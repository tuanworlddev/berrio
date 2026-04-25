package controllers

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"fmt"
	"net/http"
	"net/url"
	"strings"
	"sync"
	"time"

	"github.com/gin-gonic/gin"
	"omnituan.online/models"
	"omnituan.online/services"
)

const reportTimeout = 8 * time.Minute

var reportCache sync.Map

type cachedReport struct {
	data      []byte
	expiresAt time.Time
}

type ReportRequest struct {
	APIKey   string  `form:"apiKey" json:"apiKey" binding:"required"`
	ShopName string  `form:"shopName" json:"shopName"`
	DateFrom string  `form:"dateFrom" json:"dateFrom" binding:"required"`
	DateTo   string  `form:"dateTo" json:"dateTo" binding:"required"`
	Tax      float64 `form:"tax" json:"tax" binding:"required"`
	Discount float64 `form:"discount" json:"discount" binding:"required"`
}

// @Summary      Generate and download report files
// @Description  Generates an Excel report based on API key and date range, then returns the XLSX file for download
// @Tags         reports
// @Accept       json
// @Produce      application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
// @Param        request  body      ReportRequest  true  "Report request parameters"
// @Success      200      {file}    binary         "Excel report file"
// @Failure      400      {object}  map[string]string  "Invalid request parameters or date format"
// @Failure      500      {object}  map[string]string  "Internal server error"
// @Router       /reports [post]
func HandleReportRequest(c *gin.Context) {
	req, dateFrom, dateTo, ok := parseReportRequest(c)
	if !ok {
		return
	}

	cacheKey := reportCacheKey(req)
	if cached, ok := reportCache.Load(cacheKey); ok {
		item := cached.(cachedReport)
		if time.Now().Before(item.expiresAt) {
			writeReportExcel(c, item.data, reportFileName(req))
			return
		}
		reportCache.Delete(cacheKey)
	}

	ctx, cancel := context.WithTimeout(c.Request.Context(), reportTimeout)
	defer cancel()

	reports, err := services.GetReportDetails(ctx, req.APIKey, dateFrom, dateTo)
	if err != nil {
		writeReportError(c, err)
		return
	}

	data, err := buildReportExcel(reports, req)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Không thể tạo file Excel"})
		return
	}

	storeReportCache(cacheKey, data)
	writeReportExcel(c, data, reportFileName(req))
}

func parseReportRequest(c *gin.Context) (ReportRequest, time.Time, time.Time, bool) {
	var req ReportRequest
	if err := c.ShouldBindBodyWithJSON(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return req, time.Time{}, time.Time{}, false
	}

	if req.Tax == 0 {
		req.Tax = 0.06
	}
	if req.Discount == 0 {
		req.Discount = 3.5
	}
	if req.Discount < 0 || req.Tax < 0 {
		c.JSON(http.StatusBadRequest, gin.H{"error": "tax và discount phải lớn hơn hoặc bằng 0"})
		return req, time.Time{}, time.Time{}, false
	}

	dateFrom, err := time.Parse("2006-01-02", req.DateFrom)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "dateFrom không hợp lệ. Dùng định dạng YYYY-MM-DD"})
		return req, time.Time{}, time.Time{}, false
	}
	dateTo, err := time.Parse("2006-01-02", req.DateTo)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "dateTo không hợp lệ. Dùng định dạng YYYY-MM-DD"})
		return req, time.Time{}, time.Time{}, false
	}
	if dateTo.Before(dateFrom) {
		c.JSON(http.StatusBadRequest, gin.H{"error": "dateTo phải sau hoặc bằng dateFrom"})
		return req, time.Time{}, time.Time{}, false
	}

	return req, dateFrom, dateTo, true
}

func buildReportExcel(reports []models.ReportDetails, req ReportRequest) ([]byte, error) {
	return services.GenerateReportExcel(reports, req.Tax, req.Discount)
}

func writeReportError(c *gin.Context, err error) {
	switch {
	case errors.Is(err, services.ErrReportRateLimited):
		c.JSON(http.StatusTooManyRequests, gin.H{"error": "Wildberries đang giới hạn tần suất lấy báo cáo. Vui lòng thử lại sau vài phút."})
	case errors.Is(err, context.Canceled), errors.Is(err, context.DeadlineExceeded):
		c.JSON(http.StatusGatewayTimeout, gin.H{"error": "Request lấy báo cáo quá lâu hoặc đã bị hủy. Vui lòng thử lại với khoảng ngày ngắn hơn."})
	default:
		c.JSON(http.StatusBadRequest, gin.H{"error": "Không thể lấy báo cáo", "detail": err.Error()})
	}
}

func writeReportExcel(c *gin.Context, data []byte, filename string) {
	contentType := "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	c.Header("Content-Type", contentType)
	c.Header("Content-Disposition", fmt.Sprintf(`attachment; filename="report.xlsx"; filename*=UTF-8''%s`, url.PathEscape(filename)))
	c.Data(http.StatusOK, contentType, data)
}

func storeReportCache(cacheKey string, data []byte) {
	reportCache.Store(cacheKey, cachedReport{
		data:      append([]byte(nil), data...),
		expiresAt: time.Now().Add(10 * time.Minute),
	})
}

func reportCacheKey(req ReportRequest) string {
	hash := sha256.Sum256([]byte(fmt.Sprintf("%s|%s|%s|%.4f|%.4f", req.APIKey, req.DateFrom, req.DateTo, req.Tax, req.Discount)))
	return hex.EncodeToString(hash[:])
}

func reportFileName(req ReportRequest) string {
	shopName := sanitizeFileName(req.ShopName)
	if shopName == "" {
		shopName = "shop"
	}
	return fmt.Sprintf("%s_report_%s_%s.xlsx", shopName, req.DateFrom, req.DateTo)
}

func sanitizeFileName(value string) string {
	value = strings.TrimSpace(value)
	if value == "" {
		return ""
	}

	replacer := strings.NewReplacer(
		"\\", "-",
		"/", "-",
		":", "-",
		"*", "-",
		"?", "-",
		`"`, "-",
		"<", "-",
		">", "-",
		"|", "-",
	)
	value = replacer.Replace(value)
	value = strings.Join(strings.Fields(value), " ")
	return strings.Trim(value, ". ")
}
