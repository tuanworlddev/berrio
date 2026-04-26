package controllers

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"fmt"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"time"

	"github.com/gin-gonic/gin"
	"omnituan.online/models"
	"omnituan.online/services"
)

const (
	reportTimeout  = 8 * time.Minute
	reportCacheTTL = 24 * time.Hour
)

var reportCache sync.Map

type cachedReport struct {
	filePath  string
	expiresAt time.Time
}

type ReportRequest struct {
	APIKey   string  `form:"apiKey" json:"apiKey" binding:"required"`
	ShopID   string  `form:"shopId" json:"shopId"`
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

	cleanupReportCache(time.Now())
	cacheKey := reportCacheKey(req)
	if cached, ok := reportCache.Load(cacheKey); ok {
		item := cached.(cachedReport)
		if time.Now().Before(item.expiresAt) && fileExists(item.filePath) {
			writeReportExcelFile(c, item.filePath, reportFileName(req))
			return
		}
		deleteReportCache(cacheKey, item)
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

	if _, err := storeReportCache(cacheKey, data); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "KhÃ´ng thá»ƒ lÆ°u file bÃ¡o cÃ¡o"})
		return
	}
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

func writeReportExcelFile(c *gin.Context, filePath string, filename string) {
	c.Header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	c.FileAttachment(filePath, filename)
}

func storeReportCache(cacheKey string, data []byte) (string, error) {
	cleanupReportCache(time.Now())

	cacheDir, err := reportCacheDir()
	if err != nil {
		return "", err
	}

	filePath := filepath.Join(cacheDir, cacheKey+".xlsx")
	if old, ok := reportCache.Load(cacheKey); ok {
		oldItem := old.(cachedReport)
		reportCache.Delete(cacheKey)
		if oldItem.filePath != "" && oldItem.filePath != filePath {
			_ = os.Remove(oldItem.filePath)
		}
	}

	if err := os.WriteFile(filePath, data, 0600); err != nil {
		return "", err
	}

	reportCache.Store(cacheKey, cachedReport{
		filePath:  filePath,
		expiresAt: time.Now().Add(reportCacheTTL),
	})
	return filePath, nil
}

func deleteReportCache(cacheKey string, item cachedReport) {
	reportCache.Delete(cacheKey)
	if item.filePath != "" {
		_ = os.Remove(item.filePath)
	}
}

func reportCacheDir() (string, error) {
	dir := filepath.Join(os.TempDir(), "berrio-report-cache")
	if err := os.MkdirAll(dir, 0700); err != nil {
		return "", err
	}
	return dir, nil
}

func fileExists(path string) bool {
	if path == "" {
		return false
	}
	info, err := os.Stat(path)
	return err == nil && !info.IsDir()
}

func cleanupReportCache(now time.Time) {
	reportCache.Range(func(key, value any) bool {
		item := value.(cachedReport)
		if now.After(item.expiresAt) || !fileExists(item.filePath) {
			deleteReportCache(key.(string), item)
		}
		return true
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
