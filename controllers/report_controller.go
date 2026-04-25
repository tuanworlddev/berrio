package controllers

import (
	"archive/zip"
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"fmt"
	"net/http"
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
	DateFrom string  `form:"dateFrom" json:"dateFrom" binding:"required"`
	DateTo   string  `form:"dateTo" json:"dateTo" binding:"required"`
	Tax      float64 `form:"tax" json:"tax" binding:"required"`
	Discount float64 `form:"discount" json:"discount" binding:"required"`
}

// @Summary      Generate and download report files
// @Description  Generates Excel report files based on API key and date range, zips them, and returns the ZIP file for download
// @Tags         reports
// @Accept       json
// @Produce      application/zip
// @Param        request  body      ReportRequest  true  "Report request parameters"
// @Success      200      {file}    binary         "ZIP file containing report_total.xlsx"
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
			writeReportZip(c, item.data)
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

	data, err := buildReportZip(reports, req)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Không thể tạo file Excel"})
		return
	}

	storeReportCache(cacheKey, data)
	writeReportZip(c, data)
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

func buildReportZip(reports []models.ReportDetails, req ReportRequest) ([]byte, error) {
	report, err := services.GenerateReportExcel(reports, req.Tax, req.Discount)
	if err != nil {
		return nil, err
	}

	var zipBuffer bytes.Buffer
	zipWriter := zip.NewWriter(&zipBuffer)
	fw, err := zipWriter.Create("report_total.xlsx")
	if err != nil {
		_ = zipWriter.Close()
		return nil, err
	}
	if _, err := fw.Write(report); err != nil {
		_ = zipWriter.Close()
		return nil, err
	}
	if err := zipWriter.Close(); err != nil {
		return nil, err
	}

	return zipBuffer.Bytes(), nil
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

func writeReportZip(c *gin.Context, data []byte) {
	c.Header("Content-Type", "application/zip")
	c.Header("Content-Disposition", `attachment; filename="reports.zip"`)
	c.Data(http.StatusOK, "application/zip", data)
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
