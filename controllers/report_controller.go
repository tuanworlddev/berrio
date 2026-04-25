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
// @Description  Generates two Excel report files based on API key and date range, zips them, and returns the ZIP file for download
// @Tags         reports
// @Accept       json
// @Produce      application/zip
// @Param        request  body      ReportRequest  true  "Report request parameters"
// @Success      200      {file}    binary         "ZIP file containing report1.xlsx and report2.xlsx"
// @Failure      400      {object}  map[string]string  "Invalid request parameters or date format"
// @Failure      500      {object}  map[string]string  "Internal server error"
// @Router       /reports [post]
func HandleReportRequest(c *gin.Context) {
	var req ReportRequest

	if err := c.ShouldBindBodyWithJSON(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	if req.Tax == 0 {
		req.Tax = 0.06
	}
	if req.Discount == 0 {
		req.Discount = 3.5
	}
	if req.Discount < 0 || req.Tax < 0 {
		c.JSON(http.StatusBadRequest, gin.H{"error": "tax and discount must be greater than or equal to 0"})
		return
	}

	dateFrom, err := time.Parse("2006-01-02", req.DateFrom)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid dateFrom format. Use YYYY-MM-DD"})
		return
	}
	dateTo, err := time.Parse("2006-01-02", req.DateTo)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid dateTo format. Use YYYY-MM-DD"})
		return
	}
	if dateTo.Before(dateFrom) {
		c.JSON(http.StatusBadRequest, gin.H{"error": "dateTo must be after dateFrom"})
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
		switch {
		case errors.Is(err, services.ErrReportRateLimited):
			c.JSON(http.StatusTooManyRequests, gin.H{"error": "Wildberries đang giới hạn tần suất lấy báo cáo. Vui lòng thử lại sau vài phút."})
		case errors.Is(err, context.Canceled), errors.Is(err, context.DeadlineExceeded):
			c.JSON(http.StatusGatewayTimeout, gin.H{"error": "Request lấy báo cáo quá lâu hoặc đã bị hủy. Vui lòng thử lại với khoảng ngày ngắn hơn."})
		default:
			c.JSON(http.StatusBadRequest, gin.H{"error": "Cannot get reports", "detail": err.Error()})
		}
		return
	}
	// fmt.Println("Excel 1")
	// report1, err1 := services.GenerateDetailedExcel(reports)
	fmt.Println("Excel 2")
	report2, err2 := services.GenerateReportExcel(reports, req.Tax, req.Discount)

	// if err1 != nil {
	// 	c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to generate Excel files"})
	// 	return
	// }

	if err2 != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to generate Excel files"})
		return
	}

	var zipBuffer bytes.Buffer
	zipWriter := zip.NewWriter(&zipBuffer)

	// fw1, err := zipWriter.Create("report_vi.xlsx")
	// if err != nil {
	// 	c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create zip entry 1"})
	// 	return
	// }
	// if _, err := fw1.Write(report1); err != nil {
	// 	c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to write file 1 to zip"})
	// 	return
	// }

	fw2, err := zipWriter.Create("report_total.xlsx")
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create zip entry 1"})
		return
	}
	if _, err := fw2.Write(report2); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to write file 2 to zip"})
		return
	}

	if err := zipWriter.Close(); err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to close zip"})
		return
	}

	reportCache.Store(cacheKey, cachedReport{
		data:      append([]byte(nil), zipBuffer.Bytes()...),
		expiresAt: time.Now().Add(10 * time.Minute),
	})
	writeReportZip(c, zipBuffer.Bytes())
}

func writeReportZip(c *gin.Context, data []byte) {
	c.Header("Content-Type", "application/zip")
	c.Header("Content-Disposition", `attachment; filename="reports.zip"`)
	c.Data(http.StatusOK, "application/zip", data)
}

func reportCacheKey(req ReportRequest) string {
	hash := sha256.Sum256([]byte(fmt.Sprintf("%s|%s|%s|%.4f|%.4f", req.APIKey, req.DateFrom, req.DateTo, req.Tax, req.Discount)))
	return hex.EncodeToString(hash[:])
}
