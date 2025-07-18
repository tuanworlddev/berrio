package controllers

import (
	"archive/zip"
	"bytes"
	"net/http"
	"time"

	"github.com/gin-gonic/gin"
	"omnituan.online/services"
)

type ReportRequest struct {
	APIKey   string  `form:"apiKey" binding:"required"`
	DateFrom string  `form:"dateFrom" binding:"required"`
	DateTo   string  `form:"dateTo" binding:"required"`
	Tax      float64 `form:"tax" binding:"required"`
	Discount float64 `form:"discount" binding:"required"`
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

	reports, err := services.GetReportDetails(req.APIKey, dateFrom, dateTo)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Cannot get reports"})
		return
	}

	//report1, err1 := services.GenerateDetailedExcel(reports)
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
	c.Header("Content-Type", "application/zip")
	c.Header("Content-Disposition", `attachment; filename="reports.zip"`)
	c.Data(http.StatusOK, "application/zip", zipBuffer.Bytes())
}
