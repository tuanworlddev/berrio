package controllers

import (
	"context"
	"errors"
	"net/http"
	"time"

	"github.com/gin-gonic/gin"
	"omnituan.online/services"
)

type AnalyticOrderRequest struct {
	APIKey   string `form:"apiKey" json:"apiKey" binding:"required"`
	DateFrom string `form:"dateFrom" json:"dateFrom" binding:"required"`
	DateTo   string `form:"dateTo" json:"dateTo" binding:"required"`
}

// @Summary      Generates reports orders
// @Description  Generates reports orders
// @Tags         orders
// @Accept       json
// @Produce      application/json
// @Param        request  body      AnalyticOrderRequest  true  "Report request parameters"
// @Success      200      {object}  []services.ChartData
// @Failure      400      {object}  map[string]string  "Invalid request parameters or date format"
// @Failure      500      {object}  map[string]string  "Internal server error"
// @Router       /orders [post]
func GetOrdersReport(c *gin.Context) {
	var req AnalyticOrderRequest

	if err := c.ShouldBindBodyWithJSON(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid apiKey, dateTo, dateFrom"})
		return
	}

	begin, end, ok := parseAnalyticOrderRange(c, req)
	if !ok {
		return
	}

	ctx, cancel := context.WithTimeout(c.Request.Context(), 3*time.Minute)
	defer cancel()

	data, err := services.GetOrders(ctx, req.APIKey, begin, end)
	if err != nil {
		writeAnalyticOrderError(c, err)
		return
	}

	c.JSON(http.StatusOK, data)
}

func parseAnalyticOrderRange(c *gin.Context, req AnalyticOrderRequest) (string, string, bool) {
	dateFrom, err := time.Parse("2006-01-02", req.DateFrom)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "dateFrom không hợp lệ. Dùng định dạng YYYY-MM-DD"})
		return "", "", false
	}

	dateTo, err := time.Parse("2006-01-02", req.DateTo)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "dateTo không hợp lệ. Dùng định dạng YYYY-MM-DD"})
		return "", "", false
	}
	if dateTo.Before(dateFrom) {
		c.JSON(http.StatusBadRequest, gin.H{"error": "dateTo phải sau hoặc bằng dateFrom"})
		return "", "", false
	}

	return dateFrom.Format("2006-01-02"), dateTo.Format("2006-01-02"), true
}

func writeAnalyticOrderError(c *gin.Context, err error) {
	switch {
	case errors.Is(err, services.ErrOrdersUnauthorized):
		c.JSON(http.StatusUnauthorized, gin.H{"error": "Token Analytics không hợp lệ hoặc đã hết hạn"})
	case errors.Is(err, services.ErrOrdersForbidden):
		c.JSON(http.StatusForbidden, gin.H{"error": "Token không có quyền Analytics"})
	case errors.Is(err, services.ErrOrdersRateLimited):
		c.JSON(http.StatusTooManyRequests, gin.H{"error": "Wildberries đang giới hạn tần suất lấy đơn hàng. Vui lòng thử lại sau."})
	case errors.Is(err, context.Canceled), errors.Is(err, context.DeadlineExceeded):
		c.JSON(http.StatusGatewayTimeout, gin.H{"error": "Lấy dữ liệu đơn hàng quá lâu. Vui lòng thử khoảng ngày ngắn hơn."})
	default:
		c.JSON(http.StatusBadRequest, gin.H{"error": "Không thể lấy dữ liệu đơn hàng", "detail": err.Error()})
	}
}
