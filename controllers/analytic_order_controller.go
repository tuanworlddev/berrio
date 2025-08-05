package controllers

import (
	"net/http"

	"github.com/gin-gonic/gin"
	"omnituan.online/services"
)

type AnalyticOrderRequest struct {
	APIKey   string `form:"apiKey" binding:"required"`
	DateFrom string `form:"dateFrom" binding:"required"`
	DateTo   string `form:"dateTo" binding:"required"`
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

	data, err := services.GetOrders(req.APIKey, req.DateFrom, req.DateTo)
	if err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Error get orders reports"})
	}

	c.JSON(http.StatusOK, data)
}
