package controllers

import (
	"context"
	"net/http"
	"strconv"
	"time"

	"github.com/gin-gonic/gin"
	"omnituan.online/models"
	"omnituan.online/services"
)

const requestTimeout = 10 * time.Second

// @Summary      Create shop
// @Description  Creates a shop
// @Tags         shops
// @Accept       json
// @Produce      json
// @Param        request  body      models.CreateShopRequest  true  "Shop payload"
// @Success      201      {object}  models.Shop
// @Failure      400      {object}  map[string]string
// @Failure      500      {object}  map[string]string
// @Router       /shops [post]
func CreateShop(c *gin.Context) {
	var req models.CreateShopRequest
	if err := c.ShouldBindJSON(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	ctx, cancel := requestContext(c)
	defer cancel()

	shop, err := services.CreateShop(ctx, req)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Cannot create shop"})
		return
	}

	c.JSON(http.StatusCreated, shop)
}

// @Summary      Get shops
// @Description  Gets shops with pagination
// @Tags         shops
// @Produce      json
// @Param        page   query     int  false  "Page number"
// @Param        limit  query     int  false  "Page size"
// @Success      200    {object}  models.PaginationResponse[models.Shop]
// @Failure      500    {object}  map[string]string
// @Router       /shops [get]
func GetShops(c *gin.Context) {
	page := parseInt64Query(c, "page", 1)
	limit := parseInt64Query(c, "limit", 10)

	ctx, cancel := requestContext(c)
	defer cancel()

	shops, err := services.GetShops(ctx, page, limit)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Cannot get shops"})
		return
	}

	c.JSON(http.StatusOK, shops)
}

// @Summary      Get shop by id
// @Description  Gets one shop by id
// @Tags         shops
// @Produce      json
// @Param        id   path      string  true  "Shop ID"
// @Success      200  {object}  models.Shop
// @Failure      400  {object}  map[string]string
// @Failure      404  {object}  map[string]string
// @Router       /shops/{id} [get]
func GetShopByID(c *gin.Context) {
	ctx, cancel := requestContext(c)
	defer cancel()

	shop, err := services.GetShopByID(ctx, c.Param("id"))
	if err != nil {
		writeShopError(c, err)
		return
	}

	c.JSON(http.StatusOK, shop)
}

// @Summary      Update shop
// @Description  Updates a shop by id
// @Tags         shops
// @Accept       json
// @Produce      json
// @Param        id       path      string                    true  "Shop ID"
// @Param        request  body      models.UpdateShopRequest  true  "Shop payload"
// @Success      200      {object}  models.Shop
// @Failure      400      {object}  map[string]string
// @Failure      404      {object}  map[string]string
// @Failure      500      {object}  map[string]string
// @Router       /shops/{id} [patch]
func UpdateShop(c *gin.Context) {
	var req models.UpdateShopRequest
	if err := c.ShouldBindJSON(&req); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": err.Error()})
		return
	}

	ctx, cancel := requestContext(c)
	defer cancel()

	shop, err := services.UpdateShop(ctx, c.Param("id"), req)
	if err != nil {
		writeShopError(c, err)
		return
	}

	c.JSON(http.StatusOK, shop)
}

// @Summary      Delete shop
// @Description  Deletes a shop by id
// @Tags         shops
// @Param        id   path  string  true  "Shop ID"
// @Success      204
// @Failure      400  {object}  map[string]string
// @Failure      404  {object}  map[string]string
// @Failure      500  {object}  map[string]string
// @Router       /shops/{id} [delete]
func DeleteShop(c *gin.Context) {
	ctx, cancel := requestContext(c)
	defer cancel()

	if err := services.DeleteShop(ctx, c.Param("id")); err != nil {
		writeShopError(c, err)
		return
	}

	c.Status(http.StatusNoContent)
}

func requestContext(c *gin.Context) (context.Context, context.CancelFunc) {
	return context.WithTimeout(c.Request.Context(), requestTimeout)
}

func parseInt64Query(c *gin.Context, key string, fallback int64) int64 {
	value, err := strconv.ParseInt(c.DefaultQuery(key, strconv.FormatInt(fallback, 10)), 10, 64)
	if err != nil {
		return fallback
	}

	return value
}

func writeShopError(c *gin.Context, err error) {
	if services.IsNotFound(err) {
		c.JSON(http.StatusNotFound, gin.H{"error": "Shop not found"})
		return
	}

	c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid shop id"})
}
