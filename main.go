package main

import (
	"fmt"
	"net/http"

	"github.com/gin-contrib/cors"
	"github.com/gin-gonic/gin"
	"omnituan.online/controllers"

	swaggerFiles "github.com/swaggo/files"
	ginSwagger "github.com/swaggo/gin-swagger"
	_ "omnituan.online/docs"
)

// @title API Documentation
// @version         1.0
// @description     Report Service.
// @host            localhost:8080
// @BasePath        /api/v1
func main() {
	router := gin.Default()
	router.Use(cors.Default())

	router.GET("/", func(c *gin.Context) {
		c.JSON(http.StatusOK, gin.H{"message": "Welcome"})
	})

	v1 := router.Group("/api/v1")
	{
		v1.POST("/reports", controllers.HandleReportRequest)
	}

	router.GET("/swagger/*any", ginSwagger.WrapHandler(swaggerFiles.Handler))
	fmt.Println("Server started at: http://localhost:8080")
	router.Run(":8080")
}
