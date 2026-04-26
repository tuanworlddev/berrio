package main

import (
	"context"
	"fmt"
	"log"
	"net/http"
	"os"

	"github.com/gin-contrib/cors"
	"github.com/gin-gonic/gin"
	"omnituan.online/config"
	"omnituan.online/controllers"
	"omnituan.online/database"

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
	config.LoadEnv()

	ctx := context.Background()
	if err := database.Connect(ctx); err != nil {
		log.Fatalf("Cannot connect to MongoDB: %v", err)
	}
	defer func() {
		if err := database.Disconnect(ctx); err != nil {
			log.Printf("Cannot disconnect from MongoDB: %v", err)
		}
	}()

	router := gin.Default()
	router.Use(cors.Default())
	router.Static("/assets", "./public")

	serveIndex := func(c *gin.Context) {
		c.File("./public/index.html")
	}
	router.GET("/", serveIndex)
	router.HEAD("/", serveIndex)

	router.GET("/health", func(c *gin.Context) {
		c.JSON(http.StatusOK, gin.H{"message": "Welcome"})
	})

	v1 := router.Group("/api/v1")
	{
		v1.POST("/reports", controllers.HandleReportRequest)
		v1.GET("/reports/jobs", controllers.ListReportJobs)
		v1.POST("/reports/jobs", controllers.CreateReportJob)
		v1.GET("/reports/jobs/:id", controllers.GetReportJob)
		v1.GET("/reports/jobs/:id/download", controllers.DownloadReportJob)
		v1.POST("/orders", controllers.GetOrdersReport)

		v1.POST("/shops", controllers.CreateShop)
		v1.GET("/shops", controllers.GetShops)
		v1.GET("/shops/:id", controllers.GetShopByID)
		v1.PATCH("/shops/:id", controllers.UpdateShop)
		v1.DELETE("/shops/:id", controllers.DeleteShop)
		v1.GET("/shops/:id/token-status", controllers.CheckShopToken)
	}

	router.GET("/swagger/*any", ginSwagger.WrapHandler(swaggerFiles.Handler))

	port := os.Getenv("PORT")
	if port == "" {
		port = "8080"
	}

	fmt.Printf("Server started at: http://localhost:%s\n", port)
	if err := router.Run(":" + port); err != nil {
		log.Fatalf("Cannot start server: %v", err)
	}
}
