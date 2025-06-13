package main

import (
	"fmt"

	"github.com/gin-contrib/cors"
	"github.com/gin-gonic/gin"
	"omnituan.online/controllers"
)

func main() {
	router := gin.Default()
	router.Use(cors.Default())
	router.GET("/v1/report", controllers.HandleReportRequest)
	fmt.Println("Server started at: http://localhost:8080")
	router.Run(":8080")
}
