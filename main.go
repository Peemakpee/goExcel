package main

import (
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strconv"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

func main() {
	// Keep Gin in debug mode for testing
	gin.SetMode(gin.DebugMode)

	// Create a new Gin engine with default middleware
	r := gin.Default()

	// Create temp directory if it doesn't exist
	if err := os.MkdirAll("temp", os.ModePerm); err != nil {
		log.Fatalf("Failed to create temp directory: %v", err)
	}

	r.POST("/upload", handleExcelUpload)

	log.Println("Server is running in debug mode on http://localhost:8080")
	log.Println("Use a tool like Postman or curl to send a POST request with an Excel file to /upload")
	if err := r.Run(":8080"); err != nil {
		log.Fatalf("Failed to start server: %v", err)
	}
}

func handleExcelUpload(c *gin.Context) {
	log.Println("Received file upload request")

	// Get the file from the request
	file, err := c.FormFile("file")
	if err != nil {
		log.Printf("Error: No file uploaded - %v", err)
		c.JSON(http.StatusBadRequest, gin.H{"error": "No file uploaded"})
		return
	}

	log.Printf("Received file: %s", file.Filename)

	// Validate file extension
	if filepath.Ext(file.Filename) != ".xlsx" {
		log.Printf("Error: Invalid file format - %s", file.Filename)
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid file format. Please upload an Excel file (.xlsx)"})
		return
	}

	// Save the uploaded file temporarily
	tempFilePath := filepath.Join("temp", file.Filename)
	if err := c.SaveUploadedFile(file, tempFilePath); err != nil {
		log.Printf("Error: Failed to save file - %v", err)
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to save file"})
		return
	}
	log.Printf("File saved temporarily as: %s", tempFilePath)
	defer os.Remove(tempFilePath) // Clean up the file after processing

	// Open the Excel file
	f, err := excelize.OpenFile(tempFilePath)
	if err != nil {
		log.Printf("Error: Failed to open Excel file - %v", err)
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to open Excel file"})
		return
	}
	defer f.Close()

	// Get all sheet names
	sheets := f.GetSheetList()
	log.Printf("Found %d sheets in the Excel file", len(sheets))

	// Process each sheet
	var result []map[string]interface{}
	for _, sheet := range sheets {
		log.Printf("Processing sheet: %s", sheet)
		sheetData, err := processSheet(f, sheet)
		if err != nil {
			log.Printf("Error processing sheet %s: %v", sheet, err)
			c.JSON(http.StatusInternalServerError, gin.H{"error": fmt.Sprintf("Error processing sheet %s: %v", sheet, err)})
			return
		}
		result = append(result, sheetData)
	}

	log.Println("File processed successfully")
	c.JSON(http.StatusOK, gin.H{"message": "File processed successfully", "data": result})
}

func processSheet(f *excelize.File, sheetName string) (map[string]interface{}, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows from sheet %s: %v", sheetName, err)
	}

	if len(rows) < 2 {
		return nil, fmt.Errorf("sheet %s is empty or has no data rows", sheetName)
	}

	log.Printf("Sheet %s has %d rows", sheetName, len(rows))

	headers := rows[0]
	data := make([]map[string]interface{}, 0)

	for i, row := range rows[1:] {
		rowData := make(map[string]interface{})
		for j, cell := range row {
			if j < len(headers) {
				// Try to parse the cell as a number
				if val, err := strconv.ParseFloat(cell, 64); err == nil {
					rowData[headers[j]] = val
				} else {
					rowData[headers[j]] = cell
				}
			}
		}
		data = append(data, rowData)
		if i < 5 {
			log.Printf("Processed row %d: %v", i+1, rowData)
		} else if i == 5 {
			log.Println("...")
		}
	}

	log.Printf("Finished processing sheet %s", sheetName)

	return map[string]interface{}{
		"sheetName": sheetName,
		"headers":   headers,
		"data":      data,
	}, nil
}
