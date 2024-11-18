package main

import (
	"bytes"
	"fmt"
	"net/http"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

type Product struct {
	ID    int
	Name  string
	Price float64
}

func main() {
	// Data dummy dalam struct
	products := []Product{
		{ID: 1, Name: "Product A", Price: 100.0},
		{ID: 2, Name: "Product B", Price: 200.0},
		{ID: 3, Name: "Product C", Price: 300.0},
	}

	// Inisialisasi Gin Router
	r := gin.Default()

	// Endpoint untuk Generate dan Download Excel
	r.GET("/download", func(c *gin.Context) {
		fileBytes, err := generateExcel(products)
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": "failed to generate Excel"})
			return
		}

		// Set header untuk mendownload file
		c.Header("Content-Description", "File Transfer")
		c.Header("Content-Disposition", "attachment; filename=products.xlsx")
		c.Data(http.StatusOK, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileBytes)
	})

	// Jalankan aplikasi di serverless
	// Untuk lingkungan seperti Vercel, gunakan handler tanpa `.Run()`.
	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		r.URL.Path = "/download"
		r.Method = "GET"
		r.Header.Set("Content-Type", "application/json")
		r.Body.Close()
	})
	r.Run(":8080") // Gunakan ini jika Anda ingin menjalankan lokal
}

func generateExcel(products []Product) ([]byte, error) {
	// Buat file Excel baru
	f := excelize.NewFile()
	sheet := "Products"

	// Tambah Sheet baru
	index,err := f.NewSheet(sheet)
	if err != nil {
		return nil, err
	}
	
	f.SetActiveSheet(index)

	// Tambahkan Header
	headers := []string{"ID", "Name", "Price"}
	for i, header := range headers {
		cell := fmt.Sprintf("%s1", string('A'+i))
		f.SetCellValue(sheet, cell, header)
	}

	// Tambahkan Data
	for i, product := range products {
		f.SetCellValue(sheet, fmt.Sprintf("A%d", i+2), product.ID)
		f.SetCellValue(sheet, fmt.Sprintf("B%d", i+2), product.Name)
		f.SetCellValue(sheet, fmt.Sprintf("C%d", i+2), product.Price)
	}

	// Simpan file ke dalam buffer memory
	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
