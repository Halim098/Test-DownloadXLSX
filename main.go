package handler

import (
	"bytes"
	"fmt"
	"net/http"

	"github.com/gin-gonic/gin"
	"github.com/xuri/excelize/v2"
)

type Product struct {
	No           int
	Komoditi     string
	Kemasan      string
	HargaJual    float64
	StokAwal     int
	StokTambahan int
	Terjual      int
	Sisa         int
	Penjualan    float64
	StokAkhir    int
}

// Data Dummy
func getDummyData() []Product {
	return []Product{
		{No: 1, Komoditi: "Beras", Kemasan: "1 kg", HargaJual: 12000, StokAwal: 50, StokTambahan: 20, Terjual: 60, Sisa: 10, Penjualan: 720000, StokAkhir: 10},
		{No: 2, Komoditi: "Gula", Kemasan: "1 kg", HargaJual: 15000, StokAwal: 30, StokTambahan: 10, Terjual: 35, Sisa: 5, Penjualan: 525000, StokAkhir: 5},
		{No: 3, Komoditi: "Minyak", Kemasan: "1 liter", HargaJual: 14000, StokAwal: 40, StokTambahan: 15, Terjual: 45, Sisa: 10, Penjualan: 630000, StokAkhir: 10},
	}
}

// Fungsi utama untuk Vercel
func Handler(w http.ResponseWriter, r *http.Request) {
	// Gunakan Gin untuk routing
	gin.SetMode(gin.ReleaseMode)
	router := gin.New()

	// Endpoint untuk mendownload file
	router.GET("/download", func(c *gin.Context) {
		data := getDummyData()
		fileBytes, err := generateExcel(data)
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": "failed to generate Excel"})
			return
		}

		// Kirim file sebagai response
		c.Header("Content-Disposition", "attachment; filename=DataPenjualan.xlsx")
		c.Data(http.StatusOK, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileBytes)
	})

	// Handle request HTTP melalui Gin
	router.ServeHTTP(w, r)
}

// Fungsi untuk membuat file Excel
func generateExcel(products []Product) ([]byte, error) {
	f := excelize.NewFile()
	sheet := "Data Penjualan Toko"

	// Ganti nama sheet
	f.SetSheetName("Sheet1", sheet)

	// Tambahkan tanggal di B1
	f.SetCellValue(sheet, "B1", "20-Dec-24")

	// Tambahkan header utama
	f.MergeCell(sheet, "A2", "A3")
	f.SetCellValue(sheet, "A2", "No")
	f.MergeCell(sheet, "B2", "B3")
	f.SetCellValue(sheet, "B2", "Komoditi")
	f.MergeCell(sheet, "C2", "C3")
	f.SetCellValue(sheet, "C2", "Kemasan")
	f.MergeCell(sheet, "D2", "D3")
	f.SetCellValue(sheet, "D2", "Harga Jual (RP)")
	f.MergeCell(sheet, "E2", "F2")
	f.SetCellValue(sheet, "E2", "Stok")
	f.SetCellValue(sheet, "E3", "awal")
	f.SetCellValue(sheet, "F3", "Tambahan")
	f.MergeCell(sheet, "G2", "G3")
	f.SetCellValue(sheet, "G2", "Terjual")
	f.MergeCell(sheet, "H2", "H3")
	f.SetCellValue(sheet, "H2", "Sisa")
	f.MergeCell(sheet, "I2", "I3")
	f.SetCellValue(sheet, "I2", "Penjualan")
	f.MergeCell(sheet, "J2", "J3")
	f.SetCellValue(sheet, "J2", "Stok Akhir")

	// Tambahkan data produk
	for i, product := range products {
		row := i + 4
		f.SetCellValue(sheet, fmt.Sprintf("A%d", row), product.No)
		f.SetCellValue(sheet, fmt.Sprintf("B%d", row), product.Komoditi)
		f.SetCellValue(sheet, fmt.Sprintf("C%d", row), product.Kemasan)
		f.SetCellValue(sheet, fmt.Sprintf("D%d", row), product.HargaJual)
		f.SetCellValue(sheet, fmt.Sprintf("E%d", row), product.StokAwal)
		f.SetCellValue(sheet, fmt.Sprintf("F%d", row), product.StokTambahan)
		f.SetCellValue(sheet, fmt.Sprintf("G%d", row), product.Terjual)
		f.SetCellValue(sheet, fmt.Sprintf("H%d", row), product.Sisa)
		f.SetCellValue(sheet, fmt.Sprintf("I%d", row), product.Penjualan)
		f.SetCellValue(sheet, fmt.Sprintf("J%d", row), product.StokAkhir)
	}

	// Tambahkan header untuk "Stok Keluar"
	f.MergeCell(sheet, "A31", "B31")
	f.SetCellValue(sheet, "A31", "Nama")
	f.SetCellValue(sheet, "C31", "Komoditi")
	f.SetCellValue(sheet, "D31", "Jumlah")

	// Tambahkan footer
	f.SetCellValue(sheet, "H34", "Total")
	f.SetCellValue(sheet, "H35", "Pengeluaran")
	f.SetCellValue(sheet, "H36", "Uang Fisik")
	f.SetCellValue(sheet, "H37", "Selisih")

	// Tambahkan border untuk cell
	style, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})

	f.SetCellStyle(sheet, "A2", "J3", style)
	f.SetCellStyle(sheet, "A31", "D31", style)
	f.SetCellStyle(sheet, "H34", "I37", style)

	// Simpan file ke buffer
	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

