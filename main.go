package main

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

    // Tambahkan sheet baru
    f.SetSheetName("Sheet1", sheet)

    // Tambahkan tanggal di B1
    f.SetCellValue(sheet, "B1", "20-Dec-24")

    // Merge header kolom
    headerRanges := map[string]string{
        "B2:B3": "Komoditi",
        "C2:C3": "Kemasan",
        "D2:D3": "Harga Jual (RP)",
        "E2:E3": "Stok awal",
        "F2:F3": "Stok Tambahan",
        "G2:G3": "Terjual",
        "H2:H3": "Sisa",
        "I2:I3": "Penjualan",
        "J2:J3": "Stok Akhir",
    }
    for cells, title := range headerRanges {
        f.MergeCell(sheet, cells[:2], cells[3:])
        f.SetCellValue(sheet, cells[:2], title)
    }
    f.SetCellValue(sheet, "A2", "No")

    // Tambahkan data
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

    // Tambahkan tabel "Stok Keluar"
    f.MergeCell(sheet, "A31", "B31")
    f.SetCellValue(sheet, "A31", "Nama")
    f.SetCellValue(sheet, "C31", "Komoditi")
    f.SetCellValue(sheet, "D31", "Jumlah")

    // Tambahkan footer
    f.SetCellValue(sheet, "H34", "Total")
    f.SetCellValue(sheet, "H35", "Pengeluaran")
    f.SetCellValue(sheet, "H36", "Uang Fisik")
    f.SetCellValue(sheet, "H37", "Selisih")

    // Buat border style
    style := excelize.Style{
        Border: []excelize.Border{
            {Type: "left", Color: "000000", Style: 1},
            {Type: "top", Color: "000000", Style: 1},
            {Type: "bottom", Color: "000000", Style: 1},
            {Type: "right", Color: "000000", Style: 1},
        },
    }

    // Apply the style to cells
    styleID, err := f.NewStyle(&style)
    if err != nil {
        return nil, err
    }

    f.SetCellStyle(sheet, "A2", "J3", styleID)
    f.SetCellStyle(sheet, "A31", "D31", styleID)
    f.SetCellStyle(sheet, "H34", "I37", styleID)

    // Simpan file ke buffer
    var buf bytes.Buffer
    if err := f.Write(&buf); err != nil {
        return nil, err
    }
    return buf.Bytes(), nil
}
