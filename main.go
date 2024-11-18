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
	f.MergeCell(sheet, "A1", "B1")
	f.SetCellValue(sheet, "A1", "20 Desember 24")

	f.MergeCell(sheet, "D1", "G1")
	f.SetCellValue(sheet, "D1", "Data Penjualan Toko")

	// Tambahkan header utama
	f.MergeCell(sheet, "A2", "A4")
	f.SetCellValue(sheet, "A2", "No")
	f.MergeCell(sheet, "B2", "B4")
	f.SetCellValue(sheet, "B2", "Komoditi")
	f.MergeCell(sheet, "C2", "C4")
	f.SetCellValue(sheet, "C2", "Kemasan")
	f.MergeCell(sheet, "D2", "D4")
	f.SetCellValue(sheet, "D2", "Harga Jual (RP)")
	f.MergeCell(sheet, "E2", "G2")
	f.SetCellValue(sheet, "E2", "Dikeluarkan dari BM")
	f.MergeCell(sheet, "E3", "F3")
	f.SetCellValue(sheet, "E3", "Stok")
	f.SetCellValue(sheet, "E4", "Awal")
	f.SetCellValue(sheet, "F4", "Tambahan")
	f.MergeCell(sheet, "G3", "G4")
	f.SetCellValue(sheet, "G3", "Terjual")
	f.MergeCell(sheet, "H3", "H4")
	f.SetCellValue(sheet, "H3", "Sisa")
	f.MergeCell(sheet, "I2", "I4")
	f.SetCellValue(sheet, "I2", "Hasil Penjualan (RP)")
	f.MergeCell(sheet, "J2", "J4")
	f.SetCellValue(sheet, "J2", "Stok Akhir")

	// Tambahkan data produk
	for i, product := range products {
		row := i + 5
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
	f.SetCellValue(sheet, "C30", "Stok Keluar")
	f.SetCellValue(sheet, "B31", "Nama")
	f.SetCellValue(sheet, "C31", "Komoditi")
	f.SetCellValue(sheet, "D31", "Jumlah")

	// Tambahkan footer
	f.SetCellValue(sheet, "I30", "Total			:")
	f.SetCellValue(sheet, "I31", "Pengeluaran	:")
	f.SetCellValue(sheet, "I32", "Uang Fisik	:")
	f.SetCellValue(sheet, "I33", "Selisih		:")

    // Apply the style to cells
    styleID, _ := f.NewStyle(&excelize.Style{
        Border: []excelize.Border{
            {Type: "left", Color: "000000", Style: 1},
            {Type: "top", Color: "000000", Style: 1},
            {Type: "bottom", Color: "000000", Style: 1},
            {Type: "right", Color: "000000", Style: 1},
        },
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Font: &excelize.Font{
			Bold: true,
		},
    })

	styleFont , _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "center",
		},
		Border: []excelize.Border{
            {Type: "left", Color: "000000", Style: 1},
            {Type: "top", Color: "000000", Style: 1},
            {Type: "bottom", Color: "000000", Style: 1},
            {Type: "right", Color: "000000", Style: 1},
        },
	})

	styleCenter , _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
            {Type: "left", Color: "000000", Style: 1},
            {Type: "top", Color: "000000", Style: 1},
            {Type: "bottom", Color: "000000", Style: 1},
            {Type: "right", Color: "000000", Style: 1},
        },
	})

	styleLeft , _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "center",
		},
		Border: []excelize.Border{
            {Type: "left", Color: "000000", Style: 1},
            {Type: "top", Color: "000000", Style: 1},
            {Type: "bottom", Color: "000000", Style: 1},
            {Type: "right", Color: "000000", Style: 1},
        },
	})

	styleFontBold, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
	})

	styleBoldCenter, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold: true,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})

	f.SetColWidth(sheet, "A", "A", 4.44)
	f.SetColWidth(sheet, "B", "B", 19.11)
	f.SetColWidth(sheet, "C", "C", 16.22)
	f.SetColWidth(sheet, "D", "D", 15.56)
	f.SetColWidth(sheet, "E", "E", 4.89)
	f.SetColWidth(sheet, "F", "F", 10.22)
	f.SetColWidth(sheet, "G", "G", 6.89)
	f.SetColWidth(sheet, "H", "H", 5.78)
	f.SetColWidth(sheet, "I", "I", 16.56)
	f.SetColWidth(sheet, "J", "J", 11.67)

	f.SetCellStyle(sheet, "A2", "J4", styleID)
	f.SetCellStyle(sheet, "D1", "G1", styleBoldCenter)
	f.SetCellStyle(sheet, "B31", "D31", styleID)
	
	f.SetCellStyle(sheet, "I30", "J33", styleFont)

	f.SetCellStyle(sheet, "A5", "J29", styleCenter)

	f.SetCellStyle(sheet, "B5", "J29", styleLeft)
	f.SetCellStyle(sheet, "B32", "D34", styleLeft)

	f.SetCellStyle(sheet, "A1", "B1", styleFontBold)
	f.SetCellStyle(sheet, "A30", "D30", styleBoldCenter)

    // Simpan file ke buffer
    var buf bytes.Buffer
    if err := f.Write(&buf); err != nil {
        return nil, err
    }
    return buf.Bytes(), nil
}

