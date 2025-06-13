package services

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"time"

	"github.com/xuri/excelize/v2"
	"omnituan.online/models"
)

func GetReportDetails(apiKey string, dateFrom, dateTo time.Time) ([]models.ReportDetails, error) {
	var allReports []models.ReportDetails
	for from := dateFrom; from.Before(dateTo); {
		to := from.AddDate(0, 0, 6)
		if to.After(dateTo) {
			to = dateTo
		}
		url := fmt.Sprintf(
			"https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod?dateFrom=%s&dateTo=%s",
			from.Format(time.RFC3339),
			to.Format(time.RFC3339),
		)
		client := &http.Client{
			Timeout: 10 * time.Second,
		}

		req, err := http.NewRequest("GET", url, nil)
		if err != nil {
			fmt.Println("Error creating request:", err)
			return nil, err
		}
		req.Header.Set("Authorization", fmt.Sprintf("Bearer %s", apiKey))
		req.Header.Set("Content-Type", "application/json")

		res, err := client.Do(req)
		if err != nil {
			fmt.Println("Error making request:", err)
			return nil, err
		}
		defer res.Body.Close()

		if res.StatusCode != http.StatusOK {
			fmt.Printf("Error: Status code %d\n", res.StatusCode)
			body, _ := io.ReadAll(res.Body)
			fmt.Println("Response:", string(body))
			return nil, err
		}

		body, err := io.ReadAll(res.Body)
		if err != nil {
			fmt.Println("Error reading response:", err)
			return nil, err
		}

		var reports []models.ReportDetails
		if err := json.Unmarshal(body, &reports); err != nil {
			fmt.Println("Error decoding JSON:", err)
			return nil, err
		}
		allReports = append(allReports, reports...)

		fmt.Printf("Count: %d\n", len(reports))
		from = to.AddDate(0, 0, 1)
	}
	return allReports, nil
}

func GenerateDetailedExcel(reports []models.ReportDetails) ([]byte, error) {
	f := excelize.NewFile()
	sheet := "Report"
	f.NewSheet(sheet)
	f.DeleteSheet("Sheet1")

	// Định nghĩa tiêu đề
	headers := []string{
		"№", "Номер поставки", "Предмет", "Код номенклатуры", "Бренд", "Артикул поставщика", "Название", "Размер",
		"Баркод", "Тип документа", "Дата заказа покупателем", "Дата продажи", "Кол-во", "Цена розничная",
		"Согласованный продуктовый дисконт, %", "Промокод %", "Итоговая согласованная скидка, %",
		"Цена розничная с учетом согласованной скидки", "Стикер МП", "Номер сборочного задания", "Код маркировки", "ШК", "Srid",
	}

	// Ghi tiêu đề
	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheet, cell, h)
	}

	// Ghi dữ liệu
	for i, r := range reports {
		row := i + 2 // Dòng bắt đầu từ 2
		data := []any{
			i + 1,
			r.RealizationReportID,
			r.SubjectName,
			r.NmID,
			r.BrandName,
			r.SaName,
			r.TsName,
			"", // Размер: bạn cần map nếu có trong struct
			r.Barcode,
			r.DocTypeName,
			r.OrderDt.Format("2006-01-02"),
			r.SaleDt.Format("2006-01-02"),
			r.Quantity,
			r.RetailPrice,
			r.ProductDiscountForReport,
			r.SupplierPromo,
			r.SalePercent,
			r.RetailPriceWithDiscRub,
			r.StickerID,
			r.AssemblyID,
			r.Kiz,
			r.ShkID,
			r.Srid,
		}

		for j, val := range data {
			cell, _ := excelize.CoordinatesToCellName(j+1, row)
			f.SetCellValue(sheet, cell, val)
		}
	}

	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func GenerateReportExcel(reports []models.ReportDetails) ([]byte, error) {
	f := excelize.NewFile()
	sheet := "Report"
	f.SetSheetName("Sheet1", sheet)

	// Định dạng kiểu cho tên bảng (background nhạt, chữ trắng)
	headerStyleLight, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"33CC33"}, Pattern: 1}, // Xanh lá nhạt
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})

	// Định dạng kiểu cho tiêu đề cột (background đậm, chữ trắng)
	titleStyleDark, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"33CC33"}, Pattern: 1}, // Xanh lá đậm
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})

	// Bảng Doanh thu (A1:C3+)
	f.SetCellValue(sheet, "A1", "BẢNG DOANH THU")
	f.MergeCell(sheet, "A1", "C1")
	f.SetCellStyle(sheet, "A1", "C1", headerStyleLight)
	f.SetCellValue(sheet, "A2", "Артикул поставщика")
	f.SetCellValue(sheet, "B2", "Giá đăng bán")
	f.SetCellValue(sheet, "C2", "Tiền chuyển cho hàng hóa đã bán chưa bao gồm chi phí logistic và chi phí khác")
	f.SetCellStyle(sheet, "A2", "C2", titleStyleDark)
	row := 3
	for _, r := range reports {
		if r.SupplierOperName == "Продажа" {
			f.SetCellValue(sheet, fmt.Sprintf("A%d", row), r.SaName)
			f.SetCellValue(sheet, fmt.Sprintf("B%d", row), r.RetailPrice)
			f.SetCellValue(sheet, fmt.Sprintf("C%d", row), r.PpvzForPay)
			row++
		}
	}

	// Bảng Hàng mua bị trả lại (F1:H3+)
	headerStyleLight, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FF6600"}, Pattern: 1}, // Cam nhạt
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	titleStyleDark, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FF6600"}, Pattern: 1}, // Cam đậm
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	f.SetCellValue(sheet, "F1", "BẢNG HÀNG MUA BỊ TRẢ LẠI")
	f.MergeCell(sheet, "F1", "H1")
	f.SetCellStyle(sheet, "F1", "H1", headerStyleLight)
	f.SetCellValue(sheet, "F2", "Артикул поставщика")
	f.SetCellValue(sheet, "G2", "Giá gốc đăng bán")
	f.SetCellValue(sheet, "H2", "Giá trả lại")
	f.SetCellStyle(sheet, "F2", "H2", titleStyleDark)
	row = 3
	for _, r := range reports {
		if r.ReturnAmount > 0 {
			f.SetCellValue(sheet, fmt.Sprintf("F%d", row), r.SaName)
			f.SetCellValue(sheet, fmt.Sprintf("G%d", row), r.RetailPrice)
			f.SetCellValue(sheet, fmt.Sprintf("H%d", row), r.RetailPriceWithDiscRub)
			row++
		}
	}

	// Bảng Phí logistic (K1:L3+)
	headerStyleLight, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"0066CC"}, Pattern: 1}, // Xanh dương nhạt
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	titleStyleDark, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"0066CC"}, Pattern: 1}, // Xanh dương đậm
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	f.SetCellValue(sheet, "K1", "BẢNG PHÍ LOGISTIC")
	f.MergeCell(sheet, "K1", "L1")
	f.SetCellStyle(sheet, "K1", "L1", headerStyleLight)
	f.SetCellValue(sheet, "K2", "Артикул поставщика")
	f.SetCellValue(sheet, "L2", "Chi phí logistic")
	f.SetCellStyle(sheet, "K2", "L2", titleStyleDark)
	row = 3
	for _, r := range reports {
		if r.RebillLogisticCost > 0 {
			f.SetCellValue(sheet, fmt.Sprintf("K%d", row), r.SaName)
			f.SetCellValue(sheet, fmt.Sprintf("L%d", row), r.RebillLogisticCost)
			row++
		}
	}

	// Bảng Phí đơn hàng bị hủy/không mua (O1:P3+)
	headerStyleLight, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"CC0000"}, Pattern: 1}, // Đỏ nhạt
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	titleStyleDark, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"CC0000"}, Pattern: 1}, // Đỏ đậm
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	f.SetCellValue(sheet, "O1", "BẢNG PHÍ ĐƠN HÀNG BỊ HỦY OR KHÔNG MUA")
	f.MergeCell(sheet, "O1", "P1")
	f.SetCellStyle(sheet, "O1", "P1", headerStyleLight)
	f.SetCellValue(sheet, "O2", "Артикул поставщика")
	f.SetCellValue(sheet, "P2", "phí vận chuyển hàng trả lại")
	f.SetCellStyle(sheet, "O2", "P2", titleStyleDark)
	row = 3
	for _, r := range reports {
		if r.Penalty > 0 && r.BonusTypeName != "" {
			f.SetCellValue(sheet, fmt.Sprintf("O%d", row), r.SaName)
			f.SetCellValue(sheet, fmt.Sprintf("P%d", row), r.Penalty)
			row++
		}
	}

	// Bảng Chi phí khác (S1:T7)
	headerStyleLight, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"6600CC"}, Pattern: 1}, // Tím nhạt
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	titleStyleDark, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"6600CC"}, Pattern: 1}, // Tím đậm
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	f.SetCellValue(sheet, "S1", "BẢNG CHI PHÍ KHÁC")
	f.MergeCell(sheet, "S1", "T1")
	f.SetCellStyle(sheet, "S1", "T1", headerStyleLight)
	f.SetCellValue(sheet, "S2", "Chi phí khác")
	f.SetCellValue(sheet, "T2", "Số tiền")
	f.SetCellStyle(sheet, "S2", "T2", titleStyleDark)
	f.SetCellValue(sheet, "S3", "Tiền phạt")
	f.SetCellValue(sheet, "S4", "Chi phí lưu trữ")
	f.SetCellValue(sheet, "S5", "Chi phí quảng cáo")
	f.SetCellValue(sheet, "S6", "Chi phí chấp nhận")
	f.SetCellValue(sheet, "S7", "Tổng")
	var totalOtherCost float64
	for _, r := range reports {
		totalOtherCost += r.Penalty + r.StorageFee + r.SupplierPromo + r.Acceptance
	}
	f.SetCellValue(sheet, "T3", totalOtherCost) // Tiền phạt tổng
	f.SetCellValue(sheet, "T4", totalOtherCost) // Chi phí lưu trữ tổng
	f.SetCellValue(sheet, "T5", totalOtherCost) // Chi phí quảng cáo tổng
	f.SetCellValue(sheet, "T6", totalOtherCost) // Chi phí chấp nhận tổng
	f.SetCellValue(sheet, "T7", totalOtherCost) // Tổng chi phí khác

	// Bảng Tổng kết (W1:AG2)
	headerStyleLight, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFCC00"}, Pattern: 1}, // Vàng nhạt
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	titleStyleDark, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "FFFFFF"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFCC00"}, Pattern: 1}, // Vàng đậm
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
		},
	})
	f.SetCellValue(sheet, "W1", "BẢNG TỔNG KẾT")
	f.MergeCell(sheet, "W1", "AG1")
	f.SetCellStyle(sheet, "W1", "AG1", headerStyleLight)
	f.SetCellValue(sheet, "W2", "Doanh thu gộp")
	f.SetCellValue(sheet, "X2", "Doanh thu thuần")
	f.SetCellValue(sheet, "Y2", "Giảm trừ doanh thu")
	f.SetCellValue(sheet, "Z2", "Chi phí logistic")
	f.SetCellValue(sheet, "AA2", "Chi phí khác")
	f.SetCellValue(sheet, "AB2", "Doanh thu chưa trừ giá vốn")
	f.SetCellValue(sheet, "AC2", "Giá vốn ước lượng")
	f.SetCellValue(sheet, "AD2", "Doanh thu giảm trừ thuế")
	f.SetCellValue(sheet, "AE2", "Lãi gộp")
	f.SetCellStyle(sheet, "W2", "AE2", titleStyleDark)

	var grossRevenue, netRevenue, returnDeduction, logisticCost, otherCost, costOfGoods float64
	for _, r := range reports {
		if r.SupplierOperName == "Продажа" {
			grossRevenue += r.RetailPrice * float64(r.Quantity)
			netRevenue += r.PpvzForPay
		}
		if r.ReturnAmount > 0 {
			returnDeduction += r.RetailPriceWithDiscRub * float64(r.ReturnAmount)
		}
		logisticCost += r.RebillLogisticCost
		otherCost += r.Penalty + r.StorageFee + r.SupplierPromo + r.Acceptance
		costOfGoods += r.RetailPrice * float64(r.Quantity) * 0.7 // Giả sử giá vốn 70%
	}
	revenueBeforeCOGS := netRevenue - logisticCost - otherCost
	taxDeduction := revenueBeforeCOGS * 0.2 // Giả sử thuế 20%
	grossProfit := revenueBeforeCOGS - costOfGoods

	f.SetCellValue(sheet, "W3", grossRevenue)
	f.SetCellValue(sheet, "X3", netRevenue)
	f.SetCellValue(sheet, "Y3", returnDeduction)
	f.SetCellValue(sheet, "Z3", logisticCost)
	f.SetCellValue(sheet, "AA3", otherCost)
	f.SetCellValue(sheet, "AB3", revenueBeforeCOGS)
	f.SetCellValue(sheet, "AC3", costOfGoods)
	f.SetCellValue(sheet, "AD3", revenueBeforeCOGS-taxDeduction)
	f.SetCellValue(sheet, "AE3", grossProfit)

	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
