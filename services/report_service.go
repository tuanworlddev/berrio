package services

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"math"
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
		client := &http.Client{}

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

		fmt.Printf("Date from: %v, Date to: %v, Count: %d\n", from.Format("02-01-2006"), to.Format("02-01-2006"), len(reports))
		from = to.AddDate(0, 0, 1)
		time.Sleep(25 * time.Second)
	}
	return allReports, nil
}

func GenerateDetailedExcel(reports []models.ReportDetails) ([]byte, error) {
	f := excelize.NewFile()
	sheet1 := "Sheet1"
	sw, err := f.NewStreamWriter(sheet1)
	if err != nil {
		return nil, err
	}

	// Định nghĩa tiêu đề
	headers := []any{
		"№",
		"Номер поставки",
		"Предмет",
		"Код номенклатуры",
		"Бренд",
		"Артикул поставщика",
		"Название",
		"Размер",
		"Баркод",
		"Тип документа",
		"Обоснование для оплаты",
		"Дата заказа покупателем",
		"Дата продажи",
		"Кол-во",
		"Цена розничная",
		"Вайлдберриз реализовал Товар (Пр)",
		"Согласованный продуктовый дисконт, %",
		"Промокод %",
		"Итоговая согласованная скидка, %",
		"Цена розничная с учетом согласованной скидки",
		"Размер снижения кВВ из-за рейтинга, %",
		"Размер изменения кВВ из-за акции, %",
		"Скидка постоянного Покупателя (СПП), %",
		"Размер кВВ, %",
		"Размер  кВВ без НДС, % Базовый",
		"Итоговый кВВ без НДС, %",
		"Вознаграждение с продаж до вычета услуг поверенного, без НДС",
		"Возмещение за выдачу и возврат товаров на ПВЗ",
		"Эквайринг/Комиссии за организацию платежей",
		"Размер комиссии за эквайринг/Комиссии за организацию платежей, %",
		"Тип платежа за Эквайринг/Комиссии за организацию платежей",
		"Вознаграждение Вайлдберриз (ВВ), без НДС",
		"НДС с Вознаграждения Вайлдберриз",
		"К перечислению Продавцу за реализованный Товар",
		"Количество доставок",
		"Количество возврата",
		"Услуги по доставке товара покупателю",
		"Дата начала действия фиксации",
		"Дата конца действия фиксации",
		"Признак услуги платной доставки",
		"Общая сумма штрафов",
		"Корректировка Вознаграждения Вайлдберриз (ВВ)",
		"Виды логистики, штрафов и корректировок ВВ",
		"Стикер МП",
		"Наименование банка-эквайера",
		"Номер офиса",
		"Наименование офиса доставки",
		"ИНН партнера",
		"Партнер",
		"Склад",
		"Страна",
		"Тип коробов",
		"Номер таможенной декларации",
		"Номер сборочного задания",
		"Код маркировки",
		"ШК",
		"Srid",
		"Возмещение издержек по перевозке/по складским операциям с товаром",
		"Организатор перевозки",
		"Хранение",
		"Удержания",
		"Платная приемка",
		"Фиксированный коэффициент склада по поставке",
		"Признак продажи юридическому лицу",
		"Номер короба для платной приемки",
		"Скидка по программе софинансирования",
		"Скидка Wibes, %",
	}

	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "center", WrapText: true},
	})
	f.SetRowHeight(sheet1, 1, 30)

	// Ghi tiêu đề
	// for i, h := range headers {
	// 	cell, _ := excelize.CoordinatesToCellName(i+1, 1)
	// 	f.SetCellValue(sheet, cell, h)
	// 	f.SetCellStyle(sheet, cell, cell, headerStyle)
	// }
	if err := sw.SetRow("A1", headers, excelize.RowOpts{StyleID: headerStyle, Height: 24}); err != nil {
		return nil, err
	}

	// Ghi dữ liệu
	for i, r := range reports {
		row := i + 2
		data := []any{
			i + 1,                          // №
			r.GiID,                         // Номер поставки
			r.SubjectName,                  // Предмет
			r.NmID,                         // Код номенклатуры
			r.BrandName,                    // Бренд
			r.SaName,                       // Артикул поставщика
			"",                             // Название
			r.TsName,                       // Размер
			r.Barcode,                      // Баркод
			r.DocTypeName,                  // Тип документа
			r.SupplierOperName,             //Обоснование для оплаты
			r.OrderDt.Format("2006-01-02"), // Дата заказа покупателем
			r.SaleDt.Format("2006-01-02"),  // Дата продажи
			r.Quantity,                     // Кол-во
			r.RetailPrice,                  // Цена розничная
			r.RetailAmount,                 // Вайлдберриз реализовал Товар (Пр)
			0,                              // Согласованный продуктовый дисконт, %
			"",                             // Промокод %
			0,                              // Итоговая согласованная скидка, %
			r.RetailPrice,                  // Цена розничная с учетом согласованной скидки
			0,                              // Размер снижения кВВ из-за рейтинга, %
			0,                              // Размер изменения кВВ из-за акции, %
			r.PpvzSppPrc,                   // Скидка постоянного Покупателя (СПП), %
			math.Round(r.CommissionPercent*100) / 100, // Размер кВВ, %
			math.Round(r.PpvzKvwPrcBase*100) / 100,    // Размер  кВВ без НДС, % Базовый
			math.Round(r.PpvzKvwPrc*100) / 100,        // Итоговый кВВ без НДС, %
			r.PpvzSalesCommission,                     // Вознаграждение с продаж до вычета услуг поверенного, без НДС
			0,                                         //Возмещение за выдачу и возврат товаров на ПВЗ
			r.AcquiringFee,                            // Эквайринг/Комиссии за организацию платежей
			r.AcquiringPercent,                        // Размер комиссии за эквайринг/Комиссии за организацию платежей, %
			r.PaymentProcessing,                       // Тип платежа за Эквайринг/Комиссии за организацию платежей
			math.Round(r.PpvzVw*100) / 100,            // Вознаграждение Вайлдберриз (ВВ), без НДС
			r.PpvzVwNds,                               // НДС с Вознаграждения Вайлдберриз
			r.PpvzForPay,                              // К перечислению Продавцу за реализованный Товар
			r.DeliveryAmount,                          // Количество доставок
			r.ReturnAmount,                            // Количество возврата
			r.DeliveryRub,                             // Услуги по доставке товара покупателю
			r.FixTariffDateFrom,                       // Дата начала действия фиксации
			r.FixTariffDateTo,                         // Дата конца действия фиксации
			"",                                        // Признак услуги платной доставки
			0,                                         // Общая сумма штрафов
			0,                                         // Корректировка Вознаграждения Вайлдберриз (ВВ)
			r.BonusTypeName,                           // Виды логистики, штрафов и корректировок ВВ
			r.StickerID,                               // Стикер МП
			r.AcquiringBank,                           // Наименование банка-эквайера
			r.PpvzOfficeID,                            // Номер офиса
			r.PpvzOfficeName,                          // Наименование офиса доставки
			"",                                        // ИНН партнера
			"",                                        // Партнер
			r.OfficeName,                              // Склад
			r.SiteCountry,                             // Страна
			r.GiBoxTypeName,                           // Тип коробов
			"",                                        // Номер таможенной декларации
			r.AssemblyID,                              // Номер сборочного задания
			r.Kiz,                                     // Код маркировки
			r.ShkID,                                   // ШК
			r.Srid,                                    // Srid
			r.RebillLogisticCost,                      // Возмещение издержек по перевозке/по складским операциям с товаром
			r.RebillLogisticOrg,                       // Организатор перевозки
			r.StorageFee,                              // Хранение
			r.Deduction,                               // Удержания
			r.Acceptance,                              // Платная приемка
			r.DlvPrc,                                  // Фиксированный коэффициент склада по поставке
			"Нет",                                     // Признак продажи юридическому лицу
			0,                                         // Номер короба для платной приемки
			0,                                         // Скидка по программе софинансирования
			"",                                        // Скидка Wibes, %
		}
		if err := sw.SetRow(fmt.Sprintf("A%d", row), data); err != nil {
			return nil, fmt.Errorf("failed to write row %d: %w", row, err)
		}

	}

	if err := sw.Flush(); err != nil {
		return nil, fmt.Errorf("failed to flush stream writer: %w", err)
	}

	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func GenerateReportExcel(reports []models.ReportDetails, taxPt, discountPt float64) ([]byte, error) {
	var grossRevenue float64          // Doanh thu gộp
	var netRevenue float64            // Doanh thu thuần
	var reductionInRevenue float64    // Giảm trừ doanh thu
	var logisticsExpenses float64     // Chi phí logistic
	var otherExpenses float64         // Chi phí khác
	var revenueExcludingCOGS float64  // Doanh thu chưa trừ giá vốn
	var estimatedCOGS float64         // Giá vốn ước lượng
	var revenueExcludingTaxes float64 // Doanh thu giảm trừ thuế
	var grossProfitToal float64       // Lãi gộp
	var tax float64                   // Thuế(6%)
	var netProfit float64             // Lãi ròng

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
		if r.SaName != "" && r.DocTypeName == "Продажа" {
			grossRevenue += r.RetailPrice
			netRevenue += r.PpvzForPay
			f.SetCellValue(sheet, fmt.Sprintf("A%d", row), r.SaName)
			f.SetCellValue(sheet, fmt.Sprintf("B%d", row), r.RetailPrice)
			f.SetCellValue(sheet, fmt.Sprintf("C%d", row), r.PpvzForPay)
			row++
		}
	}

	f.SetCellValue(sheet, "F1", "BẢNG HÀNG MUA BỊ TRẢ LẠI")
	f.MergeCell(sheet, "F1", "H1")
	f.SetCellStyle(sheet, "F1", "H1", headerStyleLight)
	f.SetCellValue(sheet, "F2", "Артикул поставщика")
	f.SetCellValue(sheet, "G2", "Giá gốc đăng bán")
	f.SetCellValue(sheet, "H2", "Giá trả lại")
	f.SetCellStyle(sheet, "F2", "H2", titleStyleDark)
	row = 3
	for _, r := range reports {
		if r.DocTypeName == "Возврат" {
			revenueExcludingTaxes += r.RetailPrice
			reductionInRevenue += r.PpvzForPay
			f.SetCellValue(sheet, fmt.Sprintf("F%d", row), r.SaName)
			f.SetCellValue(sheet, fmt.Sprintf("G%d", row), r.RetailPrice)
			f.SetCellValue(sheet, fmt.Sprintf("H%d", row), r.PpvzForPay)
			row++
		}
	}

	f.SetCellValue(sheet, "K1", "BẢNG PHÍ LOGISTIC")
	f.MergeCell(sheet, "K1", "L1")
	f.SetCellStyle(sheet, "K1", "L1", headerStyleLight)
	f.SetCellValue(sheet, "K2", "Артикул поставщика")
	f.SetCellValue(sheet, "L2", "Chi phí logistic")
	f.SetCellStyle(sheet, "K2", "L2", titleStyleDark)

	f.SetCellValue(sheet, "O1", "BẢNG PHÍ ĐƠN HÀNG BỊ HỦY OR KHÔNG MUA")
	f.MergeCell(sheet, "O1", "P1")
	f.SetCellStyle(sheet, "O1", "P1", headerStyleLight)
	f.SetCellValue(sheet, "O2", "Артикул поставщика")
	f.SetCellValue(sheet, "P2", "phí vận chuyển hàng trả lại")
	f.SetCellStyle(sheet, "O2", "P2", titleStyleDark)
	row = 3
	for _, r := range reports {
		if r.SupplierOperName == "Логистика" {
			logisticsExpenses += r.DeliveryRub
			f.SetCellValue(sheet, fmt.Sprintf("K%d", row), r.SaName)
			f.SetCellValue(sheet, fmt.Sprintf("L%d", row), r.DeliveryRub)
			if r.ReturnAmount == 1 {
				f.SetCellValue(sheet, fmt.Sprintf("O%d", row), r.SaName)
				f.SetCellValue(sheet, fmt.Sprintf("P%d", row), r.DeliveryRub)
			}
			row++
		}
	}

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
	var fines float64
	var storageCosts float64
	var advCosts float64
	var acceptanceCosts float64
	for _, r := range reports {
		fines += r.Penalty
		storageCosts += r.StorageFee
		advCosts += r.Deduction
		acceptanceCosts += r.Acceptance
	}
	otherExpenses = fines + storageCosts + advCosts + acceptanceCosts
	f.SetCellValue(sheet, "T3", fines)           // Tiền phạt
	f.SetCellValue(sheet, "T4", storageCosts)    // Chi phí lưu trữ
	f.SetCellValue(sheet, "T5", advCosts)        // Chi phí quảng cáo
	f.SetCellValue(sheet, "T6", acceptanceCosts) // Chi phí chấp nhận
	f.SetCellValue(sheet, "T7", otherExpenses)   // Tổng

	revenueExcludingCOGS = netRevenue - reductionInRevenue - logisticsExpenses - otherExpenses
	estimatedCOGS = (grossRevenue - revenueExcludingTaxes) / discountPt
	grossProfitToal = revenueExcludingCOGS - estimatedCOGS
	tax = (grossRevenue - revenueExcludingTaxes) * taxPt
	netProfit = grossProfitToal - tax
	f.SetCellValue(sheet, "W1", "BẢNG TỔNG KẾT")
	f.MergeCell(sheet, "W1", "AG1")
	f.SetCellStyle(sheet, "W1", "AG1", headerStyleLight)
	f.SetCellValue(sheet, "W2", "Doanh thu gộp(từ giá gốc)")
	f.SetCellValue(sheet, "X2", "Doanh thu thuần(đã trừ phí wb)")
	f.SetCellValue(sheet, "Y2", "Giảm trừ doanh thu(hàng trả lại)")
	f.SetCellValue(sheet, "Z2", "Chi phí logistic")
	f.SetCellValue(sheet, "AA2", "Chi phí khác")
	f.SetCellValue(sheet, "AB2", "Doanh thu chưa trừ giá vốn")
	f.SetCellValue(sheet, "AC2", "Giá vốn ước lượng")
	f.SetCellValue(sheet, "AD2", "Doanh thu giảm trừ thuế")
	f.SetCellValue(sheet, "AE2", "Lãi gộp")
	f.SetCellValue(sheet, "AF2", "Thuế(6%)")
	f.SetCellValue(sheet, "AG2", "Lãi ròng")
	f.SetCellStyle(sheet, "W2", "AG2", titleStyleDark)

	f.SetCellValue(sheet, "W3", math.Round(grossRevenue*100)/100)
	f.SetCellValue(sheet, "X3", math.Round(netRevenue*100)/100)
	f.SetCellValue(sheet, "Y3", math.Round(reductionInRevenue*100)/100)
	f.SetCellValue(sheet, "Z3", math.Round(logisticsExpenses*100)/100)
	f.SetCellValue(sheet, "AA3", math.Round(otherExpenses*100)/100)
	f.SetCellValue(sheet, "AB3", math.Round(revenueExcludingCOGS*100)/100)
	f.SetCellValue(sheet, "AC3", math.Round(estimatedCOGS*100)/100)
	f.SetCellValue(sheet, "AD3", math.Round(revenueExcludingTaxes*100)/100)
	f.SetCellValue(sheet, "AE3", math.Round(grossProfitToal*100)/100)
	f.SetCellValue(sheet, "AF3", math.Round(tax*100)/100)
	f.SetCellValue(sheet, "AG3", math.Round(netProfit*100)/100)

	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
