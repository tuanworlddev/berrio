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
	limit := 100000
	rrdid := int64(0) // Bắt đầu với rrdid = 0

	for {
		// Tạo URL với dateFrom, dateTo, limit và rrdid
		url := fmt.Sprintf(
			"https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod?dateFrom=%s&dateTo=%s&limit=%d&rrdid=%d",
			dateFrom.Format("2006-01-02"),
			dateTo.Format("2006-01-02"),
			limit,
			rrdid,
		)

		client := &http.Client{}

		// Tạo request
		req, err := http.NewRequest("GET", url, nil)
		if err != nil {
			return nil, fmt.Errorf("failed to create request: %v", err)
		}
		req.Header.Set("Authorization", fmt.Sprintf("Bearer %s", apiKey))
		req.Header.Set("Content-Type", "application/json")

		// Gửi request
		res, err := client.Do(req)
		if err != nil {
			return nil, fmt.Errorf("failed to make request: %v", err)
		}
		defer res.Body.Close()

		// Xử lý rate limit (429)
		if res.StatusCode == 429 {
			fmt.Println("Rate limit exceeded (429), waiting for 1 minute...")
			time.Sleep(1 * time.Minute)
			continue
		}

		// Kiểm tra status code
		if res.StatusCode != http.StatusOK {
			body, _ := io.ReadAll(res.Body)
			return nil, fmt.Errorf("error response: status code %d, body: %s", res.StatusCode, string(body))
		}

		// Đọc và parse body
		body, err := io.ReadAll(res.Body)
		if err != nil {
			return nil, fmt.Errorf("failed to read response: %v", err)
		}

		var reports []models.ReportDetails
		if err := json.Unmarshal(body, &reports); err != nil {
			return nil, fmt.Errorf("failed to decode JSON: %v", err)
		}

		// Thoát nếu không còn dữ liệu
		if len(reports) == 0 {
			fmt.Println("No more data to fetch.")
			break
		}

		allReports = append(allReports, reports...)

		// Cập nhật rrdid từ bản ghi cuối cùng
		rrdid = reports[len(reports)-1].RrdID
		fmt.Printf("Fetched %d records, next rrdid: %d\n", len(reports), rrdid)

		if len(reports) < limit {
			fmt.Println("Reached end of data (less than limit).")
			break
		}
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

	headers := []any{
		"STT",                              // №
		"Mã giao hàng",                     // Номер поставки
		"Loại sản phẩm",                    // Предмет
		"Mã hàng",                          // Код номенклатуры
		"Thương hiệu",                      // Бренд
		"Mã nhà cung cấp",                  // Артикул поставщика
		"Tên sản phẩm",                     // Название
		"Kích thước",                       // Размер
		"Mã vạch",                          // Баркод
		"Loại tài liệu",                    // Тип документа
		"Lý do giao dịch",                  // Обоснование для оплаты
		"Ngày đặt hàng",                    // Дата заказа покупателем
		"Ngày bán",                         // Дата продажи
		"Số lượng",                         // Кол-во
		"Giá niêm yết",                     // Цена розничная
		"Doanh thu Wildberries (đã bán)",   // Вайлдберриз реализовал Товар (Пр)
		"Giảm giá theo thỏa thuận (%)",     // Согласованный продуктовый дисконт, %
		"Khuyến mãi mã giảm (%)",           // Промокод %
		"Tổng giảm giá sau thỏa thuận (%)", // Итоговая согласованная скидка, %
		"Giá sau giảm",                     // Цена розничная с учетом согласованной скидки
		"Giảm giá do đánh giá (%)",         // Размер снижения кВВ из-за рейтинга, %
		"Giảm giá do khuyến mãi (%)",       // Размер изменения кВВ из-за акции, %
		"Chiết khấu khách hàng thân thiết (SPP) (%)", // Скидка постоянного Покупателя (СПП), %
		"Hoa hồng (%)",                    // Размер кВВ, %
		"Hoa hồng cơ bản không VAT (%)",   // Размер  кВВ без НДС, % Базовый
		"Hoa hồng cuối không VAT (%)",     // Итоговый кВВ без НДС, %
		"Hoa hồng Wildberries (chưa VAT)", // Вознаграждение с продаж до вычета услуг поверенного, без НДС
		"Hoàn phí giao/hoàn trả",          // Возмещение за выдачу и возврат товаров на ПВЗ
		"Phí thanh toán",                  // Эквайринг/Комиссии за организацию платежей
		"Tỷ lệ phí thanh toán (%)",        // Размер комиссии за эквайринг/Комиссии за организацию платежей, %
		"Hình thức thanh toán",            // Тип платежа за Эквайринг/Комиссии за организацию платежей
		"Phí Wildberries (chưa VAT)",      // Вознаграждение Вайлдберриз (ВВ), без НДС
		"VAT trên phí Wildberries",        // НДС с Вознаграждения Вайлдберриз
		"Tiền thực nhận",                  // К перечислению Продавцу за реализованный Товар
		"Số lần giao",                     // Количество доставок
		"Số lần hoàn",                     // Количество возврата
		"Chi phí giao hàng",               // Услуги по доставке товара покупателю
		"Ngày bắt đầu phí cố định",        // Дата начала действия фиксации
		"Ngày kết thúc phí cố định",       // Дата конца действия фиксации
		"Dịch vụ giao hàng có tính phí",   // Признак услуги платной доставки
		"Tổng tiền phạt",                  // Общая сумма штрафов
		"Điều chỉnh phí Wildberries",      // Корректировка Вознаграждения Вайлдберриз (ВВ)
		"Loại logistics/phạt/điều chỉnh",  // Виды логистики, штрафов и корректировок ВВ
		"Mã nhãn dán (Sticker MP)",        // Стикер МП
		"Ngân hàng thanh toán",            // Наименование банка-эквайера
		"Mã văn phòng",                    // Номер офиса
		"Tên văn phòng giao hàng",         // Наименование офиса доставки
		"Mã số thuế đối tác",              // ИНН партнера
		"Tên đối tác",                     // Партнер
		"Kho hàng",                        // Склад
		"Quốc gia",                        // Страна
		"Loại hộp",                        // Тип коробов
		"Số tờ khai hải quan",             // Номер таможенной декларации
		"Mã đơn lắp ráp",                  // Номер сборочного задания
		"Mã định danh (KIZ)",              // Код маркировки
		"Mã sản phẩm (ШК)",                // ШК
		"Mã giao dịch (Srid)",             // Srid
		"Hoàn phí vận chuyển/kho",         // Возмещение издержек по перевозке/по складским операциям с товаром
		"Đơn vị vận chuyển",               // Организатор перевозки
		"Phí lưu kho",                     // Хранение
		"Khoản trừ khác",                  // Удержания
		"Phí nhận hàng",                   // Платная приемка
		"Hệ số kho cố định",               // Фиксированный коэффициент склада по поставке
		"Bán cho công ty",                 // Признак продажи юридическому лицу
		"Số hộp nhận hàng tính phí",       // Номер короба для платной приемки
		"Giảm giá đồng tài trợ",           // Скидка по программе софинансирования
		"Giảm giá Wibes (%)",              // Скидка Wibes, %
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
	var tax float64                   // Thuế(%)
	var taxFinal float64              // Thuế phải đóng
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
	row = 3
	for _, r := range reports {
		if r.SupplierOperName == "Логистика" {
			logisticsExpenses += r.DeliveryRub
			f.SetCellValue(sheet, fmt.Sprintf("K%d", row), r.SaName)
			f.SetCellValue(sheet, fmt.Sprintf("L%d", row), r.DeliveryRub)
			row++
		}
	}

	f.SetCellValue(sheet, "O1", "BẢNG PHÍ ĐƠN HÀNG BỊ HỦY OR KHÔNG MUA")
	f.MergeCell(sheet, "O1", "P1")
	f.SetCellStyle(sheet, "O1", "P1", headerStyleLight)
	f.SetCellValue(sheet, "O2", "Артикул поставщика")
	f.SetCellValue(sheet, "P2", "phí vận chuyển hàng trả lại")
	f.SetCellStyle(sheet, "O2", "P2", titleStyleDark)
	row = 3
	for _, r := range reports {
		if r.SupplierOperName == "Логистика" && r.ReturnAmount == 1 {
			f.SetCellValue(sheet, fmt.Sprintf("O%d", row), r.SaName)
			f.SetCellValue(sheet, fmt.Sprintf("P%d", row), r.DeliveryRub)
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
	taxFinal = (netRevenue - reductionInRevenue) * taxPt
	netProfit = grossProfitToal - tax
	f.SetCellValue(sheet, "W1", "BẢNG TỔNG KẾT")
	f.MergeCell(sheet, "W1", "AH1")
	f.SetCellStyle(sheet, "W1", "AH1", headerStyleLight)
	f.SetCellValue(sheet, "W2", "Doanh thu theo giá gốc sản phẩm")
	f.SetCellValue(sheet, "X2", "Doanh thu sau khi trừ phí WB")
	f.SetCellValue(sheet, "Y2", "Giảm trừ doanh thu(hàng trả lại)")
	f.SetCellValue(sheet, "Z2", "Chi phí logistic")
	f.SetCellValue(sheet, "AA2", "Chi phí khác")
	f.SetCellValue(sheet, "AB2", "Doanh thu chưa trừ giá vốn")
	f.SetCellValue(sheet, "AC2", "Giá vốn ước lượng")
	f.SetCellValue(sheet, "AD2", "Doanh thu giảm trừ thuế")
	f.SetCellValue(sheet, "AE2", "Lãi trước thuế và chi phí khác")
	f.SetCellValue(sheet, "AF2", fmt.Sprintf("Thuế(%.2f%%)", taxPt*100))
	f.SetCellValue(sheet, "AG2", "Thuế phải đóng")
	f.SetCellValue(sheet, "AH2", "Lợi nhuận thực nhận về sau khi trừ toàn bộ phí")
	f.SetCellStyle(sheet, "W2", "AH2", titleStyleDark)

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
	f.SetCellValue(sheet, "AH3", math.Round(taxFinal*100)/100)
	f.SetCellValue(sheet, "AH3", math.Round(netProfit*100)/100)

	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
