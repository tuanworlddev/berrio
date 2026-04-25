package services

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"math"
	"net/http"
	"strconv"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
	"omnituan.online/models"
)

var ErrReportRateLimited = errors.New("wildberries report API rate limited")

const reportRequestInterval = 61 * time.Second
const reportRateLimitRetries = 5
const financeReportURL = "https://finance-api.wildberries.ru/api/finance/v1/sales-reports/detailed"

type financeReportRequest struct {
	DateFrom string `json:"dateFrom"`
	DateTo   string `json:"dateTo"`
	Limit    int    `json:"limit"`
	RrdID    int64  `json:"rrdId"`
	Period   string `json:"period"`
}

type financeReportDetail struct {
	ReportID                    int64           `json:"reportId"`
	DateFrom                    string          `json:"dateFrom"`
	DateTo                      string          `json:"dateTo"`
	CreateDate                  string          `json:"createDate"`
	Currency                    string          `json:"currency"`
	ReportType                  int             `json:"reportType"`
	RrdID                       int64           `json:"rrdId"`
	GiID                        int64           `json:"giId"`
	DlvPrc                      flexibleFloat64 `json:"dlvPrc"`
	FixTariffDateFrom           string          `json:"fixTariffDateFrom"`
	FixTariffDateTo             string          `json:"fixTariffDateTo"`
	SubjectName                 string          `json:"subjectName"`
	NmID                        int64           `json:"nmId"`
	BrandName                   string          `json:"brandName"`
	VendorCode                  string          `json:"vendorCode"`
	TechSize                    string          `json:"techSize"`
	SKU                         string          `json:"sku"`
	DocTypeName                 string          `json:"docTypeName"`
	Quantity                    int             `json:"quantity"`
	RetailPrice                 flexibleFloat64 `json:"retailPrice"`
	RetailAmount                flexibleFloat64 `json:"retailAmount"`
	SalePercent                 int             `json:"salePercent"`
	CommissionPercent           flexibleFloat64 `json:"commissionPercent"`
	OfficeName                  string          `json:"officeName"`
	SellerOperName              string          `json:"sellerOperName"`
	OrderDt                     flexibleTime    `json:"orderDt"`
	SaleDt                      flexibleTime    `json:"saleDt"`
	RrDate                      string          `json:"rrDate"`
	ShkID                       int64           `json:"shkId"`
	RetailPriceWithDisc         flexibleFloat64 `json:"retailPriceWithDisc"`
	DeliveryAmount              int             `json:"deliveryAmount"`
	ReturnAmount                int             `json:"returnAmount"`
	DeliveryService             flexibleFloat64 `json:"deliveryService"`
	GiBoxTypeName               string          `json:"giBoxTypeName"`
	ProductDiscountForReport    flexibleFloat64 `json:"productDiscountForReport"`
	SellerPromo                 flexibleFloat64 `json:"sellerPromo"`
	SPP                         flexibleFloat64 `json:"spp"`
	KvwBase                     flexibleFloat64 `json:"kvwBase"`
	Kvw                         flexibleFloat64 `json:"kvw"`
	SupRatingUp                 flexibleFloat64 `json:"supRatingUp"`
	IsKgvpV2                    flexibleFloat64 `json:"isKgvpV2"`
	PpvzSalesCommission         flexibleFloat64 `json:"ppvzSalesCommission"`
	ForPay                      flexibleFloat64 `json:"forPay"`
	PpvzReward                  flexibleFloat64 `json:"ppvzReward"`
	AcquiringFee                flexibleFloat64 `json:"acquiringFee"`
	AcquiringPercent            flexibleFloat64 `json:"acquiringPercent"`
	PaymentProcessing           string          `json:"paymentProcessing"`
	AcquiringBank               string          `json:"acquiringBank"`
	Vw                          flexibleFloat64 `json:"vw"`
	VwNds                       flexibleFloat64 `json:"vwNds"`
	PpvzOfficeName              string          `json:"ppvzOfficeName"`
	PpvzOfficeID                int             `json:"ppvzOfficeId"`
	PpvzSupplierName            string          `json:"ppvzSupplierName"`
	PpvzSupplierInn             string          `json:"ppvzSupplierInn"`
	DeclarationNumber           string          `json:"declarationNumber"`
	BonusTypeName               string          `json:"bonusTypeName"`
	StickerID                   string          `json:"stickerId"`
	Country                     string          `json:"country"`
	SrvDbs                      bool            `json:"srvDbs"`
	Penalty                     flexibleFloat64 `json:"penalty"`
	AdditionalPayment           flexibleFloat64 `json:"additionalPayment"`
	RebillLogisticCost          flexibleFloat64 `json:"rebillLogisticCost"`
	RebillLogisticOrg           string          `json:"rebillLogisticOrg"`
	PaidStorage                 flexibleFloat64 `json:"paidStorage"`
	Deduction                   flexibleFloat64 `json:"deduction"`
	PaidAcceptance              flexibleFloat64 `json:"paidAcceptance"`
	OrderID                     int64           `json:"orderId"`
	Kiz                         string          `json:"kiz"`
	IsB2B                       bool            `json:"isB2b"`
	TrbxID                      string          `json:"trbxId"`
	InstallmentCofinancing      flexibleFloat64 `json:"installmentCofinancingAmount"`
	WibesDiscountPercent        int             `json:"wibesDiscountPercent"`
	Srid                        string          `json:"srid"`
}

type flexibleFloat64 float64

func (value *flexibleFloat64) UnmarshalJSON(data []byte) error {
	if string(data) == "null" || string(data) == `""` {
		*value = 0
		return nil
	}

	raw, err := strconv.Unquote(string(data))
	if err != nil {
		raw = string(data)
	}
	if raw == "" {
		*value = 0
		return nil
	}

	parsed, err := strconv.ParseFloat(raw, 64)
	if err != nil {
		return err
	}
	*value = flexibleFloat64(parsed)
	return nil
}

type flexibleTime time.Time

func (value *flexibleTime) UnmarshalJSON(data []byte) error {
	raw, err := strconv.Unquote(string(data))
	if err != nil || raw == "" {
		*value = flexibleTime(time.Time{})
		return nil
	}

	for _, layout := range []string{time.RFC3339, "2006-01-02"} {
		parsed, err := time.Parse(layout, raw)
		if err == nil {
			*value = flexibleTime(parsed)
			return nil
		}
	}
	*value = flexibleTime(time.Time{})
	return nil
}

func (value flexibleTime) Time() time.Time {
	return time.Time(value)
}

var reportLimiters sync.Map

type reportLimiter struct {
	mu       sync.Mutex
	nextCall time.Time
}

func waitReportRateLimit(ctx context.Context, apiKey string) error {
	value, _ := reportLimiters.LoadOrStore(apiKey, &reportLimiter{})
	limiter := value.(*reportLimiter)

	limiter.mu.Lock()
	defer limiter.mu.Unlock()

	now := time.Now()
	if wait := time.Until(limiter.nextCall); wait > 0 {
		timer := time.NewTimer(wait)
		select {
		case <-ctx.Done():
			timer.Stop()
			return ctx.Err()
		case <-timer.C:
		}
		now = time.Now()
	}

	limiter.nextCall = now.Add(reportRequestInterval)
	return nil
}

func GetReportDetails(ctx context.Context, apiKey string, dateFrom, dateTo time.Time) ([]models.ReportDetails, error) {
	var allReports []models.ReportDetails
	limit := 100000
	client := &http.Client{Timeout: 60 * time.Second}
	rrdid := int64(0) // Bắt đầu với rrdid = 0
	rateLimitRetries := 0

	for {
		if err := waitReportRateLimit(ctx, apiKey); err != nil {
			return nil, err
		}

		payload, err := json.Marshal(financeReportRequest{
			DateFrom: dateFrom.Format("2006-01-02"),
			DateTo:   dateTo.Format("2006-01-02"),
			Limit:    limit,
			RrdID:    rrdid,
			Period:   "daily",
		})
		if err != nil {
			return nil, fmt.Errorf("failed to encode request: %v", err)
		}

		// Tạo request
		req, err := http.NewRequestWithContext(ctx, "POST", financeReportURL, bytes.NewReader(payload))
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

		// Xử lý rate limit (429)
		if res.StatusCode == 429 {
			_ = res.Body.Close()
			rateLimitRetries++
			if rateLimitRetries > reportRateLimitRetries {
				return nil, fmt.Errorf("%w: retry later", ErrReportRateLimited)
			}
			continue
		}

		if res.StatusCode == http.StatusNoContent {
			_ = res.Body.Close()
			break
		}

		// Kiểm tra status code
		if res.StatusCode != http.StatusOK {
			body, _ := io.ReadAll(res.Body)
			_ = res.Body.Close()
			return nil, fmt.Errorf("error response: status code %d, body: %s", res.StatusCode, string(body))
		}

		// Đọc và parse body
		body, err := io.ReadAll(res.Body)
		_ = res.Body.Close()
		if err != nil {
			return nil, fmt.Errorf("failed to read response: %v", err)
		}

		var details []financeReportDetail
		if err := json.Unmarshal(body, &details); err != nil {
			return nil, fmt.Errorf("failed to decode JSON: %v", err)
		}
		rateLimitRetries = 0

		// Thoát nếu không còn dữ liệu
		if len(details) == 0 {
			fmt.Println("No more data to fetch.")
			break
		}

		reports := mapFinanceReportDetails(details)
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

func mapFinanceReportDetails(details []financeReportDetail) []models.ReportDetails {
	reports := make([]models.ReportDetails, 0, len(details))
	for _, item := range details {
		reports = append(reports, models.ReportDetails{
			RealizationReportID:          item.ReportID,
			DateFrom:                     item.DateFrom,
			DateTo:                       item.DateTo,
			CreateDt:                     item.CreateDate,
			CurrencyName:                 item.Currency,
			RrdID:                        item.RrdID,
			GiID:                         item.GiID,
			DlvPrc:                       float64(item.DlvPrc),
			FixTariffDateFrom:            item.FixTariffDateFrom,
			FixTariffDateTo:              item.FixTariffDateTo,
			SubjectName:                  item.SubjectName,
			NmID:                         item.NmID,
			BrandName:                    item.BrandName,
			SaName:                       item.VendorCode,
			TsName:                       item.TechSize,
			Barcode:                      item.SKU,
			DocTypeName:                  item.DocTypeName,
			Quantity:                     item.Quantity,
			RetailPrice:                  float64(item.RetailPrice),
			RetailAmount:                 float64(item.RetailAmount),
			SalePercent:                  item.SalePercent,
			CommissionPercent:            float64(item.CommissionPercent),
			OfficeName:                   item.OfficeName,
			SupplierOperName:             item.SellerOperName,
			OrderDt:                      item.OrderDt.Time(),
			SaleDt:                       item.SaleDt.Time(),
			RrDt:                         item.RrDate,
			ShkID:                        item.ShkID,
			RetailPriceWithDiscRub:       float64(item.RetailPriceWithDisc),
			DeliveryAmount:               item.DeliveryAmount,
			ReturnAmount:                 item.ReturnAmount,
			DeliveryRub:                  float64(item.DeliveryService),
			GiBoxTypeName:                item.GiBoxTypeName,
			ProductDiscountForReport:     float64(item.ProductDiscountForReport),
			SupplierPromo:                float64(item.SellerPromo),
			PpvzSppPrc:                   float64(item.SPP),
			PpvzKvwPrcBase:               float64(item.KvwBase),
			PpvzKvwPrc:                   float64(item.Kvw),
			SupRatingPrcUp:               float64(item.SupRatingUp),
			IsKgvpV2:                     float64(item.IsKgvpV2),
			PpvzSalesCommission:          float64(item.PpvzSalesCommission),
			PpvzForPay:                   float64(item.ForPay),
			PpvzReward:                   float64(item.PpvzReward),
			AcquiringFee:                 float64(item.AcquiringFee),
			AcquiringPercent:             float64(item.AcquiringPercent),
			PaymentProcessing:            item.PaymentProcessing,
			AcquiringBank:                item.AcquiringBank,
			PpvzVw:                       float64(item.Vw),
			PpvzVwNds:                    float64(item.VwNds),
			PpvzOfficeName:               item.PpvzOfficeName,
			PpvzOfficeID:                 item.PpvzOfficeID,
			PpvzSupplierName:             item.PpvzSupplierName,
			PpvzInn:                      item.PpvzSupplierInn,
			DeclarationNumber:            item.DeclarationNumber,
			BonusTypeName:                item.BonusTypeName,
			StickerID:                    item.StickerID,
			SiteCountry:                  item.Country,
			SrvDbs:                       item.SrvDbs,
			Penalty:                      float64(item.Penalty),
			AdditionalPayment:            float64(item.AdditionalPayment),
			RebillLogisticCost:           float64(item.RebillLogisticCost),
			RebillLogisticOrg:            item.RebillLogisticOrg,
			StorageFee:                   float64(item.PaidStorage),
			Deduction:                    float64(item.Deduction),
			Acceptance:                   float64(item.PaidAcceptance),
			AssemblyID:                   item.OrderID,
			Kiz:                          item.Kiz,
			Srid:                         item.Srid,
			ReportType:                   item.ReportType,
			IsLegalEntity:                item.IsB2B,
			TrbxID:                       item.TrbxID,
			InstallmentCofinancingAmount: float64(item.InstallmentCofinancing),
			WibesWbDiscountPercent:       item.WibesDiscountPercent,
		})
	}
	return reports
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
	netProfit = grossProfitToal - taxFinal
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
	f.SetCellValue(sheet, "AG3", math.Round(taxFinal*100)/100)
	f.SetCellValue(sheet, "AH3", math.Round(netProfit*100)/100)

	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
