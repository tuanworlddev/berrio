package services

import (
	"bytes"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"strconv"
	"time"
)

var (
	ErrOrdersUnauthorized = errors.New("wildberries analytics token is invalid")
	ErrOrdersForbidden    = errors.New("wildberries analytics access denied")
	ErrOrdersRateLimited  = errors.New("wildberries analytics API rate limited")
)

const (
	analyticOrdersURL        = "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products"
	analyticOrdersLimit      = 1000
	analyticOrdersMaxPages   = 100
	analyticOrdersMaxRetries = 3
	analyticOrdersRetryWait  = 20 * time.Second
)

type analyticOrdersRequest struct {
	SelectedPeriod periodRequest `json:"selectedPeriod"`
	PastPeriod     periodRequest `json:"pastPeriod"`
	BrandNames     []string      `json:"brandNames"`
	ObjectIDs      []int         `json:"objectIDs"`
	TagIDs         []int         `json:"tagIDs"`
	NmIDs          []int         `json:"nmIDs"`
	Timezone       string        `json:"timezone"`
	OrderBy        orderRequest  `json:"orderBy"`
	Limit          int           `json:"limit"`
	Offset         int           `json:"offset"`
	SkipDeletedNm  bool          `json:"skipDeletedNm"`
}

type periodRequest struct {
	Start string `json:"start"`
	End   string `json:"end"`
}

type orderRequest struct {
	Field string `json:"field"`
	Mode  string `json:"mode"`
}

type analyticOrdersResponse struct {
	Data struct {
		Products []struct {
			Product struct {
				NmID       int    `json:"nmId"`
				VendorCode string `json:"vendorCode"`
			} `json:"product"`
			Statistic struct {
				Selected ordersStatistic `json:"selected"`
				Past     ordersStatistic `json:"past"`
			} `json:"statistic"`
		} `json:"products"`
	} `json:"data"`
}

type ordersStatistic struct {
	OrderCount int `json:"orderCount"`
	OrderSum   int `json:"orderSum"`
}

type ChartData struct {
	NmID             int    `json:"nmID"`
	VendorCode       string `json:"vendorCode"`
	OrdersCount      int    `json:"ordersCount"`
	OrdersSumRub     int    `json:"ordersSumRub"`
	PrevOrdersCount  int    `json:"prevOrdersCount"`
	PrevOrdersSumRub int    `json:"prevOrdersSumRub"`
}

type OrdersResponse struct {
	ChartData         []ChartData `json:"chartData"`
	TotalOrders       int         `json:"totalOrders"`
	TotalPrevOrders   int         `json:"totalPrevOrders"`
	TotalOrdersSumRub int         `json:"totalOrdersSumRub"`
	TotalPrevSumRub   int         `json:"totalPrevOrdersSumRub"`
	PagesLoaded       int         `json:"pagesLoaded"`
	HasMorePages      bool        `json:"hasMorePages"`
}

func GetOrders(ctx context.Context, apiKey, begin, end string) (OrdersResponse, error) {
	selectedStart, err := time.Parse("2006-01-02", begin)
	if err != nil {
		return OrdersResponse{}, err
	}
	selectedEnd, err := time.Parse("2006-01-02", end)
	if err != nil {
		return OrdersResponse{}, err
	}

	client := &http.Client{Timeout: 30 * time.Second}
	result := OrdersResponse{ChartData: []ChartData{}}

	for page := 0; page < analyticOrdersMaxPages; page++ {
		response, err := fetchAnalyticOrdersPage(ctx, client, apiKey, selectedStart, selectedEnd, page*analyticOrdersLimit)
		if err != nil {
			return OrdersResponse{}, err
		}

		result.PagesLoaded = page + 1
		appendAnalyticOrders(&result, response)

		if len(response.Data.Products) < analyticOrdersLimit {
			return result, nil
		}
	}

	result.HasMorePages = true
	return result, nil
}

func fetchAnalyticOrdersPage(ctx context.Context, client *http.Client, apiKey string, selectedStart, selectedEnd time.Time, offset int) (analyticOrdersResponse, error) {
	payload := newAnalyticOrderPayload(selectedStart, selectedEnd, offset)
	payloadBytes, err := json.Marshal(payload)
	if err != nil {
		return analyticOrdersResponse{}, err
	}

	var lastErr error
	for attempt := 0; attempt <= analyticOrdersMaxRetries; attempt++ {
		req, err := http.NewRequestWithContext(ctx, http.MethodPost, analyticOrdersURL, bytes.NewReader(payloadBytes))
		if err != nil {
			return analyticOrdersResponse{}, err
		}
		req.Header.Set("Authorization", "Bearer "+apiKey)
		req.Header.Set("Content-Type", "application/json")

		resp, err := client.Do(req)
		if err != nil {
			return analyticOrdersResponse{}, err
		}

		if resp.StatusCode == http.StatusTooManyRequests {
			_ = resp.Body.Close()
			lastErr = ErrOrdersRateLimited
			if attempt == analyticOrdersMaxRetries {
				break
			}
			if err := sleepWithContext(ctx, retryAfter(resp.Header.Get("Retry-After"))); err != nil {
				return analyticOrdersResponse{}, err
			}
			continue
		}

		if err := ensureAnalyticOrdersStatus(resp); err != nil {
			return analyticOrdersResponse{}, err
		}

		var decoded analyticOrdersResponse
		if err := json.NewDecoder(resp.Body).Decode(&decoded); err != nil {
			_ = resp.Body.Close()
			return analyticOrdersResponse{}, err
		}
		_ = resp.Body.Close()
		return decoded, nil
	}

	return analyticOrdersResponse{}, lastErr
}

func newAnalyticOrderPayload(selectedStart, selectedEnd time.Time, offset int) analyticOrdersRequest {
	days := int(selectedEnd.Sub(selectedStart).Hours()/24) + 1
	pastEnd := selectedStart.AddDate(0, 0, -1)
	pastStart := pastEnd.AddDate(0, 0, -days+1)

	return analyticOrdersRequest{
		SelectedPeriod: periodRequest{
			Start: selectedStart.Format("2006-01-02"),
			End:   selectedEnd.Format("2006-01-02"),
		},
		PastPeriod: periodRequest{
			Start: pastStart.Format("2006-01-02"),
			End:   pastEnd.Format("2006-01-02"),
		},
		BrandNames:    []string{},
		ObjectIDs:     []int{},
		TagIDs:        []int{},
		NmIDs:         []int{},
		Timezone:      "Europe/Moscow",
		OrderBy:       orderRequest{Field: "orderCount", Mode: "desc"},
		Limit:         analyticOrdersLimit,
		Offset:        offset,
		SkipDeletedNm: true,
	}
}

func ensureAnalyticOrdersStatus(resp *http.Response) error {
	if resp.StatusCode == http.StatusOK {
		return nil
	}
	defer resp.Body.Close()

	switch resp.StatusCode {
	case http.StatusUnauthorized:
		return ErrOrdersUnauthorized
	case http.StatusForbidden, http.StatusPaymentRequired:
		return ErrOrdersForbidden
	}

	body, _ := io.ReadAll(io.LimitReader(resp.Body, 2048))
	return fmt.Errorf("wildberries analytics API error: status %d: %s", resp.StatusCode, string(body))
}

func appendAnalyticOrders(result *OrdersResponse, response analyticOrdersResponse) {
	for _, item := range response.Data.Products {
		ordersCount := item.Statistic.Selected.OrderCount
		prevOrdersCount := item.Statistic.Past.OrderCount
		ordersSum := item.Statistic.Selected.OrderSum
		prevOrdersSum := item.Statistic.Past.OrderSum

		result.TotalOrders += ordersCount
		result.TotalPrevOrders += prevOrdersCount
		result.TotalOrdersSumRub += ordersSum
		result.TotalPrevSumRub += prevOrdersSum

		if ordersCount > 0 || prevOrdersCount > 0 {
			result.ChartData = append(result.ChartData, ChartData{
				NmID:             item.Product.NmID,
				VendorCode:       item.Product.VendorCode,
				OrdersCount:      ordersCount,
				OrdersSumRub:     ordersSum,
				PrevOrdersCount:  prevOrdersCount,
				PrevOrdersSumRub: prevOrdersSum,
			})
		}
	}
}

func retryAfter(value string) time.Duration {
	seconds, err := strconv.Atoi(value)
	if err != nil || seconds <= 0 {
		return analyticOrdersRetryWait
	}
	return time.Duration(seconds) * time.Second
}

func sleepWithContext(ctx context.Context, duration time.Duration) error {
	timer := time.NewTimer(duration)
	defer timer.Stop()

	select {
	case <-ctx.Done():
		return ctx.Err()
	case <-timer.C:
		return nil
	}
}
