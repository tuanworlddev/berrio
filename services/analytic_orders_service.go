package services

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"time"
)

type AnalyticOrderRequest struct {
	Timezone string `json:"timezone"`
	Period   struct {
		Begin string `json:"begin"`
		End   string `json:"end"`
	} `json:"period"`
	OrderBy struct {
		Field string `json:"field"`
		Mode  string `json:"mode"`
	} `json:"orderBy"`
	Page int `json:"page"`
}

type AnalyticOrderResponse struct {
	Data struct {
		Page       int  `json:"page"`
		IsNextPage bool `json:"isNextPage"`
		Cards      []struct {
			NmID       int    `json:"nmID"`
			VendorCode string `json:"vendorCode"`
			Statistics struct {
				SelectedPeriod struct {
					OrdersCount  int `json:"ordersCount"`
					OrdersSumRub int `json:"ordersSumRub"`
				} `json:"selectedPeriod"`
				PreviousPeriod struct {
					OrdersCount  int `json:"ordersCount"`
					OrdersSumRub int `json:"ordersSumRub"`
				} `json:"previousPeriod"`
			} `json:"statistics"`
		} `json:"cards"`
	} `json:"data"`
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
	ChartData       []ChartData `json:"chartData"`
	TotalOrders     int         `json:"totalOrders"`
	TotalPrevOrders int         `json:"totalPrevOrders"`
}

func GetOrders(apiKey, begin, end string) (OrdersResponse, error) {
	payload := AnalyticOrderRequest{
		Timezone: "Europe/Moscow",
		Period: struct {
			Begin string `json:"begin"`
			End   string `json:"end"`
		}{
			Begin: begin,
			End:   end,
		},
		OrderBy: struct {
			Field string `json:"field"`
			Mode  string `json:"mode"`
		}{
			Field: "orders",
			Mode:  "desc",
		},
		Page: 1,
	}

	payloadBytes, err := json.Marshal(payload)
	if err != nil {
		fmt.Println(err)
		return OrdersResponse{}, err
	}

	req, err := http.NewRequest("POST", "https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail", bytes.NewBuffer(payloadBytes))
	if err != nil {
		fmt.Println(err)
		return OrdersResponse{}, err
	}

	req.Header.Set("Authorization", fmt.Sprintf("Bearer %s", apiKey))
	req.Header.Set("Content-Type", "application/json")

	client := &http.Client{Timeout: 10 * time.Second}
	resp, err := client.Do(req)
	if err != nil {
		fmt.Println(err)
		return OrdersResponse{}, err
	}
	defer resp.Body.Close()

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		fmt.Println(err)
		return OrdersResponse{}, err
	}

	var analyticOrderResponse AnalyticOrderResponse
	if err := json.Unmarshal(body, &analyticOrderResponse); err != nil {
		fmt.Println(err)
		return OrdersResponse{}, err
	}

	var chartData []ChartData
	var totalCountOrders int
	var totalPrevCountOders int
	for _, card := range analyticOrderResponse.Data.Cards {
		totalCountOrders += card.Statistics.SelectedPeriod.OrdersCount
		totalPrevCountOders += card.Statistics.PreviousPeriod.OrdersCount
		ordersCount := card.Statistics.SelectedPeriod.OrdersCount
		prevOrdersCount := card.Statistics.PreviousPeriod.OrdersCount

		if ordersCount > 0 || prevOrdersCount > 0 {
			chartData = append(chartData, ChartData{
				NmID:             card.NmID,
				VendorCode:       card.VendorCode,
				OrdersCount:      ordersCount,
				OrdersSumRub:     card.Statistics.SelectedPeriod.OrdersSumRub,
				PrevOrdersCount:  prevOrdersCount,
				PrevOrdersSumRub: card.Statistics.PreviousPeriod.OrdersSumRub,
			})
		}
	}

	return OrdersResponse{
		ChartData:       chartData,
		TotalOrders:     totalCountOrders,
		TotalPrevOrders: totalPrevCountOders,
	}, nil
}
