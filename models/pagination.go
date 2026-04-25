package models

type PaginationResponse[T any] struct {
	Items      []T   `json:"items"`
	Page       int64 `json:"page"`
	Limit      int64 `json:"limit"`
	Total      int64 `json:"total"`
	TotalPages int64 `json:"totalPages"`
}
