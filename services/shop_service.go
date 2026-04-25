package services

import (
	"context"
	"encoding/json"
	"errors"
	"math"
	"net/http"
	"os"
	"sync"
	"time"

	"go.mongodb.org/mongo-driver/bson"
	"go.mongodb.org/mongo-driver/bson/primitive"
	"go.mongodb.org/mongo-driver/mongo"
	"go.mongodb.org/mongo-driver/mongo/options"
	"omnituan.online/database"
	"omnituan.online/models"
)

const (
	defaultShopCollection = "shops"
	defaultPage           = int64(1)
	defaultLimit          = int64(10)
	maxLimit              = int64(100)
	wbPingURL             = "https://common-api.wildberries.ru/ping"
	tokenStatusCacheTTL   = 5 * time.Minute
)

var tokenStatusCache sync.Map

type cachedTokenStatus struct {
	status    models.ShopTokenStatus
	expiresAt time.Time
}

func shopCollection() *mongo.Collection {
	name := os.Getenv("MONGO_SHOP_COLLECTION")
	if name == "" {
		name = defaultShopCollection
	}

	return database.Collection(name)
}

func CreateShop(ctx context.Context, req models.CreateShopRequest) (models.Shop, error) {
	now := time.Now().UTC()
	isActive := true
	if req.IsActive != nil {
		isActive = *req.IsActive
	}

	shop := models.Shop{
		ID:          primitive.NewObjectID(),
		Name:        req.Name,
		Marketplace: req.Marketplace,
		APIKey:      req.APIKey,
		Description: req.Description,
		IsActive:    isActive,
		Metadata:    req.Metadata,
		CreatedAt:   now,
		UpdatedAt:   now,
	}

	_, err := shopCollection().InsertOne(ctx, shop)
	return shop, err
}

func GetShops(ctx context.Context, page, limit int64) (models.PaginationResponse[models.Shop], error) {
	page, limit = normalizePagination(page, limit)
	skip := (page - 1) * limit

	filter := bson.M{}
	findOptions := options.Find().
		SetSkip(skip).
		SetLimit(limit).
		SetSort(bson.D{{Key: "createdAt", Value: -1}})

	total, err := shopCollection().CountDocuments(ctx, filter)
	if err != nil {
		return models.PaginationResponse[models.Shop]{}, err
	}

	cursor, err := shopCollection().Find(ctx, filter, findOptions)
	if err != nil {
		return models.PaginationResponse[models.Shop]{}, err
	}
	defer cursor.Close(ctx)

	var shops []models.Shop
	if err := cursor.All(ctx, &shops); err != nil {
		return models.PaginationResponse[models.Shop]{}, err
	}
	if shops == nil {
		shops = []models.Shop{}
	}

	return models.PaginationResponse[models.Shop]{
		Items:      shops,
		Page:       page,
		Limit:      limit,
		Total:      total,
		TotalPages: int64(math.Ceil(float64(total) / float64(limit))),
	}, nil
}

func GetShopByID(ctx context.Context, id string) (models.Shop, error) {
	objectID, err := primitive.ObjectIDFromHex(id)
	if err != nil {
		return models.Shop{}, err
	}

	var shop models.Shop
	err = shopCollection().FindOne(ctx, bson.M{"_id": objectID}).Decode(&shop)
	return shop, err
}

func UpdateShop(ctx context.Context, id string, req models.UpdateShopRequest) (models.Shop, error) {
	objectID, err := primitive.ObjectIDFromHex(id)
	if err != nil {
		return models.Shop{}, err
	}

	set := bson.M{"updatedAt": time.Now().UTC()}
	if req.Name != nil {
		set["name"] = *req.Name
	}
	if req.Marketplace != nil {
		set["marketplace"] = *req.Marketplace
	}
	if req.APIKey != nil {
		set["apiKey"] = *req.APIKey
	}
	if req.Description != nil {
		set["description"] = *req.Description
	}
	if req.IsActive != nil {
		set["isActive"] = *req.IsActive
	}
	if req.Metadata != nil {
		set["metadata"] = req.Metadata
	}

	updateOptions := options.FindOneAndUpdate().SetReturnDocument(options.After)
	var shop models.Shop
	err = shopCollection().
		FindOneAndUpdate(ctx, bson.M{"_id": objectID}, bson.M{"$set": set}, updateOptions).
		Decode(&shop)
	return shop, err
}

func DeleteShop(ctx context.Context, id string) error {
	objectID, err := primitive.ObjectIDFromHex(id)
	if err != nil {
		return err
	}

	result, err := shopCollection().DeleteOne(ctx, bson.M{"_id": objectID})
	if err != nil {
		return err
	}
	if result.DeletedCount == 0 {
		return mongo.ErrNoDocuments
	}

	return nil
}

func CheckShopToken(ctx context.Context, id string) (models.ShopTokenStatus, error) {
	shop, err := GetShopByID(ctx, id)
	if err != nil {
		return models.ShopTokenStatus{}, err
	}

	cacheKey := shop.ID.Hex() + ":" + shop.UpdatedAt.Format(time.RFC3339Nano)
	if cached, ok := tokenStatusCache.Load(cacheKey); ok {
		item := cached.(cachedTokenStatus)
		if time.Now().Before(item.expiresAt) {
			return item.status, nil
		}
		tokenStatusCache.Delete(cacheKey)
	}

	status := models.ShopTokenStatus{
		ShopID:    shop.ID.Hex(),
		CheckedAt: time.Now().UTC().Format(time.RFC3339),
	}

	if shop.APIKey == "" {
		status.Status = "missing"
		status.Message = "Shop chưa có API key"
		return status, nil
	}

	req, err := http.NewRequestWithContext(ctx, http.MethodGet, wbPingURL, nil)
	if err != nil {
		return models.ShopTokenStatus{}, err
	}
	req.Header.Set("Authorization", "Bearer "+shop.APIKey)

	client := &http.Client{Timeout: 10 * time.Second}
	resp, err := client.Do(req)
	if err != nil {
		status.Status = "error"
		status.Message = "Không kết nối được WB API"
		return status, nil
	}
	defer resp.Body.Close()

	var body struct {
		TS     string `json:"TS"`
		Status string `json:"Status"`
	}
	_ = json.NewDecoder(resp.Body).Decode(&body)

	status.WBTimestamp = body.TS
	switch resp.StatusCode {
	case http.StatusOK:
		status.Valid = body.Status == "OK"
		status.Status = "ok"
		if !status.Valid {
			status.Status = "invalid"
			status.Message = "WB API không trả trạng thái OK"
		}
	case http.StatusUnauthorized:
		status.Status = "invalid"
		status.Message = "Token không hợp lệ, hết hạn hoặc sai category"
	case http.StatusTooManyRequests:
		status.Status = "rate_limited"
		status.Message = "WB đang giới hạn kiểm tra token, thử lại sau"
	default:
		status.Status = "error"
		status.Message = "WB API trả lỗi khi kiểm tra token"
	}

	tokenStatusCache.Store(cacheKey, cachedTokenStatus{
		status:    status,
		expiresAt: time.Now().Add(tokenStatusCacheTTL),
	})
	return status, nil
}

func IsNotFound(err error) bool {
	return errors.Is(err, mongo.ErrNoDocuments)
}

func normalizePagination(page, limit int64) (int64, int64) {
	if page < 1 {
		page = defaultPage
	}
	if limit < 1 {
		limit = defaultLimit
	}
	if limit > maxLimit {
		limit = maxLimit
	}

	return page, limit
}
