package services

import (
	"context"
	"errors"
	"math"
	"os"
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
)

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
