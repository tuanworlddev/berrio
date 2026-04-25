package models

import (
	"time"

	"go.mongodb.org/mongo-driver/bson"
	"go.mongodb.org/mongo-driver/bson/primitive"
)

type Shop struct {
	ID          primitive.ObjectID `bson:"_id,omitempty" json:"id"`
	Name        string             `bson:"name" json:"name" binding:"required"`
	Marketplace string             `bson:"marketplace,omitempty" json:"marketplace,omitempty"`
	APIKey      string             `bson:"apiKey,omitempty" json:"apiKey,omitempty"`
	Description string             `bson:"description,omitempty" json:"description,omitempty"`
	IsActive    bool               `bson:"isActive" json:"isActive"`
	Metadata    bson.M             `bson:"metadata,omitempty" json:"metadata,omitempty"`
	CreatedAt   time.Time          `bson:"createdAt" json:"createdAt"`
	UpdatedAt   time.Time          `bson:"updatedAt" json:"updatedAt"`
}

type CreateShopRequest struct {
	Name        string `json:"name" binding:"required"`
	Marketplace string `json:"marketplace,omitempty"`
	APIKey      string `json:"apiKey,omitempty"`
	Description string `json:"description,omitempty"`
	IsActive    *bool  `json:"isActive,omitempty"`
	Metadata    bson.M `json:"metadata,omitempty"`
}

type UpdateShopRequest struct {
	Name        *string `json:"name,omitempty"`
	Marketplace *string `json:"marketplace,omitempty"`
	APIKey      *string `json:"apiKey,omitempty"`
	Description *string `json:"description,omitempty"`
	IsActive    *bool   `json:"isActive,omitempty"`
	Metadata    bson.M  `json:"metadata,omitempty"`
}
