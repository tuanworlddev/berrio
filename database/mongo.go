package database

import (
	"context"
	"fmt"
	"os"
	"time"

	"go.mongodb.org/mongo-driver/mongo"
	"go.mongodb.org/mongo-driver/mongo/options"
)

var Client *mongo.Client

func Connect(ctx context.Context) error {
	uri := os.Getenv("MONGO_URI")
	if uri == "" {
		return fmt.Errorf("MONGO_URI is required")
	}

	connectCtx, cancel := context.WithTimeout(ctx, 10*time.Second)
	defer cancel()

	client, err := mongo.Connect(connectCtx, options.Client().ApplyURI(uri))
	if err != nil {
		return err
	}

	pingCtx, cancel := context.WithTimeout(ctx, 5*time.Second)
	defer cancel()
	if err := client.Ping(pingCtx, nil); err != nil {
		return err
	}

	Client = client
	return nil
}

func Disconnect(ctx context.Context) error {
	if Client == nil {
		return nil
	}

	disconnectCtx, cancel := context.WithTimeout(ctx, 5*time.Second)
	defer cancel()
	return Client.Disconnect(disconnectCtx)
}

func Collection(name string) *mongo.Collection {
	dbName := os.Getenv("MONGO_DATABASE")
	if dbName == "" {
		dbName = "shopdb"
	}

	return Client.Database(dbName).Collection(name)
}
