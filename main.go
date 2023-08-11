package main

import (
	"context"
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"time"

	"github.com/xuri/excelize/v2"
	"go.mongodb.org/mongo-driver/bson"
	"go.mongodb.org/mongo-driver/mongo"
	"go.mongodb.org/mongo-driver/mongo/options"
)

func main() {
	var startDateStr string
	var endDateStr string
	flag.StringVar(&startDateStr, "start-date", "", "Start date for the filter (YYYY-MM-DD)")
	flag.StringVar(&endDateStr, "end-date", "", "End date for the filter (YYYY-MM-DD)")
	flag.Parse()
	
	// Parse dates from command line arguments
	startDate, err := time.Parse("2006-01-02", startDateStr)
	if err != nil {
		log.Fatal("Error parsing start date:", err)
	}

	endDate, err := time.Parse("2006-01-02", endDateStr)
	if err != nil {
		log.Fatal("Error parsing end date:", err)
	}

	// Define configuration database
	clientOptions := options.Client().ApplyURI("your-uri")
	client, err := mongo.Connect(context.Background(), clientOptions)
	if err != nil {
		log.Fatal(err)
	}
	defer client.Disconnect(context.Background())

	// Test connection
	err = client.Ping(context.Background(), nil)
	if err != nil {
		log.Fatal(err)
	}

	// Access database and collection
	database := client.Database("simpleAPI")
	collection := database.Collection("gov-history")

	// Create the filter to query
	filter := bson.M{
		"createdAt": bson.M{
			"$gte": startDate,
			"$lte": endDate,
		},
	}

	// Find result
	cur, err := collection.Find(context.Background(), filter)
	if err != nil {
		log.Fatal(err)
	}
	defer cur.Close(context.Background())

	// Create a new Excel file
	f := excelize.NewFile()

	row := 1
	var headers []string

	for cur.Next(context.Background()) {
		var result bson.M
		err := cur.Decode(&result)
		if err != nil {
			log.Fatal(err)
		}

		if len(headers) == 0 {
			// Extract headers from the first document
			for key := range result {
				headers = append(headers, key)
			}

			// Write headers to Excel
			for col, header := range headers {
				cell := toAlphaString(col) + fmt.Sprint(row)
				f.SetCellValue("Sheet1", cell, header)
			}

			row++
		}

		// Write data to Excel
		for col, header := range headers {
			cell := toAlphaString(col) + fmt.Sprint(row)
			value := result[header]
	
			// Handle nested fields using JSON marshaling
			nestedJSON, err := json.Marshal(value)
			if err != nil {
				log.Fatal(err)
			}
	
			f.SetCellValue("Sheet1", cell, string(nestedJSON))
		}

		row++
	}

	// Save the Excel file
	if err := f.SaveAs("output.xlsx"); err != nil {
		log.Fatal(err)
	}
}

func toAlphaString(index int) string {
    if index <= 0 {
        return ""
    }

    result := ""
    for index > 0 {
        index--
        result = fmt.Sprintf("%c%s", 'A'+(index%26), result)
        index /= 26
    }
    return result
}
