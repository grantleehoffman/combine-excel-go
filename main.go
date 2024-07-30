package main

import (
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

// Function to combine Excel files
func combineExcelFiles(inputDir string, outputFile string, keywords []string, keywordRow int) error {
	files, err := os.ReadDir(inputDir)
	if err != nil {
		return err
	}

	// Create a new Excel file for the output
	outFile := xlsx.NewFile()
	outSheet, err := outFile.AddSheet("CombinedData")
	if err != nil {
		return err
	}

	// Boolean to track if header row has been written
	headerWritten := false

	for _, file := range files {
		if file.IsDir() {
			continue
		}

		filePath := inputDir + "/" + file.Name()
		fmt.Printf("Processing file: %s\n", filePath)

		// Open the Excel file
		inFile, err := xlsx.OpenFile(filePath)
		if err != nil {
			return err
		}

		for _, inSheet := range inFile.Sheets {
			// Check if keywordRow is within bounds
			if keywordRow < 1 || keywordRow > len(inSheet.Rows) {
				return fmt.Errorf("keywordRow %d is out of range", keywordRow)
			}

			// Identify columns to copy based on keywords
			keywordRowCells := inSheet.Rows[keywordRow-1].Cells
			columnsToCopy := make(map[int]bool)

			for colIndex, cell := range keywordRowCells {
				cellValue := cell.String()
				for _, keyword := range keywords {
					if strings.Contains(cellValue, keyword) {
						columnsToCopy[colIndex] = true
					}
				}
			}

			if !headerWritten {
				// Write header row
				headerRow := inSheet.Rows[keywordRow-1]
				newHeaderRow := outSheet.AddRow()
				for colIndex := range columnsToCopy {
					newCell := newHeaderRow.AddCell()
					newCell.SetValue(headerRow.Cells[colIndex].String())
				}
				headerWritten = true
			}

			// Copy relevant rows
			for i := keywordRow; i < len(inSheet.Rows); i++ {
				row := inSheet.Rows[i]

				// Skip empty rows
				if isEmptyRow(row) {
					continue
				}

				newRow := outSheet.AddRow()
				for colIndex := range columnsToCopy {
					if colIndex < len(row.Cells) {
						newCell := newRow.AddCell()
						newCell.SetValue(row.Cells[colIndex].String())
					}
				}
			}
		}
	}

	// Save the combined output file
	return outFile.Save(outputFile)
}

// Helper function to check if a row is empty
func isEmptyRow(row *xlsx.Row) bool {
	for _, cell := range row.Cells {
		if cell.String() != "" {
			return false
		}
	}
	return true
}

// Helper function to check if a value matches any of the keywords
func containsKeyword(value string, keywords []string) bool {
	for _, keyword := range keywords {
		if value == keyword {
			return true
		}
	}
	return false
}

func main() {
	inputDir := flag.String("i", "", "Directory containing input Excel files")
	outputFile := flag.String("o", "", "Path to the output Excel file")
	keywords := flag.String("k", "", "Comma-separated list of keywords to filter columns")
	keywordRow := flag.Int("r", 5, "Row number where the keyword line is located")
	flag.Parse()

	if *inputDir == "" || *outputFile == "" || *keywords == "" {
		fmt.Println("Usage: combine-excel -inputDir <inputDir> -outputFile <outputFile> -keywords <keywords>")
		os.Exit(1)
	}

	keywordList := strings.Split(*keywords, ",")
	if err := combineExcelFiles(*inputDir, *outputFile, keywordList, *keywordRow); err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}
	fmt.Println("Process complete")
}
