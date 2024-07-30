package main

import (
	"flag"
	"fmt"
	"os"
	"regexp"
	"runtime"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Function to combine Excel files
func combineExcelFiles(inputDir string, outputFile string, keywords []string, keywordRow int) error {
	files, err := os.ReadDir(inputDir)
	if err != nil {
		return err
	}

	// Create a new Excel file for the output
	outFile := excelize.NewFile()
	defaultSheetName := "Sheet1" // Default sheet name

	// Get the default sheet index
	outSheetIndex, err := outFile.GetSheetIndex(defaultSheetName)
	if err != nil || outSheetIndex == -1 {
		return fmt.Errorf("default sheet %s not found", defaultSheetName)
	}

	// Map to hold the column indices in order based on keywords
	headerMap := make(map[int]string) // column index to header name
	keywordOrder := []int{}           // order of columns to be copied
	dataRows := [][]string{}          // rows to be copied

	for _, file := range files {
		if file.IsDir() {
			continue
		}

		filePath := inputDir + "/" + file.Name()
		fmt.Printf("Processing file: %s\n", filePath)

		// Open the Excel file
		inFile, err := excelize.OpenFile(filePath)
		if err != nil {
			return err
		}

		sheetName := inFile.GetSheetList()[0]
		// Get the rows from the input sheet
		rows, err := inFile.GetRows(sheetName)
		if err != nil {
			return err
		}

		if keywordRow < 1 || keywordRow > len(rows) {
			return fmt.Errorf("keywordRow %d is out of range", keywordRow)
		}

		// Identify columns to copy based on keywords
		keywordRowCells := rows[keywordRow-1]
		columnsToCopy := make(map[int]bool)

		for colIndex, cellValue := range keywordRowCells {
			for _, keyword := range keywords {
				if containsKeyword(cellValue, keyword) {
					normalizedCellValue := normalizeWhitespace(cellValue)
					columnsToCopy[colIndex] = true
					if _, exists := headerMap[colIndex]; !exists {
						headerMap[colIndex] = normalizedCellValue
						keywordOrder = append(keywordOrder, colIndex)
					}
				}
			}
		}

		// Copy relevant rows
		for i := keywordRow; i < len(rows); i++ {
			row := rows[i]

			// Skip empty rows
			if isEmptyRow(row) {
				continue
			}

			dataRow := make([]string, len(keywordOrder))
			for j, colIndex := range keywordOrder {
				if colIndex < len(row) {
					dataRow[j] = row[colIndex]
				}
			}
			dataRows = append(dataRows, dataRow)
		}
	}

	if len(headerMap) == 0 {
		return fmt.Errorf("no columns match the specified keywords")
	}

	// Write header row
	headerRow := make([]string, len(keywordOrder))
	for i, colIndex := range keywordOrder {
		headerRow[i] = headerMap[colIndex]
	}
	for i, value := range headerRow {
		colName, _ := excelize.ColumnNumberToName(i + 1) // Adjust index for 1-based column numbering
		cellAddress := fmt.Sprintf("%s%d", colName, 1)   // Row 1 for header
		outFile.SetCellValue(defaultSheetName, cellAddress, value)
	}

	// Write data rows
	for rowIndex, row := range dataRows {
		for colIndex, value := range row {
			colName, _ := excelize.ColumnNumberToName(colIndex + 1) // Adjust index for 1-based column numbering
			cellAddress := fmt.Sprintf("%s%d", colName, rowIndex+2) // Row numbers start from 2
			outFile.SetCellValue(defaultSheetName, cellAddress, value)
		}
	}

	// Save the combined output file
	return outFile.SaveAs(outputFile)
}

// Helper function to check if a row is empty
func isEmptyRow(row []string) bool {
	for _, cell := range row {
		if cell != "" {
			return false
		}
	}
	return true
}

// Helper function to normalize whitespace (replace tabs and newlines with a single space)
func normalizeWhitespace(s string) string {
	// Replace all types of whitespace (including tabs and newlines) with a single space
	re := regexp.MustCompile(`\s+`)
	return re.ReplaceAllString(s, " ")
}

// Helper function to check if a value matches any of the keywords
func containsKeyword(value string, keyword string) bool {
	normalizedValue := normalizeWhitespace(value)
	normalizedKeyword := normalizeWhitespace(keyword)
	return strings.Contains(normalizedValue, normalizedKeyword)
}

func main() {
	inputDir := flag.String("i", "", "Directory containing input Excel files")
	outputFile := flag.String("o", "", "Path to the output Excel file")
	keywords := flag.String("k", "", "Comma-separated list of keywords to filter columns")
	keywordRow := flag.Int("r", 5, "Row number where the keyword line is located")
	flag.Parse()

	binaryName := "combine-excel"
	if runtime.GOOS == "windows" {
		binaryName += ".exe"
	}

	if *inputDir == "" || *outputFile == "" || *keywords == "" {
		fmt.Printf("Usage: %s -i <inputDir> -o <outputFile.xlsx> -k <keywords> -r <row_number>\n", binaryName)
		os.Exit(1)
	}

	keywordList := strings.Split(*keywords, ",")
	if err := combineExcelFiles(*inputDir, *outputFile, keywordList, *keywordRow); err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}
	fmt.Println("Process complete")
}
