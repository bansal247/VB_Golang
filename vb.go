package main

import (
	"fmt"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

var (
	colCarrier          = "B"  // "Carrier"
	colAgentName        = "A"  // "Agent Name"
	colAgentID          = "C"  // "Agent ID"
	colStatementDate    = "E"  // "Statement Date"
	colClientFullName   = "G"  // "Client Full Name"
	colCarrierMemberID  = "K"  // "Carrier Member ID"
	colPolicyNumber     = "L"  // "Policy Number"
	colEffectiveDate    = "M"  // "Effective Date"
	colLine             = "R"  // "Line"
	colSubLine          = "S"  // "Sub-line"
	colPlanType         = "T"  // "Plan Type"
	colContract         = "V"  // "Contract"
	colPremium          = "Z"  // "Premium"
	colAgentSplit       = "AA" // "Agent Split"
	colCompRate         = "AB" // "Comp Rate"
	colCommission       = "AD" // "Commission"
	colCommissionAction = "AG" // "Commission Action"
	colStatementLink    = "AH" // "Statement link"
	colClientFirstName  = "H"  //Client First Name
	colClientMiddleName = "I"  //Client Middle Name/Initial
	colClientLastName   = "J"  //Client Last name
)

var headers = map[string]string{
	colAgentName:        "Agent Name",
	colCarrier:          "Carrier",
	colAgentID:          "Agent ID",
	colStatementDate:    "Statement Date",
	colClientFullName:   "Client Full Name",
	colCarrierMemberID:  "Carrier Member ID",
	colPolicyNumber:     "Policy Number",
	colEffectiveDate:    "Effective Date",
	colLine:             "Line",
	colSubLine:          "Sub-line",
	colPlanType:         "Plan Type",
	colContract:         "Contract",
	colPremium:          "Premium",
	colAgentSplit:       "Agent Split",
	colCompRate:         "Comp Rate",
	colCommission:       "Commission",
	colCommissionAction: "Commission Action",
	colStatementLink:    "Statement Link",
	colClientFirstName:  "Client First Name",
	colClientMiddleName: "Client Middle Name/Initial",
	colClientLastName:   "Client Last name",
}

func main() {
	filePath := "files/Humana ay64 big file (1).xlsx" // Change to your actual file path
	err := processHumanaWorkbook(filePath)
	if err != nil {
		fmt.Printf("error: %v\n", err)
	}
}

func processHumanaWorkbook(filePath string) error {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Errorf("processHumanaWorkbook: failed to open file: %w", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Printf("warning: failed to close file: %v\n", err)
		}
	}()

	sourceSheetName := f.GetSheetName(0)
	if sourceSheetName == "" {
		return fmt.Errorf("processHumanaWorkbook: failed to get first sheet name")
	}

	if err := restructureFileHumana(f, sourceSheetName); err != nil {
		return fmt.Errorf("processHumanaWorkbook: restructureFileHumana failed: %w", err)
	}

	if err := amendColumnsHumana(f, sourceSheetName); err != nil {
		return fmt.Errorf("processHumanaWorkbook: amendColumnsHumana failed: %w", err)
	}

	dataSheetName := "Data_Sh" // since not given in VBA
	if index, _ := f.GetSheetIndex(dataSheetName); index == -1 {
		f.NewSheet(dataSheetName)
	}
	for col, header := range headers {
		cell := fmt.Sprintf("%s1", col)
		if err := f.SetCellValue(dataSheetName, cell, header); err != nil {
			return fmt.Errorf("processHumanaWorkbook: failed to set header %s at %s: %w", header, cell, err)
		}
	}
	nextRow := getLastRow(f, dataSheetName) + 1

	sourceRows, err := f.GetRows(sourceSheetName)
	if err != nil {
		return fmt.Errorf("processHumanaWorkbook: failed to get rows from source sheet: %w", err)
	}
	lastSourceRow := len(sourceRows)

	if lastSourceRow <= 1 {
		return nil // No data to copy
	}

	for i := 2; i <= lastSourceRow; i++ {
		rowOffset := nextRow + i - 2

		copyCell := func(srcCol, dstCol string) error {
			val, err := f.GetCellValue(sourceSheetName, fmt.Sprintf("%s%d", srcCol, i))
			if err != nil {
				return fmt.Errorf("copyCell: failed to get value from %s%d: %w", srcCol, i, err)
			}
			if err := f.SetCellValue(dataSheetName, fmt.Sprintf("%s%d", dstCol, rowOffset), val); err != nil {
				return fmt.Errorf("copyCell: failed to set value to %s%d: %w", dstCol, rowOffset, err)
			}
			return nil
		}

		if err := f.SetCellValue(dataSheetName, fmt.Sprintf("%s%d", colLine, rowOffset), "Health"); err != nil {
			return fmt.Errorf("processHumanaWorkbook: failed to set Line at row %d: %w", rowOffset, err)
		}

		columnPairs := map[string]string{
			"AU": colCarrier,
			"C":  colAgentName,
			"D":  colAgentID,
			"B":  colStatementDate,
			"E":  colClientFullName,
			"F":  colCarrierMemberID,
			"AM": colPolicyNumber,
			"AF": colEffectiveDate,
			"AV": colSubLine,
			"T":  colPlanType,
			"AN": colContract,
			"W":  colPremium,
			"X":  colAgentSplit,
			"V":  colCompRate,
			"Y":  colCommission,
			"AB": colCommissionAction,
		}

		for src, dst := range columnPairs {
			if err := copyCell(src, dst); err != nil {
				return fmt.Errorf("processHumanaWorkbook: copyCell %s -> %s failed at row %d: %w", src, dst, i, err)
			}
		}

		formula := fmt.Sprintf(`=HYPERLINK("%s", "%s")`, filePath, filepath.Base(filePath))
		if err := f.SetCellFormula(dataSheetName, fmt.Sprintf("%s%d", colStatementLink, rowOffset), formula); err != nil {
			return fmt.Errorf("processHumanaWorkbook: failed to set hyperlink formula at row %d: %w", rowOffset, err)
		}
	}

	if err := splitClientName3(f, dataSheetName, nextRow, nextRow+lastSourceRow-2); err != nil {
		return fmt.Errorf("processHumanaWorkbook: failed to set split client name %w", err)
	}

	if err := formatAgentName1(f, dataSheetName, nextRow, nextRow+lastSourceRow-2); err != nil {
		return fmt.Errorf("processHumanaWorkbook: failed to format agent name %w", err)
	}

	newFilePath := strings.Replace(filePath, ".xlsx", "_processed.xlsx", 1)
	if err := f.SaveAs(newFilePath); err != nil {
		return fmt.Errorf("processHumanaWorkbook: failed to save new file: %w", err)
	}

	fmt.Println("Processing done and file saved successfully!")
	return nil
}

func formatAgentName1(f *excelize.File, dataSheetName string, startRow, endRow int) error {
	for i := startRow; i <= endRow; i++ {
		fullName, err := f.GetCellValue(dataSheetName, fmt.Sprintf("%s%d", colAgentName, i))
		if err != nil {
			return fmt.Errorf("failed to get full name at row %d: %w", i, err)
		}

		properFullName := strings.Title(strings.ToLower(fullName))

		// Split the full name into components
		nameParts := strings.Fields(properFullName)

		// Determine first, middle, and last names
		var firstName, lastName string

		// First Name
		if len(nameParts) > 0 {
			firstName = nameParts[1]
		}

		// Last Name
		if len(nameParts) >= 2 {
			lastName = nameParts[0]
		}

		// Swap First and Last Name (i.e., "Albert Mikesh" from "MIKESH ALBERT")
		swappedFullName := fmt.Sprintf("%s %s", firstName, lastName)

		if err := f.SetCellValue(dataSheetName, fmt.Sprintf("%s%d", colAgentName, i), swappedFullName); err != nil {
			return fmt.Errorf("failed to set swapped full name at row %d: %w", i, err)
		}
	}
	return nil
}

func splitClientName3(f *excelize.File, dataSheetName string, startRow, endRow int) error {
	for i := startRow; i <= endRow; i++ {
		fullName, err := f.GetCellValue(dataSheetName, fmt.Sprintf("%s%d", colClientFullName, i))
		if err != nil {
			return fmt.Errorf("failed to get full name at row %d: %w", i, err)
		}

		properFullName := strings.Title(strings.ToLower(fullName))

		// Split the full name into components
		nameParts := strings.Fields(properFullName)

		// Determine first, middle, and last names
		var firstName, middleName, lastName string

		// First Name
		if len(nameParts) > 0 {
			firstName = nameParts[1]
		}

		// Middle Name or Initial
		if len(nameParts) == 3 {
			middleName = nameParts[2]
		}

		// Last Name
		if len(nameParts) >= 2 {
			lastName = nameParts[0]
		}

		// Swap First and Last Name (i.e., "Albert Mikesh" from "MIKESH ALBERT")
		swappedFullName := fmt.Sprintf("%s %s", firstName, lastName)

		// Set the reversed name back in the clientFullNameCol if required
		if err := f.SetCellValue(dataSheetName, fmt.Sprintf("%s%d", colClientFullName, i), swappedFullName); err != nil {
			return fmt.Errorf("failed to set swapped full name at row %d: %w", i, err)
		}

		// Set First Name
		if err := f.SetCellValue(dataSheetName, fmt.Sprintf("%s%d", colClientFirstName, i), firstName); err != nil {
			return fmt.Errorf("failed to set first name at row %d: %w", i, err)
		}

		// Set Middle Name or Initial (empty if not available)
		if err := f.SetCellValue(dataSheetName, fmt.Sprintf("%s%d", colClientMiddleName, i), middleName); err != nil {
			return fmt.Errorf("failed to set middle name/initial at row %d: %w", i, err)
		}

		// Set Last Name
		if err := f.SetCellValue(dataSheetName, fmt.Sprintf("%s%d", colClientLastName, i), lastName); err != nil {
			return fmt.Errorf("failed to set last name at row %d: %w", i, err)
		}
	}
	return nil
}

func getLastRow(f *excelize.File, sheet string) int {
	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Printf("getLastRow: warning: failed to get rows: %v\n", err)
		return 0
	}
	return len(rows)
}

func restructureFileHumana(f *excelize.File, sheetName string) error {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("restructureFileHumana: failed to read rows: %w", err)
	}

	for i := len(rows); i >= 2; i-- {
		cell := fmt.Sprintf("C%d", i)
		val, err := f.GetCellValue(sheetName, cell)
		if err != nil {
			return fmt.Errorf("restructureFileHumana: failed to get cell value at %s: %w", cell, err)
		}

		if strings.TrimSpace(val) == "" {
			if err := f.RemoveRow(sheetName, i); err != nil {
				return fmt.Errorf("restructureFileHumana: failed to remove row %d: %w", i, err)
			}
		}
	}
	return nil
}

func amendColumnsHumana(f *excelize.File, sheetName string) error {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("amendColumnsHumana: failed to get rows: %w", err)
	}

	for i := 1; i < len(rows); i++ {
		rowNum := i + 1

		cellX := fmt.Sprintf("X%d", rowNum)
		valueX, err := f.GetCellValue(sheetName, cellX)
		if err == nil && valueX != "" {
			if num, err := parseFloat(valueX); err == nil {
				if err := f.SetCellValue(sheetName, cellX, num/100); err != nil {
					return fmt.Errorf("amendColumnsHumana: failed to set adjusted value in X%d: %w", rowNum, err)
				}
			}
		}

		product := strings.TrimSpace(strings.ToLower(getCellValueSafe(f, sheetName, fmt.Sprintf("S%d", rowNum))))
		blkBusCd := strings.TrimSpace(strings.ToLower(getCellValueSafe(f, sheetName, fmt.Sprintf("J%d", rowNum))))
		cellT := strings.TrimSpace(strings.ToLower(getCellValueSafe(f, sheetName, fmt.Sprintf("T%d", rowNum))))
		cellAB := strings.TrimSpace(getCellValueSafe(f, sheetName, fmt.Sprintf("AB%d", rowNum)))

		carrier := ""
		subline := ""

		switch {
		case product == "dental":
			carrier = "Humana Dental"
			subline = "Dental"
		case product == "vision":
			carrier = "Humana Vision"
			subline = "Vision"
		case blkBusCd == "ms":
			carrier = "Humana Med Supp"
			subline = "Med Supp"
		case blkBusCd == "ma" && cellT == "pdp":
			carrier = "Humana PDP"
			subline = "PDP"
		case blkBusCd == "pdp":
			carrier = "Humana PDP"
			subline = "PDP"
		case blkBusCd == "ma":
			carrier = "Humana MAPD"
			subline = "Med Adv"
		default:
			carrier = "Humana"
		}

		if strings.Contains(strings.ToUpper(cellAB), "OVERRIDE") {
			carrier += " override"
		}

		if err := f.SetCellValue(sheetName, fmt.Sprintf("AU%d", rowNum), carrier); err != nil {
			return fmt.Errorf("amendColumnsHumana: failed to set Carrier AU%d: %w", rowNum, err)
		}
		if subline != "" {
			if err := f.SetCellValue(sheetName, fmt.Sprintf("AV%d", rowNum), subline); err != nil {
				return fmt.Errorf("amendColumnsHumana: failed to set Sub-line AV%d: %w", rowNum, err)
			}
		}
	}

	return nil
}

func parseFloat(s string) (float64, error) {
	return strconv.ParseFloat(s, 64)
}

func getCellValueSafe(f *excelize.File, sheet, cell string) string {
	val, err := f.GetCellValue(sheet, cell)
	if err != nil {
		fmt.Printf("getCellValueSafe: warning: failed to get value from %s: %v\n", cell, err)
		return ""
	}
	return val
}
