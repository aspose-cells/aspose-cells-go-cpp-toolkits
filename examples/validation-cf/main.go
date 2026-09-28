// Command validation-cf demonstrates the editor DSL's data validation and
// conditional formatting support. It shows how to add validation rules to
// control user input and apply conditional formatting to visualize data.
//
// Nothing here imports the underlying engine binding. The workbook is seeded
// with datasource.NewEmptyWorkbook and every operation goes through editor
// and query, which is the point of the toolkit: a caller builds validations
// and formatting without ever holding an engine handle.
//
// Output lands in examples/validation-cf/out/.
package main

import (
	"fmt"
	"log"
	"os"

	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/datasource"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/editor"
	examples "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/examples/common"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
)

func main() {
	if err := examples.SetLicense(); err != nil {
		log.Printf("license: %v", err)
	}

	// Seed an empty workbook with sample data
	seed, err := datasource.NewEmptyWorkbook()
	if err != nil {
		log.Fatal(err)
	}

	const sheet = "Data"

	// Build a table with sample data for validation and formatting
	seeded, err := editor.EditSpreadsheet(
		seed,
		editor.WithAddWorksheet(sheet),
		editor.InWorksheet(sheet,
			// Headers
			editor.SetCellValue(0, 0, "Name"),
			editor.SetCellValue(0, 1, "Age"),
			editor.SetCellValue(0, 2, "Score"),
			editor.SetCellValue(0, 3, "Status"),
			editor.SetCellValue(0, 4, "Email"),
			// Data rows
			editor.SetCellValue(1, 0, "Alice"),
			editor.SetCellValue(1, 1, int32(30)),
			editor.SetCellValue(1, 2, int32(95)),
			editor.SetCellValue(1, 3, "Active"),
			editor.SetCellValue(1, 4, "alice@example.com"),
			editor.SetCellValue(2, 0, "Bob"),
			editor.SetCellValue(2, 1, int32(25)),
			editor.SetCellValue(2, 2, int32(85)),
			editor.SetCellValue(2, 3, "Active"),
			editor.SetCellValue(2, 4, "bob@example.com"),
			editor.SetCellValue(3, 0, "Charlie"),
			editor.SetCellValue(3, 1, int32(35)),
			editor.SetCellValue(3, 2, int32(72)),
			editor.SetCellValue(3, 3, "Inactive"),
			editor.SetCellValue(3, 4, "charlie@example.com"),
			editor.SetCellValue(4, 0, "Diana"),
			editor.SetCellValue(4, 1, int32(28)),
			editor.SetCellValue(4, 2, int32(88)),
			editor.SetCellValue(4, 3, "Pending"),
			editor.SetCellValue(4, 4, "diana@example.com"),
		),
	)
	if err != nil {
		log.Fatalf("seed data: %v", err)
	}

	// Example 1: Data Validation
	// Add various validation rules to control user input
	withValidation, err := editor.EditSpreadsheet(
		datasource.BytesSource(seeded),
		editor.InWorksheet(sheet,
			// Whole number validation for Age (18-100)
			editor.AddDataValidation("B2:B100",
				editor.WithValidationType(editor.ValidationTypeWholeNumber),
				editor.WithValidationOperator(editor.OperatorTypeBetween),
				editor.WithValidationFormula1("18"),
				editor.WithValidationFormula2("100"),
				editor.WithValidationErrorMessage("Age must be between 18 and 100"),
				editor.WithValidationErrorTitle("Invalid Age"),
				editor.WithValidationInputTitle("Enter Age"),
				editor.WithValidationInputMessage("Please enter an age between 18 and 100"),
				editor.WithValidationShowInput(true),
				editor.WithValidationShowError(true),
			),
			// List validation for Status
			editor.AddDataValidation("D2:D100",
				editor.WithValidationList([]string{"Active", "Inactive", "Pending"}),
				editor.WithValidationInCellDropDown(true),
				editor.WithValidationInputTitle("Select Status"),
				editor.WithValidationInputMessage("Choose from the dropdown list"),
			),
			// Decimal validation for Score (0-100)
			editor.AddDataValidation("C2:C100",
				editor.WithValidationType(editor.ValidationTypeDecimal),
				editor.WithValidationOperator(editor.OperatorTypeBetween),
				editor.WithValidationFormula1("0"),
				editor.WithValidationFormula2("100"),
				editor.WithValidationErrorMessage("Score must be between 0 and 100"),
			),
		),
	)
	if err != nil {
		log.Fatalf("add validation: %v", err)
	}
	writeOut("validation.xlsx", withValidation)

	// Example 2: Conditional Formatting - Color Scale
	// Apply a 3-color scale to the Score column (red-yellow-green)
	withColorScale, err := editor.EditSpreadsheet(
		datasource.BytesSource(withValidation),
		editor.AddConditionalFormatting(sheet, "C2:C100",
			editor.WithColorScale("#F8696B", "#63BE7B", "#FFEB84"),
		),
	)
	if err != nil {
		log.Fatalf("add color scale: %v", err)
	}
	writeOut("color-scale.xlsx", withColorScale)

	// Example 3: Conditional Formatting - Data Bar
	// Add data bars to the Score column
	withDataBar, err := editor.EditSpreadsheet(
		datasource.BytesSource(withValidation),
		editor.AddConditionalFormatting(sheet, "C2:C100",
			editor.WithDataBar("#63BE7B"),
		),
	)
	if err != nil {
		log.Fatalf("add data bar: %v", err)
	}
	writeOut("data-bar.xlsx", withDataBar)

	// Example 4: Conditional Formatting - Icon Set
	// Add traffic light icons based on score
	withIconSet, err := editor.EditSpreadsheet(
		datasource.BytesSource(withValidation),
		editor.AddConditionalFormatting(sheet, "C2:C100",
			editor.WithIconSet(editor.IconSetTrafficLights31),
		),
	)
	if err != nil {
		log.Fatalf("add icon set: %v", err)
	}
	writeOut("icon-set.xlsx", withIconSet)

	// Example 5: Conditional Formatting - Cell Value Rule
	// Highlight scores above 90 in green
	withCellValueRule, err := editor.EditSpreadsheet(
		datasource.BytesSource(withValidation),
		editor.AddConditionalFormatting(sheet, "C2:C100",
			editor.WithCellValueRule(editor.OperatorTypeGreaterThan, "90", "",
				editor.WithFontColor("#006100"),
				editor.WithBackgroundColor("#C6EFCE"),
				editor.WithFontIsBold(true),
			),
		),
	)
	if err != nil {
		log.Fatalf("add cell value rule: %v", err)
	}
	writeOut("cell-value-rule.xlsx", withCellValueRule)

	// Example 6: Conditional Formatting - Expression Rule
	// Highlight rows where age is greater than 30
	withExpressionRule, err := editor.EditSpreadsheet(
		datasource.BytesSource(withValidation),
		editor.AddConditionalFormatting(sheet, "A2:E100",
			editor.WithExpressionRule("=$B2>30",
				editor.WithBackgroundColor("#FFEB9C"),
			),
		),
	)
	if err != nil {
		log.Fatalf("add expression rule: %v", err)
	}
	writeOut("expression-rule.xlsx", withExpressionRule)

	// Example 7: Conditional Formatting - Top 10
	// Highlight top 3 scores
	withTop10, err := editor.EditSpreadsheet(
		datasource.BytesSource(withValidation),
		editor.AddConditionalFormatting(sheet, "C2:C100",
			editor.WithTop10Rule(3, true,
				editor.WithFontColor("#9C5700"),
				editor.WithBackgroundColor("#FFEB9C"),
			),
		),
	)
	if err != nil {
		log.Fatalf("add top 10 rule: %v", err)
	}
	writeOut("top10-rule.xlsx", withTop10)

	// Example 8: Query validations and conditional formatting
	fmt.Println("\n=== Querying Validations ===")
	count, err := query.ValidationCount(
		datasource.BytesSource(withValidation),
		query.WithSheet(sheet),
	)
	if err != nil {
		log.Fatalf("query validation count: %v", err)
	}
	fmt.Printf("Total validations: %d\n", count)

	validations, err := query.AllValidations(
		datasource.BytesSource(withValidation),
		query.WithSheet(sheet),
	)
	if err != nil {
		log.Fatalf("query all validations: %v", err)
	}
	for _, v := range validations {
		fmt.Printf("  Validation %d: Type=%s, Operator=%s, Formula1=%s, Formula2=%s\n",
			v.Index, v.Type, v.Operator, v.Formula1, v.Formula2)
		if v.ErrorMessage != "" {
			fmt.Printf("    Error Message: %s\n", v.ErrorMessage)
		}
	}

	fmt.Println("\n=== Querying Conditional Formatting ===")
	cfCount, err := query.ConditionalFormattingCount(
		datasource.BytesSource(withColorScale),
		query.WithSheet(sheet),
	)
	if err != nil {
		log.Fatalf("query conditional formatting count: %v", err)
	}
	fmt.Printf("Total conditional formattings: %d\n", cfCount)

	cfList, err := query.AllConditionalFormattings(
		datasource.BytesSource(withColorScale),
		query.WithSheet(sheet),
	)
	if err != nil {
		log.Fatalf("query all conditional formattings: %v", err)
	}
	for _, cf := range cfList {
		fmt.Printf("  Conditional Formatting %d:\n", cf.Index)
		fmt.Printf("    Areas: %v\n", cf.Areas)
		for _, cond := range cf.Conditions {
			fmt.Printf("    Condition %d: Type=%s, Operator=%s\n",
				cond.Index, cond.Type, cond.Operator)
		}
	}

	log.Println("validation and conditional formatting workbooks written to examples/validation-cf/out/")
}

func writeOut(name string, data []byte) {
	if err := os.MkdirAll(examples.OutDir("validation-cf"), 0o755); err != nil {
		log.Fatal(err)
	}
	if err := os.WriteFile(examples.OutPath("validation-cf", name), data, 0o644); err != nil {
		log.Fatal(err)
	}
	fmt.Printf("wrote %s\n", examples.OutPath("validation-cf", name))
}
