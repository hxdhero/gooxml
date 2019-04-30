// Copyright 2017 Baliance. All rights reserved.
package main

import (
	"fmt"

	"baliance.com/gooxml/spreadsheet"
	"baliance.com/gooxml/spreadsheet/formula"
)

func main() {
	//f_if()
	//f_sum()
	f_sumif()
}

func f_sumif(){
	fmt.Println("Currently support", len(formula.SupportedFunctions()), "functions")
	fmt.Println(formula.SupportedFunctions())
	ss := spreadsheet.New()
	sheet := ss.AddSheet()
	sheet.Cell("A1").SetString("机票")
	sheet.Cell("A2").SetString("酒店")
	sheet.Cell("A3").SetString("酒店")
	sheet.Cell("A4").SetString("饮食")

	sheet.Cell("B1").SetNumber(800.00)
	sheet.Cell("B2").SetNumber(375.00)
	sheet.Cell("B3").SetNumber(375.00)
	sheet.Cell("B4").SetNumber(150.00)

	sheet.Cell("C1").SetNumber(921.58)
	sheet.Cell("C2").SetNumber(324.98)
	sheet.Cell("C3").SetNumber(324.98)
	sheet.Cell("C4").SetNumber(174.38)

	formEv := formula.NewEvaluator()
	formula := `SUMIF(B1:B4,">10",C1:C2)`
	result := formEv.Eval(sheet.FormulaContext(), formula)
	fmt.Println("result: ",result.Value())
}

func f_if() {
	fmt.Println("Currently support", len(formula.SupportedFunctions()), "functions")
	fmt.Println(formula.SupportedFunctions())
	ss := spreadsheet.New()
	sheet := ss.AddSheet()
	sheet.Cell("A1").SetString("机票")
	sheet.Cell("A2").SetString("酒店")
	sheet.Cell("A3").SetString("汽车")
	sheet.Cell("A4").SetString("饮食")

	sheet.Cell("B1").SetNumber(800.00)
	sheet.Cell("B2").SetNumber(375.00)
	sheet.Cell("B3").SetNumber(150.00)
	sheet.Cell("B4").SetNumber(150.00)

	sheet.Cell("C1").SetNumber(921.58)
	sheet.Cell("C2").SetNumber(324.98)
	sheet.Cell("C3").SetNumber(128.43)
	sheet.Cell("C4").SetNumber(174.38)

	formEv := formula.NewEvaluator()

	// the formula context allows the formula evaluator to pull data from a
	// sheet
	a1Cell := sheet.FormulaContext().Cell("A1", formEv)
	fmt.Println("A1 is", a1Cell.Value())

	// So that when evaluating formulas, live workbook data is used. Formulas
	// can be evaluated directly in the context of a sheet.
	formula := `IF(C2=324.989,"超预算","范围内")`
	result := formEv.Eval(sheet.FormulaContext(), formula)
	fmt.Println("SUM(A1:A3) is", result.Value())
}

func f_sum() {
	fmt.Println("Currently support", len(formula.SupportedFunctions()), "functions")
	fmt.Println(formula.SupportedFunctions())
	ss := spreadsheet.New()
	sheet := ss.AddSheet()
	sheet.Cell("A1").SetNumber(1.2)
	sheet.Cell("A2").SetNumber(2.3)
	sheet.Cell("A3").SetNumber(2.3)

	sheet.Cell("B1").SetNumber(1.2)
	sheet.Cell("B2").SetNumber(2.3)
	sheet.Cell("B3").SetNumber(2.3)

	formEv := formula.NewEvaluator()

	// the formula context allows the formula evaluator to pull data from a
	// sheet
	a1Cell := sheet.FormulaContext().Cell("A1", formEv)
	fmt.Println("A1 is", a1Cell.Value())

	// So that when evaluating formulas, live workbook data is used. Formulas
	// can be evaluated directly in the context of a sheet.
	formula := `SUM(A1:B3)`
	result := formEv.Eval(sheet.FormulaContext(), formula)
	fmt.Println("SUM(A1:A3) is", result.Value())

	//// Or, stored in a cell and the cell evaulated.
	//sheet.Cell("A4").SetFormulaRaw("SUM(A1:A3)+SUM(A1:A3)")
	//a4Value := formEv.Eval(sheet.FormulaContext(), "A4")
	//fmt.Println("A4 is", a4Value.Value())
}
