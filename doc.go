// Package xlsx generates and parses Excel (.xlsx) workbooks from a simple,
// JSON-friendly description.
//
// A [Workbook] lists named styles and sheets of cells. It can be built in Go
// or decoded from JSON, and written out with [Workbook.Write]:
//
//	wb := &xlsx.Workbook{
//		Styles: []xlsx.Style{{Name: "header", Bold: true}},
//		Sheets: []xlsx.Sheet{{
//			Name:  "Report",
//			Cells: []xlsx.Cell{{Cell: "A1", Value: "Hello", Style: "header"}},
//		}},
//	}
//	err := wb.Write(w)
//
// [Parse] reads a workbook back into a [Document] of plain string rows.
//
// [GenerateHandler] and [ParseHandler] expose both directions as
// [net/http] handlers for services that accept JSON and uploads.
//
// The package is a thin layer over [github.com/xuri/excelize/v2].
package xlsx
