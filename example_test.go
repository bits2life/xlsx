package xlsx_test

import (
	"bytes"
	"fmt"
	"log"

	"go.bits2life.com/xlsx"
)

func Example() {
	wb := &xlsx.Workbook{
		Meta:   map[string]any{"version": "1.0"},
		Styles: []xlsx.Style{{Name: "header", Bold: true, BgColor: "DDDDDD"}},
		Sheets: []xlsx.Sheet{{
			Name:    "Sales",
			Columns: []xlsx.Column{{Width: 20}, {Width: 12}},
			Cells: []xlsx.Cell{
				{Cell: "A1", Value: "Product", Style: "header"},
				{Cell: "B1", Value: "Price", Style: "header"},
				{Cell: "A2", Value: "Widget"},
				{Cell: "B2", Value: 99.99},
			},
		}},
	}

	var buf bytes.Buffer
	if err := wb.Write(&buf); err != nil {
		log.Fatal(err)
	}

	doc, err := xlsx.Parse(&buf)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println(doc.Meta["version"], doc.Sheets["Sales"])
	// Output: 1.0 [[Product Price] [Widget 99.99]]
}
