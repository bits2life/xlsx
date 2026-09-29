package xlsx

import (
	"io"
	"slices"

	"github.com/xuri/excelize/v2"
)

// Document is the content of a parsed workbook.
type Document struct {
	// Meta holds the metadata written from Workbook.Meta, if any.
	Meta map[string]any `json:"meta,omitempty"`
	// Sheets maps each sheet name to its rows of formatted cell values.
	Sheets map[string][][]string `json:"sheets"`
	// Order lists the sheet names in workbook order.
	Order []string `json:"order"`
}

// Parse reads an .xlsx workbook from r. The metadata sheet is returned as
// Document.Meta rather than as a sheet.
func Parse(r io.Reader) (*Document, error) {
	f, err := excelize.OpenReader(r)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	doc := &Document{Sheets: make(map[string][][]string)}
	names := f.GetSheetList()
	if slices.Contains(names, MetaSheetName) {
		if doc.Meta, err = readMeta(f); err != nil {
			return nil, err
		}
	}
	for _, name := range names {
		if name == MetaSheetName {
			continue
		}
		rows, err := f.GetRows(name)
		if err != nil {
			return nil, err
		}
		if rows == nil {
			rows = [][]string{}
		}
		doc.Sheets[name] = rows
		doc.Order = append(doc.Order, name)
	}
	return doc, nil
}
