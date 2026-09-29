package xlsx

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Workbook describes a workbook to generate: optional metadata, a set of
// named styles and the sheets themselves.
type Workbook struct {
	// Meta is stored as JSON-encoded key/value pairs in a very hidden sheet
	// and returned again by Parse.
	Meta   map[string]any `json:"meta,omitempty"`
	Styles []Style        `json:"styles"`
	Sheets []Sheet        `json:"sheets"`
}

// Sheet describes one worksheet.
type Sheet struct {
	Name        string       `json:"name"`
	Cells       []Cell       `json:"cells"`
	Columns     []Column     `json:"columns,omitempty"`     // Indexed by position: 0 is column A
	Protection  string       `json:"protection,omitempty"`  // Password for sheet protection
	Validations []Validation `json:"validations,omitempty"` // Data validation rules
}

// Column holds settings for a single column.
type Column struct {
	Width  float64 `json:"width"`
	Style  string  `json:"style,omitempty"`  // Name of a Style
	Hidden bool    `json:"hidden,omitempty"` // Whether the column is hidden
}

// Cell is a value placed at an A1-style reference.
//
// String values starting with "=" are written as formulas, and strings in
// YYYY-MM-DD form as dates. Any other string is written as text. Numbers and
// booleans keep their type, nil writes an empty cell, and anything else is
// formatted with fmt.
type Cell struct {
	Cell    string `json:"cell"`
	Value   any    `json:"value"`
	Style   string `json:"style,omitempty"`   // Name of a Style
	Comment string `json:"comment,omitempty"` // Optional cell comment
}

// Write generates the workbook and writes it to w in .xlsx format.
func (wb *Workbook) Write(w io.Writer) error {
	f, err := wb.build()
	if err != nil {
		return err
	}
	defer f.Close()
	return f.Write(w)
}

// Bytes generates the workbook and returns it in .xlsx format.
func (wb *Workbook) Bytes() ([]byte, error) {
	var buf bytes.Buffer
	if err := wb.Write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// Generate decodes a JSON Workbook from r and writes the generated .xlsx to w.
func Generate(r io.Reader, w io.Writer) error {
	var wb Workbook
	if err := json.NewDecoder(r).Decode(&wb); err != nil {
		return fmt.Errorf("xlsx: decoding workbook: %w", err)
	}
	return wb.Write(w)
}

func (wb *Workbook) build() (*excelize.File, error) {
	f := excelize.NewFile()
	ok := false
	defer func() {
		if !ok {
			f.Close()
		}
	}()

	styles, err := newStyleSet(f, wb.Styles)
	if err != nil {
		return nil, err
	}

	for i, sheet := range wb.Sheets {
		name := sheet.Name
		if name == "" {
			name = fmt.Sprintf("Sheet%d", i+1)
		}
		if i == 0 {
			err = f.SetSheetName("Sheet1", name)
		} else {
			_, err = f.NewSheet(name)
		}
		if err != nil {
			return nil, fmt.Errorf("xlsx: sheet %q: %w", name, err)
		}
		if err := writeSheet(f, name, sheet, styles); err != nil {
			return nil, fmt.Errorf("xlsx: sheet %q: %w", name, err)
		}
	}

	if len(wb.Meta) > 0 {
		if err := writeMeta(f, wb.Meta); err != nil {
			return nil, err
		}
	}

	f.SetActiveSheet(0)
	ok = true
	return f, nil
}

func writeSheet(f *excelize.File, name string, sheet Sheet, styles *styleSet) error {
	for i, col := range sheet.Columns {
		colName, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			return err
		}
		if col.Width > 0 {
			if err := f.SetColWidth(name, colName, colName, col.Width); err != nil {
				return fmt.Errorf("column %s: %w", colName, err)
			}
		}
		if col.Hidden {
			if err := f.SetColVisible(name, colName, false); err != nil {
				return fmt.Errorf("column %s: %w", colName, err)
			}
		}
		if col.Style != "" {
			id, err := styles.id(col.Style)
			if err != nil {
				return fmt.Errorf("column %s: %w", colName, err)
			}
			if err := f.SetColStyle(name, colName, id); err != nil {
				return fmt.Errorf("column %s: %w", colName, err)
			}
		}
	}

	for _, cell := range sheet.Cells {
		if err := setCellValue(f, name, cell); err != nil {
			return fmt.Errorf("cell %s: %w", cell.Cell, err)
		}
		if cell.Style != "" {
			id, err := styles.id(cell.Style)
			if err != nil {
				return fmt.Errorf("cell %s: %w", cell.Cell, err)
			}
			if err := f.SetCellStyle(name, cell.Cell, cell.Cell, id); err != nil {
				return fmt.Errorf("cell %s: %w", cell.Cell, err)
			}
		}
	}

	for _, v := range sheet.Validations {
		if err := applyValidation(f, name, v); err != nil {
			return fmt.Errorf("validation %s: %w", v.Range, err)
		}
	}

	if sheet.Protection != "" {
		err := f.ProtectSheet(name, &excelize.SheetProtectionOptions{
			Password:            sheet.Protection,
			SelectLockedCells:   true,
			SelectUnlockedCells: true,
			AlgorithmName:       "SHA-512",
		})
		if err != nil {
			return fmt.Errorf("protection: %w", err)
		}
	}
	return nil
}

func setCellValue(f *excelize.File, sheet string, cell Cell) error {
	var err error
	switch v := cell.Value.(type) {
	case string:
		if strings.HasPrefix(v, "=") {
			// OOXML stores formulas without the leading "=".
			err = f.SetCellFormula(sheet, cell.Cell, v[1:])
		} else if date, perr := parseDate(v); perr == nil {
			err = f.SetCellValue(sheet, cell.Cell, date)
		} else {
			// Write as text so Excel doesn't reinterpret it.
			err = f.SetCellStr(sheet, cell.Cell, v)
		}
	case float64, float32, int, int8, int16, int32, int64,
		uint, uint8, uint16, uint32, uint64, json.Number:
		if n, ok := v.(json.Number); ok {
			var fv float64
			if fv, err = n.Float64(); err == nil {
				err = f.SetCellValue(sheet, cell.Cell, fv)
			}
		} else {
			err = f.SetCellValue(sheet, cell.Cell, v)
		}
	case bool:
		err = f.SetCellBool(sheet, cell.Cell, v)
	case time.Time:
		err = f.SetCellValue(sheet, cell.Cell, v)
	case nil:
		err = f.SetCellStr(sheet, cell.Cell, "")
	default:
		err = f.SetCellStr(sheet, cell.Cell, fmt.Sprintf("%v", v))
	}
	if err != nil {
		return err
	}

	if cell.Comment != "" {
		return f.AddComment(sheet, excelize.Comment{
			Cell:   cell.Cell,
			Text:   cell.Comment,
			Width:  320,
			Height: 72,
		})
	}
	return nil
}

// parseDate accepts dates in YYYY-MM-DD form only.
func parseDate(s string) (time.Time, error) {
	if len(s) != len("2006-01-02") {
		return time.Time{}, fmt.Errorf("not a date: %q", s)
	}
	return time.Parse("2006-01-02", s)
}
