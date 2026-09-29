# xlsx

[![Go Reference](https://pkg.go.dev/badge/go.bits2life.com/xlsx.svg)](https://pkg.go.dev/go.bits2life.com/xlsx)

Generate and parse Excel (`.xlsx`) workbooks from a small, JSON-friendly
description. A thin layer over [excelize](https://github.com/xuri/excelize).

```sh
go get go.bits2life.com/xlsx
```

## Usage

Build a workbook in Go:

```go
wb := &xlsx.Workbook{
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
err := wb.Write(w) // or wb.Bytes()
```

Or go straight from JSON (see [JSON format](#json-format)):

```go
err := xlsx.Generate(jsonReader, xlsxWriter)
```

Read a workbook back:

```go
doc, err := xlsx.Parse(r)
// doc.Meta   map[string]any         metadata written from Workbook.Meta
// doc.Sheets map[string][][]string  rows of formatted cell values per sheet
// doc.Order  []string               sheet names in workbook order
```

### HTTP handlers

`GenerateHandler` and `ParseHandler` are plain `net/http` handler functions:

```go
mux := http.NewServeMux()
mux.HandleFunc("POST /xlsx", xlsx.GenerateHandler)       // JSON in, .xlsx attachment out
mux.HandleFunc("POST /xlsx/parse", xlsx.ParseHandler)    // multipart field "file" in, JSON out
```

Errors are returned as `{"error": "..."}` with status 400. Uploads to
`ParseHandler` are capped at `xlsx.MaxUploadSize` (32 MiB by default).
`GenerateHandler` does not limit the request body, so wrap it with
`http.MaxBytesHandler` if it faces untrusted clients.

The parse response looks like this:

```json
{
  "meta": { "created_by": "John Doe", "version": "1.0" },
  "sheets": {
    "Sheet1": [["Product", "Price"], ["Widget A", "99.99"]],
    "Sheet2": [["Notes"], ["This is a note"]]
  },
  "order": ["Sheet1", "Sheet2"]
}
```

## JSON format

```json
{
  "meta":   { ... },  // Optional: metadata stored in a very hidden sheet
  "styles": [ ... ],  // Named style definitions
  "sheets": [ ... ]   // Sheet definitions
}
```

### Root Object

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `meta` | object | No | Key-value pairs stored in a very hidden metadata sheet |
| `styles` | array | No | Array of style definitions (see Style Object) |
| `sheets` | array | Yes | Array of sheet definitions (see Sheet Object) |

### Style Object

Defines a reusable style that can be referenced by name in cells and columns.

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `name` | string | Yes | Unique name for the style (used for reference) |
| `bold` | boolean | No | Make text bold |
| `italic` | boolean | No | Make text italic |
| `bg_color` | string | No | Background color in hex format (e.g., "DDDDDD") |
| `font_color` | string | No | Font color in hex format (e.g., "000000") |
| `border` | object | No | Border settings (see Border Object) |
| `align` | string | No | Text alignment: "left", "center", or "right" |
| `number_format` | string | No | Excel number format code (e.g., "#,##0.00") |
| `locked` | boolean | No | Whether cell is locked when sheet is protected (default: true) |

### Border Object

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `top` | boolean | No | Show top border |
| `right` | boolean | No | Show right border |
| `bottom` | boolean | No | Show bottom border |
| `left` | boolean | No | Show left border |

### Sheet Object

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `name` | string | No | Sheet name (defaults to `SheetN`) |
| `cells` | array | Yes | Array of cell definitions (see Cell Object) |
| `columns` | array | No | Array of column definitions (see Column Object) |
| `protection` | string | No | Password for sheet protection |
| `validations` | array | No | Array of data validation rules (see Validation Object) |

### Column Object

Column definitions are indexed by position (0-based). The first column is index 0 (column A), second is index 1 (column B), etc.

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `width` | number | No | Column width |
| `style` | string | No | Reference to a named style |
| `hidden` | boolean | No | Whether the column should be hidden |

### Cell Object

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `cell` | string | Yes | Cell reference in A1 notation (e.g., "A1", "B2") |
| `value` | any | Yes | Cell value (string, number, boolean, or formula starting with "=") |
| `style` | string | No | Reference to a named style |
| `comment` | string | No | Comment text to attach to the cell |

**Value Types:**
- **String**: Treated as text. Strings starting with "=" are interpreted as formulas.
- **Number**: Stored as numeric value.
- **Boolean**: Stored as boolean.
- **Date**: Strings in "YYYY-MM-DD" format are automatically parsed as dates.
- **Formula**: Strings starting with "=" are treated as Excel formulas.

### Validation Object

Defines data validation rules for a cell range.

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `range` | string | Yes | Cell range in A1 notation (e.g., "A1:A10") |
| `rule` | object | Yes | Validation rule (see Validation Rule Object) |

### Validation Rule Object

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `type` | string | Yes | Validation type: "list", "integer", or "date" |
| `list` | array | No | Array of strings for dropdown list (required if type is "list") |
| `min_value` | number | No | Minimum integer value (for "integer" type) |
| `max_value` | number | No | Maximum integer value (for "integer" type) |
| `min_date` | string | No | Minimum date in "YYYY-MM-DD" format (for "date" type) |
| `max_date` | string | No | Maximum date in "YYYY-MM-DD" format (for "date" type) |
| `error_msg` | string | No | Custom error message shown when validation fails |
| `alert_type` | string | No | Alert style: "stop", "warning", or "information" |

## Examples

#### Basic Example

```json
{
  "styles": [
    {
      "name": "header",
      "bold": true,
      "bg_color": "DDDDDD",
      "border": {
        "bottom": true
      },
      "align": "center"
    },
    {
      "name": "currency",
      "align": "right",
      "number_format": "#,##0.00"
    }
  ],
  "sheets": [
    {
      "name": "Sales Report",
      "columns": [
        { "width": 20 },
        { "width": 12 },
        { "width": 15 }
      ],
      "cells": [
        {
          "cell": "A1",
          "value": "Product",
          "style": "header"
        },
        {
          "cell": "B1",
          "value": "Price",
          "style": "header"
        },
        {
          "cell": "A2",
          "value": "Widget A"
        },
        {
          "cell": "B2",
          "value": 99.99,
          "style": "currency"
        }
      ]
    }
  ]
}
```

#### Example with Metadata

```json
{
  "meta": {
    "created_by": "John Doe",
    "created_at": "2024-01-15T10:30:00Z",
    "version": "1.0"
  },
  "styles": [
    {
      "name": "default",
      "align": "left"
    }
  ],
  "sheets": [
    {
      "name": "Data",
      "cells": [
        {
          "cell": "A1",
          "value": "Sample Data"
        }
      ]
    }
  ]
}
```

#### Example with Column Configuration

```json
{
  "styles": [
    {
      "name": "hidden-col",
      "font_color": "FFFFFF"
    }
  ],
  "sheets": [
    {
      "name": "Report",
      "columns": [
        { "width": 20 },
        { "width": 15, "hidden": true },
        { "width": 10, "style": "hidden-col" }
      ],
      "cells": [
        {
          "cell": "A1",
          "value": "Visible Column"
        },
        {
          "cell": "B1",
          "value": "Hidden Column"
        },
        {
          "cell": "C1",
          "value": "Styled Column"
        }
      ]
    }
  ]
}
```

#### Example with Sheet Protection

```json
{
  "styles": [
    {
      "name": "locked",
      "locked": true
    },
    {
      "name": "unlocked",
      "locked": false
    }
  ],
  "sheets": [
    {
      "name": "Protected Sheet",
      "protection": "mypassword",
      "cells": [
        {
          "cell": "A1",
          "value": "This cell is locked",
          "style": "locked"
        },
        {
          "cell": "A2",
          "value": "This cell can be edited",
          "style": "unlocked"
        }
      ]
    }
  ]
}
```

#### Example with Data Validations

```json
{
  "styles": [],
  "sheets": [
    {
      "name": "Form",
      "validations": [
        {
          "range": "A1:A10",
          "rule": {
            "type": "list",
            "list": ["Option 1", "Option 2", "Option 3"],
            "error_msg": "Please select a valid option",
            "alert_type": "stop"
          }
        },
        {
          "range": "B1:B10",
          "rule": {
            "type": "integer",
            "min_value": 1,
            "max_value": 100,
            "error_msg": "Value must be between 1 and 100"
          }
        },
        {
          "range": "C1:C10",
          "rule": {
            "type": "date",
            "min_date": "2024-01-01",
            "max_date": "2024-12-31",
            "error_msg": "Date must be in 2024"
          }
        }
      ],
      "cells": [
        {
          "cell": "A1",
          "value": "Select option"
        },
        {
          "cell": "B1",
          "value": "Enter number"
        },
        {
          "cell": "C1",
          "value": "Enter date"
        }
      ]
    }
  ]
}
```

#### Example with Formulas and Comments

```json
{
  "styles": [
    {
      "name": "total",
      "bold": true,
      "border": {
        "top": true
      }
    }
  ],
  "sheets": [
    {
      "name": "Calculations",
      "cells": [
        {
          "cell": "A1",
          "value": 10
        },
        {
          "cell": "A2",
          "value": 20
        },
        {
          "cell": "A3",
          "value": "=SUM(A1:A2)",
          "style": "total",
          "comment": "This cell calculates the sum of A1 and A2"
        }
      ]
    }
  ]
}
```

#### Example with Date Values

```json
{
  "styles": [],
  "sheets": [
    {
      "name": "Dates",
      "cells": [
        {
          "cell": "A1",
          "value": "2024-01-15"
        },
        {
          "cell": "A2",
          "value": "2024-12-31"
        }
      ]
    }
  ]
}
```

Note: Date strings in "YYYY-MM-DD" format are automatically parsed as dates. Other date formats will be treated as text.

## Notes

- **Style references** must name a style in `styles`; an unknown name is an error.
- **Column indexing** is 0-based: index 0 is column A.
- **Cell references** use A1 notation (`A1`, `B2`, `AA10`).
- **Formulas** start with `=` and use Excel formula syntax.
- **Dates** in `YYYY-MM-DD` form become Excel dates; other date formats stay text.
- **Metadata** is stored as JSON values in a very hidden sheet named `_meta_`, after the data sheets. `Parse` returns it as `meta` and leaves it out of `sheets`.
- **Sheet protection**: with a password set, cells are locked unless their style has `locked: false`.

## License

[MIT](LICENSE)
