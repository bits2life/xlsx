package xlsx

import (
	"encoding/json"
	"fmt"
	"sort"

	"github.com/xuri/excelize/v2"
)

// MetaSheetName is the very hidden sheet that holds Workbook.Meta as
// JSON-encoded values in columns A (key) and B (value), below a header row.
const MetaSheetName = "_meta_"

func writeMeta(f *excelize.File, meta map[string]any) error {
	if _, err := f.NewSheet(MetaSheetName); err != nil {
		return fmt.Errorf("xlsx: creating metadata sheet: %w", err)
	}

	rows := [][]any{{"Key", "Value"}}
	keys := make([]string, 0, len(meta))
	for k := range meta {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	for _, k := range keys {
		v, err := json.Marshal(meta[k])
		if err != nil {
			return fmt.Errorf("xlsx: encoding metadata %q: %w", k, err)
		}
		rows = append(rows, []any{k, string(v)})
	}
	for i, row := range rows {
		cell, _ := excelize.CoordinatesToCellName(1, i+1)
		if err := f.SetSheetRow(MetaSheetName, cell, &row); err != nil {
			return fmt.Errorf("xlsx: writing metadata: %w", err)
		}
	}

	if err := f.SetSheetVisible(MetaSheetName, false, true); err != nil {
		return fmt.Errorf("xlsx: hiding metadata sheet: %w", err)
	}
	return nil
}

func readMeta(f *excelize.File) (map[string]any, error) {
	rows, err := f.GetRows(MetaSheetName)
	if err != nil {
		return nil, fmt.Errorf("xlsx: reading metadata: %w", err)
	}
	meta := make(map[string]any)
	for _, row := range rows[min(1, len(rows)):] { // Skip the header row
		if len(row) < 2 {
			continue
		}
		var v any
		if err := json.Unmarshal([]byte(row[1]), &v); err != nil {
			return nil, fmt.Errorf("xlsx: decoding metadata %q: %w", row[0], err)
		}
		meta[row[0]] = v
	}
	return meta, nil
}
