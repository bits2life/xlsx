package xlsx

import (
	"bytes"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"reflect"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func sampleWorkbook() *Workbook {
	unlocked := false
	lo, hi := 1, 100
	return &Workbook{
		Meta: map[string]any{"version": "1.0", "count": 3.0, "tags": []any{"a", "b"}},
		Styles: []Style{
			{Name: "header", Bold: true, BgColor: "DDDDDD", FontColor: "FF0000", Border: Border{Bottom: true}, Align: "center"},
			{Name: "money", NumberFormat: "#,##0.00"},
			{Name: "open", Locked: &unlocked},
		},
		Sheets: []Sheet{
			{
				Name:       "Report",
				Columns:    []Column{{Width: 20}, {Width: 12, Style: "money"}, {Hidden: true}},
				Protection: "secret",
				Cells: []Cell{
					{Cell: "A1", Value: "Product", Style: "header", Comment: "Name of the product"},
					{Cell: "B1", Value: "Price", Style: "header"},
					{Cell: "A2", Value: "Widget"},
					{Cell: "B2", Value: 99.5, Style: "money"},
					{Cell: "A3", Value: "0012"},
					{Cell: "B3", Value: "=B2*2"},
					{Cell: "C1", Value: true},
					{Cell: "C2", Value: nil, Style: "open"},
				},
				Validations: []Validation{
					{Range: "D1:D10", Rule: ValidationRule{Type: "list", List: []string{"x", "y"}, AlertType: "stop"}},
					{Range: "E1:E10", Rule: ValidationRule{Type: "integer", MinValue: &lo, MaxValue: &hi}},
					{Range: "F1:F10", Rule: ValidationRule{Type: "integer", MinValue: &lo}},
					{Range: "G1:G10", Rule: ValidationRule{Type: "date", MaxDate: ptr("2024-12-31"), ErrorMsg: "Too late"}},
				},
			},
			{
				Name:  "Dates",
				Cells: []Cell{{Cell: "A1", Value: "2024-01-15"}},
			},
		},
	}
}

func TestRoundTrip(t *testing.T) {
	data, err := sampleWorkbook().Bytes()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := Parse(bytes.NewReader(data))
	if err != nil {
		t.Fatal(err)
	}

	if want := []string{"Report", "Dates"}; !reflect.DeepEqual(doc.Order, want) {
		t.Errorf("Order = %v, want %v", doc.Order, want)
	}
	if _, ok := doc.Sheets[MetaSheetName]; ok {
		t.Error("metadata sheet returned as a data sheet")
	}
	wantMeta := map[string]any{"version": "1.0", "count": 3.0, "tags": []any{"a", "b"}}
	if !reflect.DeepEqual(doc.Meta, wantMeta) {
		t.Errorf("Meta = %v, want %v", doc.Meta, wantMeta)
	}

	report := doc.Sheets["Report"]
	for _, c := range []struct {
		row, col int
		want     string
	}{
		{0, 0, "Product"},
		{1, 0, "Widget"},
		{1, 1, "99.50"},
		{2, 0, "0012"},
		{0, 2, "TRUE"},
	} {
		if got := report[c.row][c.col]; got != c.want {
			t.Errorf("Report[%d][%d] = %q, want %q", c.row, c.col, got, c.want)
		}
	}
}

func TestCellTypesAndStyles(t *testing.T) {
	data, err := sampleWorkbook().Bytes()
	if err != nil {
		t.Fatal(err)
	}
	f, err := excelize.OpenReader(bytes.NewReader(data))
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	if formula, _ := f.GetCellFormula("Report", "B3"); formula != "B2*2" {
		t.Errorf("B3 formula = %q", formula)
	}
	if typ, _ := f.GetCellType("Report", "A3"); typ != excelize.CellTypeSharedString && typ != excelize.CellTypeInlineString {
		t.Errorf("A3 type = %v, want string", typ)
	}
	if raw, _ := f.GetCellValue("Dates", "A1", excelize.Options{RawCellValue: true}); raw != "45306" {
		t.Errorf("Dates!A1 raw = %q, want serial 45306", raw)
	}

	id, _ := f.GetCellStyle("Report", "A1")
	st, err := f.GetStyle(id)
	if err != nil {
		t.Fatal(err)
	}
	if !st.Font.Bold || st.Font.Color != "FF0000" || st.Alignment.Horizontal != "center" || len(st.Border) != 1 {
		t.Errorf("header style not applied: font=%+v align=%+v borders=%d", st.Font, st.Alignment, len(st.Border))
	}
	id, _ = f.GetCellStyle("Report", "C2")
	if st, _ := f.GetStyle(id); st.Protection == nil || st.Protection.Locked {
		t.Error("C2 should be unlocked")
	}

	if visible, _ := f.GetColVisible("Report", "C"); visible {
		t.Error("column C should be hidden")
	}
	if w, _ := f.GetColWidth("Report", "A"); w != 20 {
		t.Errorf("column A width = %v", w)
	}
	if visible, _ := f.GetSheetVisible(MetaSheetName); visible {
		t.Error("metadata sheet should be hidden")
	}
	if cs, _ := f.GetComments("Report"); len(cs) != 1 || cs[0].Cell != "A1" {
		t.Errorf("comments = %+v", cs)
	}

	dvs, err := f.GetDataValidations("Report")
	if err != nil {
		t.Fatal(err)
	}
	ops := map[string]string{}
	for _, dv := range dvs {
		ops[dv.Sqref] = dv.Type + " " + dv.Operator
	}
	want := map[string]string{
		"D1:D10": "list ",
		"E1:E10": "whole between",
		"F1:F10": "whole greaterThanOrEqual",
		"G1:G10": "date lessThanOrEqual",
	}
	if !reflect.DeepEqual(ops, want) {
		t.Errorf("validations = %v, want %v", ops, want)
	}
}

func TestErrors(t *testing.T) {
	for name, wb := range map[string]*Workbook{
		"unknown style":   {Sheets: []Sheet{{Cells: []Cell{{Cell: "A1", Value: 1, Style: "nope"}}}}},
		"duplicate style": {Styles: []Style{{Name: "a"}, {Name: "a"}}},
		"bad cell":        {Sheets: []Sheet{{Cells: []Cell{{Cell: "1A", Value: 1}}}}},
		"bad validation":  {Sheets: []Sheet{{Validations: []Validation{{Range: "A1", Rule: ValidationRule{Type: "regex"}}}}}},
		"bad date":        {Sheets: []Sheet{{Validations: []Validation{{Range: "A1", Rule: ValidationRule{Type: "date", MinDate: ptr("2024-13-01")}}}}}},
	} {
		if _, err := wb.Bytes(); err == nil {
			t.Errorf("%s: expected an error", name)
		}
	}
}

func TestGenerateJSON(t *testing.T) {
	in := `{"styles":[],"sheets":[{"name":"S","cells":[{"cell":"A1","value":42},{"cell":"A2","value":"hi"}]}]}`
	var out bytes.Buffer
	if err := Generate(strings.NewReader(in), &out); err != nil {
		t.Fatal(err)
	}
	doc, err := Parse(&out)
	if err != nil {
		t.Fatal(err)
	}
	if got := doc.Sheets["S"]; !reflect.DeepEqual(got, [][]string{{"42"}, {"hi"}}) {
		t.Errorf("S = %v", got)
	}
}

func TestHandlers(t *testing.T) {
	body := `{"meta":{"k":"v"},"sheets":[{"name":"S","cells":[{"cell":"A1","value":"x"}]}]}`
	rec := httptest.NewRecorder()
	GenerateHandler(rec, httptest.NewRequest("POST", "/", strings.NewReader(body)))
	if rec.Code != http.StatusOK || rec.Header().Get("Content-Type") != ContentType {
		t.Fatalf("generate: %d %s", rec.Code, rec.Body)
	}

	var form bytes.Buffer
	mw := multipart.NewWriter(&form)
	fw, _ := mw.CreateFormFile("file", "test.xlsx")
	fw.Write(rec.Body.Bytes())
	mw.Close()
	req := httptest.NewRequest("POST", "/", &form)
	req.Header.Set("Content-Type", mw.FormDataContentType())
	rec = httptest.NewRecorder()
	ParseHandler(rec, req)
	if rec.Code != http.StatusOK {
		t.Fatalf("parse: %d %s", rec.Code, rec.Body)
	}
	if want := `{"meta":{"k":"v"},"sheets":{"S":[["x"]]},"order":["S"]}`; strings.TrimSpace(rec.Body.String()) != want {
		t.Errorf("parse body = %s", rec.Body)
	}

	rec = httptest.NewRecorder()
	GenerateHandler(rec, httptest.NewRequest("POST", "/", strings.NewReader("{")))
	if rec.Code != http.StatusBadRequest {
		t.Errorf("bad JSON: status %d", rec.Code)
	}
	rec = httptest.NewRecorder()
	ParseHandler(rec, httptest.NewRequest("POST", "/", nil))
	if rec.Code != http.StatusBadRequest {
		t.Errorf("no upload: status %d", rec.Code)
	}
}
