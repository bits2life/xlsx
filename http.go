package xlsx

import (
	"bytes"
	"encoding/json"
	"net/http"
)

// ContentType is the MIME type of .xlsx files.
const ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

// MaxUploadSize caps the size of the request body ParseHandler accepts.
var MaxUploadSize int64 = 32 << 20

// GenerateHandler responds to a JSON Workbook in the request body with the
// generated .xlsx file as an attachment named generated.xlsx, or with a JSON
// {"error": "..."} and status 400.
func GenerateHandler(w http.ResponseWriter, r *http.Request) {
	var wb Workbook
	if err := json.NewDecoder(r.Body).Decode(&wb); err != nil {
		writeError(w, http.StatusBadRequest, "invalid workbook: "+err.Error())
		return
	}
	var buf bytes.Buffer
	if err := wb.Write(&buf); err != nil {
		writeError(w, http.StatusBadRequest, err.Error())
		return
	}
	w.Header().Set("Content-Type", ContentType)
	w.Header().Set("Content-Disposition", `attachment; filename="generated.xlsx"`)
	w.WriteHeader(http.StatusOK)
	w.Write(buf.Bytes())
}

// ParseHandler parses an .xlsx file uploaded as the multipart form field
// "file" and responds with the Document as JSON, or with a JSON
// {"error": "..."} and status 400.
func ParseHandler(w http.ResponseWriter, r *http.Request) {
	r.Body = http.MaxBytesReader(w, r.Body, MaxUploadSize)
	file, _, err := r.FormFile("file")
	if err != nil {
		writeError(w, http.StatusBadRequest, "no file uploaded: "+err.Error())
		return
	}
	defer file.Close()

	doc, err := Parse(file)
	if err != nil {
		writeError(w, http.StatusBadRequest, "invalid Excel file: "+err.Error())
		return
	}
	writeJSON(w, http.StatusOK, doc)
}

func writeError(w http.ResponseWriter, status int, msg string) {
	writeJSON(w, status, map[string]string{"error": msg})
}

func writeJSON(w http.ResponseWriter, status int, v any) {
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(status)
	json.NewEncoder(w).Encode(v)
}
