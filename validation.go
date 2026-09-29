package xlsx

import (
	"fmt"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

// Validation applies a rule to a cell range such as "A1:A10".
type Validation struct {
	Range string         `json:"range"`
	Rule  ValidationRule `json:"rule"`
}

// ValidationRule restricts what can be entered in a range.
type ValidationRule struct {
	Type      string   `json:"type"`                 // "list", "integer" or "date"
	List      []string `json:"list,omitempty"`       // Choices for "list"
	MinValue  *int     `json:"min_value,omitempty"`  // Bounds for "integer"
	MaxValue  *int     `json:"max_value,omitempty"`  //
	MinDate   *string  `json:"min_date,omitempty"`   // Bounds for "date", as YYYY-MM-DD
	MaxDate   *string  `json:"max_date,omitempty"`   //
	ErrorMsg  string   `json:"error_msg,omitempty"`  // Message shown on invalid input
	AlertType string   `json:"alert_type,omitempty"` // "stop", "warning" or "information"
}

func applyValidation(f *excelize.File, sheet string, v Validation) error {
	rule := v.Rule
	dv := excelize.NewDataValidation(true)
	dv.Sqref = v.Range

	switch rule.AlertType {
	case "":
	case "stop", "warning", "information":
		style := rule.AlertType
		dv.ShowErrorMessage = true
		dv.ErrorStyle = &style
	default:
		return fmt.Errorf("unsupported alert type %q", rule.AlertType)
	}
	if rule.ErrorMsg != "" {
		msg := rule.ErrorMsg
		dv.ShowErrorMessage = true
		dv.Error = &msg
	}

	switch rule.Type {
	case "list":
		if err := dv.SetDropList(rule.List); err != nil {
			return err
		}

	case "integer":
		dv.Type = "whole"
		setBounds(dv, intBound(rule.MinValue), intBound(rule.MaxValue))

	case "date":
		dv.Type = "date"
		prompt, title := "Please enter date in YYYY-MM-DD format", "Date Format"
		dv.ShowInputMessage = true
		dv.PromptTitle, dv.Prompt = &title, &prompt
		if rule.ErrorMsg != "" {
			errTitle := "Invalid Date"
			dv.ErrorTitle = &errTitle
		}

		lo, err := dateBound(rule.MinDate)
		if err != nil {
			return fmt.Errorf("min_date: %w", err)
		}
		hi, err := dateBound(rule.MaxDate)
		if err != nil {
			return fmt.Errorf("max_date: %w", err)
		}
		if lo == "" && hi == "" {
			// No bounds: accept any valid date.
			lo, _ = dateBound(ptr("1900-01-01"))
			hi, _ = dateBound(ptr("9999-12-31"))
		}
		setBounds(dv, lo, hi)

	default:
		return fmt.Errorf("unsupported validation type %q", rule.Type)
	}

	return f.AddDataValidation(sheet, dv)
}

// setBounds picks the operator for whichever of lo and hi are set.
func setBounds(dv *excelize.DataValidation, lo, hi string) {
	switch {
	case lo != "" && hi != "":
		dv.Operator, dv.Formula1, dv.Formula2 = "between", lo, hi
	case lo != "":
		dv.Operator, dv.Formula1 = "greaterThanOrEqual", lo
	case hi != "":
		dv.Operator, dv.Formula1 = "lessThanOrEqual", hi
	default:
		// Any whole number.
		dv.Operator, dv.Formula1, dv.Formula2 = "between", "-2147483648", "2147483647"
	}
}

func intBound(n *int) string {
	if n == nil {
		return ""
	}
	return strconv.Itoa(*n)
}

// dateBound converts a YYYY-MM-DD date to an Excel serial day number.
func dateBound(s *string) (string, error) {
	if s == nil {
		return "", nil
	}
	t, err := time.Parse("2006-01-02", *s)
	if err != nil {
		return "", err
	}
	// Excel counts days from 1899-12-30, which absorbs its 1900 leap-year bug
	// for every date from March 1900 on.
	days := int(t.Sub(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)).Hours() / 24)
	return strconv.Itoa(days), nil
}

func ptr[T any](v T) *T { return &v }
