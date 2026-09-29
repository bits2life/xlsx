package xlsx

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// Style is a named, flat cell style that cells and columns refer to by name.
type Style struct {
	Name         string `json:"name"`
	Bold         bool   `json:"bold,omitempty"`
	Italic       bool   `json:"italic,omitempty"`
	BgColor      string `json:"bg_color,omitempty"`      // Hex color, e.g. "DDDDDD"
	FontColor    string `json:"font_color,omitempty"`    // Hex color, e.g. "FF0000"
	Border       Border `json:"border,omitempty"`        // Thin black borders per side
	Align        string `json:"align,omitempty"`         // "left", "center" or "right"
	NumberFormat string `json:"number_format,omitempty"` // Excel number format code
	Locked       *bool  `json:"locked,omitempty"`        // Locked when the sheet is protected; defaults to true
}

// Border selects which sides of a cell get a thin black border.
type Border struct {
	Top    bool `json:"top,omitempty"`
	Right  bool `json:"right,omitempty"`
	Bottom bool `json:"bottom,omitempty"`
	Left   bool `json:"left,omitempty"`
}

// styleSet creates excelize styles on first use and caches their IDs.
type styleSet struct {
	f      *excelize.File
	styles map[string]Style
	ids    map[string]int
}

func newStyleSet(f *excelize.File, styles []Style) (*styleSet, error) {
	s := &styleSet{f: f, styles: make(map[string]Style, len(styles)), ids: make(map[string]int)}
	for _, st := range styles {
		if st.Name == "" {
			return nil, fmt.Errorf("xlsx: style without a name")
		}
		if _, dup := s.styles[st.Name]; dup {
			return nil, fmt.Errorf("xlsx: duplicate style %q", st.Name)
		}
		s.styles[st.Name] = st
	}
	return s, nil
}

func (s *styleSet) id(name string) (int, error) {
	if id, ok := s.ids[name]; ok {
		return id, nil
	}
	st, ok := s.styles[name]
	if !ok {
		return 0, fmt.Errorf("unknown style %q", name)
	}
	id, err := s.f.NewStyle(st.excelize())
	if err != nil {
		return 0, fmt.Errorf("style %q: %w", name, err)
	}
	s.ids[name] = id
	return id, nil
}

func (st Style) excelize() *excelize.Style {
	locked := true
	if st.Locked != nil {
		locked = *st.Locked
	}
	out := &excelize.Style{
		Font: &excelize.Font{
			Bold:   st.Bold,
			Italic: st.Italic,
			Color:  st.FontColor,
		},
		Protection: &excelize.Protection{Locked: locked},
	}

	if st.BgColor != "" {
		out.Fill = excelize.Fill{Type: "pattern", Color: []string{st.BgColor}, Pattern: 1}
	}

	for _, side := range []struct {
		on   bool
		name string
	}{
		{st.Border.Left, "left"},
		{st.Border.Top, "top"},
		{st.Border.Right, "right"},
		{st.Border.Bottom, "bottom"},
	} {
		if side.on {
			out.Border = append(out.Border, excelize.Border{Type: side.name, Color: "000000", Style: 1})
		}
	}

	if st.Align != "" {
		out.Alignment = &excelize.Alignment{Horizontal: st.Align}
	}
	if st.NumberFormat != "" {
		format := st.NumberFormat
		out.CustomNumFmt = &format
	}
	return out
}
