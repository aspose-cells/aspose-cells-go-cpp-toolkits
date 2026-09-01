package tests

import (
	"errors"
	"testing"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	"github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/query"
)

func TestParseCellRef(t *testing.T) {
	cases := []struct {
		in   string
		want query.CellRef
	}{
		{"A1", query.CellRef{Row: 0, Col: 0}},
		{"B3", query.CellRef{Row: 2, Col: 1}},
		{"C3", query.CellRef{Row: 2, Col: 2}},
		{"Z26", query.CellRef{Row: 25, Col: 25}},
		{"AA1", query.CellRef{Row: 0, Col: 26}},
		{"AZ1", query.CellRef{Row: 0, Col: 51}},
		{"a1", query.CellRef{Row: 0, Col: 0}}, // case-insensitive
	}
	for _, tc := range cases {
		got, err := query.ParseCellRef(tc.in)
		if err != nil {
			t.Errorf("ParseCellRef(%q) error: %v", tc.in, err)
			continue
		}
		if got != tc.want {
			t.Errorf("ParseCellRef(%q) = %+v, want %+v", tc.in, got, tc.want)
		}
	}
}

func TestParseCellRefInvalid(t *testing.T) {
	for _, in := range []string{"", "1", "B", "A0", "B3x", "1A", "A-1", "AA"} {
		if _, err := query.ParseCellRef(in); !errors.Is(err, toolkiterrors.ErrInvalidCellRef) {
			t.Errorf("ParseCellRef(%q) error = %v, want ErrInvalidCellRef", in, err)
		}
	}
}

func TestCellRefStringRoundTrip(t *testing.T) {
	cases := []struct {
		in   query.CellRef
		want string
	}{
		{query.CellRef{Row: 0, Col: 0}, "A1"},
		{query.CellRef{Row: 2, Col: 1}, "B3"},
		{query.CellRef{Row: 25, Col: 25}, "Z26"},
		{query.CellRef{Row: 0, Col: 26}, "AA1"},
		{query.CellRef{Row: 9, Col: 51}, "AZ10"},
	}
	for _, tc := range cases {
		if got := tc.in.String(); got != tc.want {
			t.Errorf("%+v.String() = %q, want %q", tc.in, got, tc.want)
		}
		// Every rendered reference must parse back to the same coordinate.
		back, err := query.ParseCellRef(tc.in.String())
		if err != nil {
			t.Errorf("ParseCellRef(%q) error: %v", tc.in.String(), err)
			continue
		}
		if back != tc.in {
			t.Errorf("round trip %+v -> %q -> %+v", tc.in, tc.in.String(), back)
		}
	}
}

func TestParseArea(t *testing.T) {
	got, err := query.ParseArea("A1:C3")
	if err != nil {
		t.Fatalf("ParseArea(A1:C3) error: %v", err)
	}
	want := query.Area{Start: query.CellRef{Row: 0, Col: 0}, End: query.CellRef{Row: 2, Col: 2}}
	if got != want {
		t.Errorf("ParseArea(A1:C3) = %+v, want %+v", got, want)
	}

	// A single cell is accepted as a one-cell area.
	single, err := query.ParseArea("B2")
	if err != nil {
		t.Fatalf("ParseArea(B2) error: %v", err)
	}
	if single != (query.Area{Start: query.CellRef{Row: 1, Col: 1}, End: query.CellRef{Row: 1, Col: 1}}) {
		t.Errorf("ParseArea(B2) = %+v, want one-cell area", single)
	}

	if got := (query.Area{Start: query.CellRef{Row: 2, Col: 1}, End: query.CellRef{Row: 4, Col: 2}}).String(); got != "B3:C5" {
		t.Errorf("Area.String() = %q, want %q", got, "B3:C5")
	}
}

func TestParseAreaInvalid(t *testing.T) {
	for _, in := range []string{"C3:A1", "A1:B2:C3", "A1:Z"} {
		if _, err := query.ParseArea(in); err == nil {
			t.Errorf("ParseArea(%q) succeeded, want error", in)
		}
	}
	if _, err := query.ParseArea("C3:A1"); !errors.Is(err, toolkiterrors.ErrInvalidRange) {
		t.Errorf("ParseArea(C3:A1) error = %v, want ErrInvalidRange", err)
	}
	if _, err := query.ParseArea("A1:Z"); !errors.Is(err, toolkiterrors.ErrInvalidCellRef) {
		t.Errorf("ParseArea(A1:Z) error = %v, want ErrInvalidCellRef", err)
	}
}
