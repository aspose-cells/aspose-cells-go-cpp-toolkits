package query

import (
	"math"
	"testing"
	"time"
)

func TestNumericValueClassifiesIntVsFloat(t *testing.T) {
	if got := numericValue(42); got.Kind() != KindInt {
		t.Errorf("numericValue(42) kind = %v, want KindInt", got.Kind())
	} else if i, ok := got.Int(); !ok || i != 42 {
		t.Errorf("numericValue(42) Int = %d, %v; want 42, true", i, ok)
	}

	if got := numericValue(3.14); got.Kind() != KindFloat {
		t.Errorf("numericValue(3.14) kind = %v, want KindFloat", got.Kind())
	} else if f, ok := got.Float(); !ok || math.Abs(f-3.14) > 1e-9 {
		t.Errorf("numericValue(3.14) Float = %f, %v", f, ok)
	}

	if got := numericValue(-7); got.Kind() != KindInt {
		t.Errorf("numericValue(-7) kind = %v, want KindInt", got.Kind())
	}

	// A whole double within int64 range is still an integer.
	if got := numericValue(1e15); got.Kind() != KindInt {
		t.Errorf("numericValue(1e15) kind = %v, want KindInt", got.Kind())
	}

	// A whole double beyond int64 range must not overflow into KindInt.
	if got := numericValue(1e19); got.Kind() != KindFloat {
		t.Errorf("numericValue(1e19) kind = %v, want KindFloat", got.Kind())
	}
}

func TestCellValueAccessorsRejectMismatchedKind(t *testing.T) {
	text := CellValue{kind: KindText, s: "hi"}
	if _, ok := text.Int(); ok {
		t.Error("KindText.Int() ok = true, want false")
	}
	if _, ok := text.Float(); ok {
		t.Error("KindText.Float() ok = true, want false")
	}
	if _, ok := text.Bool(); ok {
		t.Error("KindText.Bool() ok = true, want false")
	}
	if _, ok := text.Time(); ok {
		t.Error("KindText.Time() ok = true, want false")
	}
	if text.IsEmpty() {
		t.Error("KindText IsEmpty() = true, want false")
	}

	empty := CellValue{}
	if !empty.IsEmpty() {
		t.Error("empty CellValue IsEmpty() = false, want true")
	}
	if _, ok := empty.String(); ok {
		t.Error("empty CellValue String() ok = true, want false")
	}
}

func TestCellValueStringFormats(t *testing.T) {
	cases := []struct {
		val  CellValue
		want string
	}{
		{CellValue{kind: KindInt, i: 42}, "42"},
		{CellValue{kind: KindFloat, f: 1.5}, "1.5"},
		{CellValue{kind: KindBool, b: true}, "true"},
		{CellValue{kind: KindText, s: "hello"}, "hello"},
		{CellValue{kind: KindError, s: "#DIV/0!"}, "#DIV/0!"},
		{CellValue{kind: KindDateTime, t: time.Date(2024, 1, 2, 3, 4, 5, 0, time.UTC)}, "2024-01-02T03:04:05Z"},
	}
	for _, tc := range cases {
		got, ok := tc.val.String()
		if !ok {
			t.Errorf("%+v String() ok = false, want a value", tc.val)
			continue
		}
		if got != tc.want {
			t.Errorf("%+v String() = %q, want %q", tc.val, got, tc.want)
		}
	}
}
