# internal/rows

Package rows maps struct fields to spreadsheet columns. It is shared by the `query` package's structured reader (`ReadRows`) and the `editor` package's structured writer (`WriteRows`) so both sides agree on the exact same mapping rules.

## Overview

The `rows` package provides the column mapping logic that `ReadRows` and `WriteRows` share. This ensures that:

1. A struct field mapped by `ReadRows` can be written back with `WriteRows`
2. Both operations use identical column resolution rules
3. Column mapping behavior never drifts between read and write

## Column mapping rules

Each struct field is mapped to a spreadsheet column based on its `excel` tag or field name:

| Tag | Effect |
|-----|--------|
| `excel:"name"` | Column header is `name` (match is case-insensitive) |
| `excel:"-"` | Field is skipped entirely |
| *(no tag)* | Column header is the field's own name |

### Example

```go
type Employee struct {
    ID       int    `excel:"id"`
    Name     string `excel:"name"`
    SkipThis string `excel:"-"`      // this field is ignored
    HireDate time.Time               // uses field name "HireDate" as header
}
```

This struct would map to a spreadsheet with headers: `id`, `name`, and `HireDate` (case-insensitive match).

## Functions

### Columns

```go
func Columns(t reflect.Type) ([]Column, error)
```

Validates `t` as a table row type and returns its ordered columns.

#### Parameters

- **t**: The struct type to analyze. Can be a struct type or a pointer to a struct.

#### Returns

- **[]Column**: An ordered slice of `Column` structs, each containing:
  - `Name`: The resolved column name (from `excel` tag or field name)
  - `Index`: The position of the field in the struct's top-level fields
  - `Type`: The field's type

- **error**: An error if:
  - `t` is not a struct type
  - `t` has no exported fields with valid mappings
  - All exported fields are skipped with `excel:"-"`

#### Example

```go
typ := reflect.TypeOf(Employee{})
cols, err := rows.Columns(typ)
if err != nil {
    log.Fatal(err)
}

for _, col := range cols {
    fmt.Printf("Field: %s, Column: %s, Type: %s\n",
        typ.Field(col.Index).Name, col.Name, col.Type)
}
```

## Column type

```go
type Column struct {
    Name  string
    Index int
    Type  reflect.Type
}
```

- **Name**: The resolved column name. This is the `excel` tag value if present, otherwise the struct field name.
- **Index**: The zero-based position of the field in the struct's top-level fields (not nested fields).
- **Type**: The Go type of the field (e.g., `reflect.TypeOf(int(0))`).

## Usage in ReadRows

`query.ReadRows` uses `rows.Columns` to:

1. Parse the struct type's column mappings
2. Match headers in the spreadsheet to struct fields
3. Validate that all required fields have corresponding headers
4. Map header positions to struct field indices

## Usage in WriteRows

`editor.WriteRows` uses `rows.Columns` to:

1. Parse the struct type's column mappings
2. Determine which headers to write (if `WithWriteHeader` is true)
3. Map struct field indices to column positions when writing data rows

## Field type support

The `rows` package doesn't validate field types; that's handled by `ReadRows` and `WriteRows`. However, the supported types are:

- **String**: `string`
- **Integer types**: `int`, `int8`, `int16`, `int32`, `int64`
- **Unsigned integer types**: `uint`, `uint8`, `uint16`, `uint32`, `uint64`
- **Float types**: `float32`, `float64`
- **Boolean**: `bool`
- **Time**: `time.Time`

## Important notes

1. **Exported fields only**: Only exported (capitalized) fields are mapped. Unexported fields are always ignored.

2. **Top-level fields only**: Embedded structs are not flattened. Only top-level struct fields are considered.

3. **Order matters**: The returned `Column` slice is in the same order as the fields appear in the struct definition.

4. **Case sensitivity**: Header matching in `ReadRows` is case-insensitive, but the `Name` stored in `Column` preserves the original tag/field name.

5. **No tag = field name**: If no `excel` tag is present, the field name itself becomes the column header.

## Example: Complete flow

```go
type Product struct {
    SKU      string  `excel:"sku"`
    Name     string  `excel:"product_name"`
    Price    float64 `excel:"price"`
    Quantity int     `excel:"qty"`
    Hidden   string  `excel:"-"` // ignored
}

// Step 1: Parse column mappings
typ := reflect.TypeOf(Product{})
cols, err := rows.Columns(typ)
if err != nil {
    log.Fatal(err)
}

// cols = [
//   {Name: "sku", Index: 0, Type: string},
//   {Name: "product_name", Index: 1, Type: string},
//   {Name: "price", Index: 2, Type: float64},
//   {Name: "qty", Index: 3, Type: int},
// ]

// Step 2: Use in ReadRows
products, err := query.ReadRows[Product](source, opts...)
// ReadRows uses cols to match headers and decode rows

// Step 3: Use in WriteRows
err = editor.WriteRows(source, sink, products, opts...)
// WriteRows uses cols to write headers and data
```

## Why a separate package?

By extracting the column mapping logic into a separate package:

1. **Consistency**: Both `ReadRows` and `WriteRows` use identical mapping rules
2. **Maintainability**: Changes to mapping rules only need to be made in one place
3. **Testability**: The mapping logic can be tested independently of I/O operations
4. **Clear separation**: Mapping concerns are separate from spreadsheet I/O concerns