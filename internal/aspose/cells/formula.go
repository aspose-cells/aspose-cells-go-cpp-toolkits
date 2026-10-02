package cells

import "strings"

// engineFormulaPrefix is the leading "=" the engine prepends to every formula it
// stores and hands back through its getters.
const engineFormulaPrefix = "="

// StripFormulaPrefix removes the single leading "=" that the engine adds to a
// formula when it is read back.
//
// The engine normalizes every formula it stores to a "=" -prefixed form, so
// validation and conditional-formatting formulas set as "1" read back as "=1",
// and one set as "=A1>0" also reads back as "=A1>0". The engine is idempotent
// here: it never doubles the prefix. The toolkit's DSL takes formulas in their
// bare form (WithValidationFormula1("1"), WithDataRangeRule("A1>0")), so the
// read side strips the prefix to keep a value round-tripping as it was set.
//
// Exactly one prefix is removed, and only at the start: "=" becomes "", and the
// inner "=" of "=A1=(B1)" survives.
//
// Do NOT apply this to a list validation's Formula1. A list holds literal
// comma-separated values rather than a formula, and measured, the engine does
// not prefix it: a list of ["Yes","No"] reads back as "Yes,No". Stripping there
// would corrupt a list whose first entry legitimately begins with "=".
func StripFormulaPrefix(formula string) string {
	return strings.TrimPrefix(formula, engineFormulaPrefix)
}
