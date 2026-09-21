# ADR-0005: Literal, Enum, Pattern, and Email value-constraint boundaries

## Status

`accepted`

## Background

If `Literal` and `EnumOf` depend on implicit Variant coercion, the boundaries between String, Number, Boolean, Date, Null, and Empty become ambiguous. Sharing the input array for `EnumOf` would also allow external mutation after schema construction to change validation meaning.

`Pattern` and `Email` require a regular-expression runtime. VBA-Schema verifies Windows 64-bit Office and remains compatibility-conscious but unverified and unsupported on macOS and Windows 32-bit, so RegExp configuration, failure boundaries, and the Email scope must be fixed.

## Decision

- Limit `Literal` candidates to the scalar categories String, Number, Boolean, Date, Null, and Empty. Reject Error Variants, Objects, and Arrays during schema construction with `vbObjectError + 2100`.
- `EnumOf` accepts one one-dimensional native array and normalizes and snapshots candidate values at construction. Typed scalar arrays are accepted; uninitialized, multidimensional, or empty arrays, and Error/Object/Array elements, raise `vbObjectError + 2100`.
- Duplicate `EnumOf` candidates are meaningless definitions and raise `vbObjectError + 2101`. Use binary comparison for Strings, the DG-002 lossless comparison for Numbers, and same-category comparison for other categories.
- Late-bind `VBScript.RegExp` for `Pattern`, fixing `IgnoreCase=False`, `Global=False`, and `MultiLine=False`. The expression succeeds only on a full-input match. Compile RegExp lazily during validation.
- Do not convert Pattern-expression compilation failures into validation Issues; raise `vbObjectError + 2100) at the first validation. Failure to create the RegExp component raises `vbObjectError + 2200`.
- Define `Email` as a practical ASCII form implemented with `VBScript.RegExp`. Require exactly one `@` between local and domain and a total length of at most 254 UTF-16 code units. Accept common ASCII atoms and dot-separated local/domain labels, but exclude empty values, local-only values, multiple `@`, whitespace, Unicode, quoted local parts, comments, IP literals, and all other RFC extensions.
- When `Pattern` and `Email` are combined, preserve the existing contract order: Pattern, then Email. Both apply only to Text schemas.
- Enum child-schema behavior and Union Collection ownership are outside this ADR; the Union decision is recorded separately in ADR-0006 for milestone M6.

## Consequences

- Mutating the input array after Enum construction does not change validation results.
- Candidate categories and duplicate detection are deterministic; `"1"` and `1`, Boolean and Number, and Null and Empty are not conflated.
- Pattern is a full match rather than a search, so substring matching requires an explicit expression.
- Email detects common input mistakes and does not provide full RFC compliance. The scope must be stated in the README.
- A host without VBScript.RegExp produces an environment failure rather than a validation failure. Separate evidence is required for macOS and 32-bit Office.

## Rationale

- Tests: M5 focused tests fix category, snapshot, duplicate, Pattern-flag, full-match, and Email boundary behavior.
- Code: `VSchema.cls` centralizes scalar comparison, Enum snapshots, and lazy RegExp creation without adding external references.
- Related specs: `docs/specs/v1-contract.md`, `docs/design.md`

## Supersedes

- None

## Superseded by

- None
