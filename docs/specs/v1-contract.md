# VBA-Schema v1 Contract

## 1. Status and authority

This document defines the public API, validation semantics, error contract, compatibility, and distribution boundary for VBA-Schema v1. If an example or conceptual explanation in `docs/design.md` conflicts with this document, this document takes precedence.

## 2. Public API

### 2.1 Factory API

`Schema.bas` exposes the following factories.

```vb
Public Function AnyValue() As VSchema
Public Function Text() As VSchema
Public Function Number() As VSchema
Public Function Bool() As VSchema
Public Function DateTime() As VSchema
Public Function ObjectSchema() As VSchema
Public Function ArrayOf(ByVal ItemSchema As VSchema) As VSchema
Public Function Literal(ByVal ExpectedValue As Variant) As VSchema
Public Function EnumOf(ByVal Values As Variant) As VSchema
Public Function UnionOf(ByVal Schemas As Collection) As VSchema
```

- `ArrayOf` accepts a `VSchema`.
- `EnumOf` accepts a one-dimensional native array containing scalar values and snapshots the candidates when the schema is constructed. Typed scalar arrays are accepted.
- `UnionOf` accepts a non-empty `Collection` containing only `VSchema` instances. An `Array()` containing objects is not part of the v1 public contract. The branch order and element set of the Collection are snapshotted at construction time, while each child `VSchema` instance remains shared by reference rather than being cloned.
- `ArrayOf(Nothing)`, `UnionOf(Nothing)`, uninitialized, multidimensional, or empty Enum arrays, empty Unions, invalid Enum elements, duplicate Enum candidates, uninitialized `VSchema` instances, and Union elements that are not `VSchema` instances raise `Err.Raise` as programmer misuse.

### 2.2 Fluent API

`VSchema` exposes the following builders.

```vb
Public Function OptionalField() As VSchema
Public Function Nullable() As VSchema
Public Function Min(ByVal Boundary As Variant) As VSchema
Public Function Max(ByVal Boundary As Variant) As VSchema
Public Function Length(ByVal RequiredLength As Variant) As VSchema
Public Function WholeNumber() As VSchema
Public Function Pattern(ByVal Expression As String) As VSchema
Public Function Email() As VSchema
Public Function Field(ByVal FieldName As String, ByVal FieldSchema As VSchema) As VSchema
Public Function Strict() As VSchema
```

`OptionalField` allows an Object field to be absent. It has no effect on `Empty`, `Null`, or standalone values. `WholeNumber` requires a Number that is mathematically integral.

Each builder mutates and returns the same instance. Builders do not silently overwrite an existing setting. Duplicate use of the same constraint, `Min > Max`, negative lengths, and modifiers that do not apply to the schema kind are programmer misuse.

Constraint evaluation order is fixed for each schema node. After type and Null/Empty checks, constraints are evaluated in this order: `Length`, `Min`, `Max`, `WholeNumber`, `Pattern`, and `Email`. Only the first failed constraint for that node is added as an Issue. Object fields apply the same rule independently, and the overall Issue order preserves field declaration order.

`Length(n).Min(m)` is programmer misuse when `m > n`, and `Length(n).Max(m)` is programmer misuse when `m < n`. The builder detects the same condition even when the chain is written in the reverse order. The combination of `Min`/`Max` and `WholeNumber` is allowed only for Number schemas. Combining `Pattern` and `Email` is allowed only for Text schemas and follows the fixed order above.

The supported combinations are:

| Modifier | Schema kind |
| --- | --- |
| `OptionalField` | All kinds; meaningful only when used as an Object field |
| `Nullable` | All kinds |
| `Min`, `Max` | Text, Number, Array |
| `Length` | Text, Array |
| `WholeNumber` | Number |
| `Pattern`, `Email` | Text |
| `Field`, `Strict` | ObjectSchema |

### 2.3 Internal initialization (unsupported)

Because of VBA class initialization constraints, `VSchema` exposes the following procedure as Public.

```vb
Public Function InternalInitialize(ByVal KindCode As Long) As VSchema
```

This is not a stable public API for v1 users. Only factories in `Schema.bas` pass an internal encoded kind. A direct call that passes an ordinary schema kind is rejected with `vbObjectError + 2100`. Invalid encoded kinds and reinitialization are also programmer misuse, and use from outside the factories is not supported. This is a limitation caused by VBA's public visibility rules, not a security boundary.

Because VBA cannot access a Private member through another instance of the same class, the `Internal*` hooks required to compose child schemas, perform nested validation with paths, and preflight the schema graph are also Public. These are unsupported internal-only APIs, not stable user APIs or security boundaries.

### 2.4 Validation API

```vb
Public Function SafeParse(ByVal InputValue As Variant) As VValidationResult

Public Property Get Success() As Boolean
Public Property Get Value() As Variant
Public Property Get Issues() As Collection
Public Property Get ErrorText() As String
```

Functions returning `VSchema` or `VValidationResult`, and the `Value` property when it contains an Object, follow VBA Object assignment rules and must be received with `Set`. Scalar `Value` results use ordinary assignment.

`SafeParse` does not raise an error for validation failures. Its success and failure contract is:

| Property | Success | Failure |
| --- | --- | --- |
| `Success` | `True` | `False` |
| `Value` | The input value; Objects keep the same reference | `Empty` |
| `Issues` | An empty Collection | A snapshot Collection with one or more items |
| `ErrorText` | An empty string | A deterministic string generated from Issues |

`Issues` returns a snapshot on every call. Mutating the returned Collection or an Issue Dictionary does not change the internal Result state or `ErrorText`.

On success, `Value` is the original input value. Validation never transforms, coerces, or copies the input; Object input keeps the same reference in `Value`.

`ErrorText` uses this fixed grammar:

```text
<path>: <message> (expected=<expected>, received=<received>)
```

Multiple Issues are joined with `vbCrLf`, with no extra leading or trailing line break. `expected` and `received` use the strings from the Issue snapshots as-is; the raw input or secrets are not emitted.

## 3. Error boundary

### 3.1 Validation failure

When an input does not satisfy a schema, return `Success=False`. Type mismatches, missing required fields, constraint violations, and unknown fields in a strict Object are validation failures.

### 3.2 Programmer misuse

Invalid schema construction raises `Err.Raise` in the range `vbObjectError + 2100` through `vbObjectError + 2199`. Examples include:

- `ArrayOf(Nothing)`
- An empty field name or `Field(name, Nothing)`
- Duplicate field names
- A modifier that does not apply to the schema kind
- Contradictory or duplicate constraints
- An empty Enum or Union
- A cyclic schema graph (`vbObjectError + 2103`)

The initial allocation is:

```text
vbObjectError + 2100  invalid argument
vbObjectError + 2101  invalid constraint
vbObjectError + 2102  invalid schema composition
vbObjectError + 2103  cyclic schema
```

### 3.3 Runtime or environment failure

When a required runtime component is unavailable or an internal invariant is violated, raise an error in the range `vbObjectError + 2200` through `vbObjectError + 2299`. `SafeParse` does not convert these into validation Issues.

```text
vbObjectError + 2200  required runtime component unavailable
vbObjectError + 2299  internal invariant failure
```

A normal direct call to `InternalInitialize`, an invalid kind, or a corrupt factory-encoded kind uses `vbObjectError + 2100`. Reinitializing an initialized instance uses `vbObjectError + 2102`.

## 4. Issue contract

Each Issue is a snapshot of a late-bound `Scripting.Dictionary` and always contains these String keys.

| Key | Type | Contract |
| --- | --- | --- |
| `path` | String | Canonical path |
| `code` | String | Public error code |
| `message` | String | Human-readable message |
| `expected` | String | Stable description of the expected condition |
| `received` | String | Safe summary of the input type and value |

`received` uses this structured grammar and never exposes an Object reference or the array itself.

```text
String("<escaped>")
Number(<invariant-number>)
Boolean(True|False)
Date(<yyyy-mm-ddThh:nn:ss>)
Null
Empty
Error(<number>)
Missing
Object(<TypeName>)
Array(<TypeName>)
```

String values escape backslashes, double quotes, and control characters, and the complete escaped result is limited to 80 UTF-16 code units. When it exceeds the limit, it is truncated with the `...` marker included. Numbers use a locale-independent decimal point (`.`) and no thousands separators. Dates discard sub-second precision and use `yyyy-mm-ddThh:nn:ss`. This prevents an Issue from depending on the lifetime or mutation of the input Object or the host locale.

Issue order must be deterministic:

1. Depth-first
2. Object fields in schema declaration order
3. Arrays in logical index order
4. Non-string Object keys in binary ordinal order
5. Strict unknown fields in binary ordinal order

The public v1 error codes are:

```text
required
invalid_type
too_small
too_big
invalid_length
invalid_integer
invalid_pattern
invalid_email
invalid_literal
invalid_enum
invalid_union
invalid_key
unknown_field
invalid_array_rank
```

## 5. Canonical path

- The root is always `$`.
- Fields matching `[A-Za-z_][A-Za-z0-9_]*` use the `$.user.name` form.
- Other fields use the `$["a.b"]` form. Backslashes, double quotes, line feeds, carriage returns, and tabs are escaped as `\\`, `\"`, `\n`, `\r`, and `\t`; other control characters use `\uXXXX`.
- Arrays and Collections normalize enumeration order to zero-based logical indexes regardless of the underlying base index.
- Native arrays accepted by v1 are one-dimensional only.
- Multidimensional arrays produce an `invalid_array_rank` Issue.
- An uninitialized dynamic array is treated as an array with zero elements.
- Array length constraints (`Length`, `Min`, and `Max`) are evaluated before element validation, and only the first constraint Issue is added to the array node.
- Native array rank/bounds probing is contained in a limited helper. Errors from `LBound`/`UBound` used to detect an uninitialized array must not escape as validation errors. Input arrays and Collections are not modified.

Examples:

```text
$
$.user.name
$.users[0]
$.users[2].email
$["field.with.dot"]
```

## 6. Object field semantics

- Field names use `vbBinaryCompare`-equivalent case-sensitive comparison.
- Field matching does not depend on the input Dictionary's `CompareMode`; it uses key enumeration and `StrComp(..., vbBinaryCompare)`.
- Registering the same field name twice is programmer misuse.
- v1 limits Object input to `Scripting.Dictionary`.
- Object detection begins with `TypeName(value) = "Dictionary"` and performs limited capability checks for `Count`, `Exists`, and `Keys`. A TypeName mismatch or failed capability check produces an `invalid_type` Issue.
- A `Collection` or arbitrary class instance is not treated as a key-value Object.
- A non-String input key produces an `invalid_key` Issue regardless of whether `.Strict()` is used. The path is the Object node path, `expected` is `String key`, and `received` is a safe descriptor. Multiple keys use binary ordinal descriptor order.
- Unknown fields pass through by default; `.Strict()` produces an `unknown_field` Issue.
- A strict unknown-field path is canonical, and `received` is the String descriptor from `DescribeValue`; order is binary ordinal.
- A missing required field uses `Missing` as `received`. This is distinct from a present `Empty` or `Null`.
- Validation does not modify the input Dictionary.

## 7. Type and comparison semantics

- `AnyValue` accepts every Variant state, including Null, Empty, Error Variant, and Nothing.
- No implicit String/Number/Boolean/Date conversion is performed.
- Number accepts Byte, Integer, Long, Single, Double, and Currency.
- Decimal Variants and LongLong are not accepted in v1 until the same contract can be verified on both 32-bit and 64-bit hosts.
- Number type acceptance does not depend on conversion to Decimal. A finite Double outside the Decimal range is accepted by Number type checking alone.
- Number constraint boundaries are stored as Variants and are not unconditionally converted to Double.
- Number `Min`/`Max` boundaries must be supported numeric Variants.
- Text/Array `Min`/`Max` and `Length` accept numeric Variants that are mathematically integral, non-negative, and within the Long range. They are normalized to Long for internal storage.
- Passing Null, Empty, an Error Variant, Object, Array, Boolean, Date, or String as a Number boundary is an `invalid argument`.
- Passing a fraction, a value outside the Long range, a negative value, or a non-numeric value as a Text/Array boundary or `Length` is an `invalid argument`.
- Comparisons that would produce NaN, Infinity, or overflow are rejected.
- Byte/Integer/Long values use integer comparison, Single values use lossless widening to Double, and Currency values prefer Currency comparison. When different numeric categories are compared, conversion to a common representation is allowed only when round-trip checks prove that the value does not change. If overflow or precision loss cannot be proven absent, reject with `vbObjectError + 2100` rather than performing unconditional `CDbl`.
- A Double/Currency boundary involving NaN, Infinity, or an indeterminate comparison is rejected.
- `WholeNumber` succeeds when a Number is mathematically integral. Single/Double values are checked as floating-point values, and Currency values preserve Currency precision; conversion to Decimal is not required.
- Literal/Enum distinguish String, Number, Boolean, Date, Null, and Empty categories. Error Variants, Objects, and Arrays are not accepted as definition candidates.
- Within the Number category, values match when numerically equal even if their numeric subtypes differ.
- String comparisons and field-name comparisons use binary comparison.
- `EnumOf` snapshots candidate values at construction time; later changes to the source array do not change schema meaning. Duplicate candidates raise `vbObjectError + 2101`.
- String `Length` is the same UTF-16 code-unit count as VBA `Len`.

`UnionOf` validates branches in order and succeeds on the first success. Nested Unions are not flattened; their structure is preserved. When every branch fails validation, internal branch Issues are discarded and one `invalid_union` Issue (`expected=Union`) is added at the failed path. Programmer misuse, runtime/environment failures, and internal invariant failures raised inside a branch are not converted into Union validation failures; they are propagated unchanged.

`Pattern` late-binds `VBScript.RegExp` and fixes `IgnoreCase=False`, `Global=False`, and `MultiLine=False`. The expression succeeds only when it matches the entire input. The RegExp instance is compiled lazily, and an invalid regular expression raises `vbObjectError + 2100` at the first validation. If the RegExp runtime is unavailable, raise the `vbObjectError + 2200` environment failure.

`Email` uses the same RegExp runtime and is a practical ASCII-only form. It separates local and domain with one `@`, limits total length to 254 UTF-16 code units, and excludes whitespace, Unicode, quoted local parts, comments, and IP literals. It does not claim full RFC compliance. When `Pattern` and `Email` are combined, evaluate Pattern first and Email second.

## 8. Schema graph and mutation

- Builders are mutable; changes made after sharing a schema are visible through every reference.
- `Field` and `ArrayOf` do not clone child schemas.
- Self-referential and indirectly cyclic schema graphs are outside the v1 contract. A DFS preflight rejects them with `vbObjectError + 2103` before validation begins. The active path is tracked with a Collection, schema identity is checked with VBA `Is`, and non-cyclic reuse of a shared child is allowed.

## 9. Compatibility contract

| Environment | Status |
| --- | --- |
| Windows 64-bit Office 2016+ | Development and verification target |
| Windows 32-bit Office 2016+ | Compatibility-conscious, unverified, not guaranteed |
| macOS Office | Compatibility-conscious, unverified, not guaranteed |

To preserve compatibility, the core must not depend on the Windows API, pointer-size-specific declarations, or the Excel Object Model. `Scripting.Dictionary` and `VBScript.RegExp` are late-bound, but their availability is not guaranteed on every target. A missing required component is a runtime/environment failure.

Documentation and release notes must not describe an unverified environment as supported or verified.

## 10. Distribution contract

The canonical production source consists of these three files:

```text
src/modules/Schema.bas
src/classes/VSchema.cls
src/classes/VValidationResult.cls
```

The release artifact contains only these three files. It excludes `src/modules/Tests`, `src/modules/Xlflow`, `App.bas`, `Main.bas`, `Ui.bas`, and workbook document modules.

`src/` is for development and tests, `dist/VBA-Schema/` is a generated, non-committed import payload built from the three-file allowlist, and the tracked `.xlsm` is an xlflow development fixture. The complete xlflow source tree is not a release artifact. Release staging creates a temporary project from the explicit three-file allowlist.

Staged VBA source must be UTF-8 without a BOM and use LF line endings. README, LICENSE, CHANGELOG, sample workbooks, ZIP files, and other supporting assets are provided separately from the three-file import payload and must not increase its component count.

The GitHub Release workflow accepts only stable `vMAJOR.MINOR.PATCH` tag pushes. It runs the reusable `source-check` workflow first and generates `VBA-Release-vX.Y.Z.zip` only after it succeeds. The ZIP root is `VBA-Release/` and contains exactly these three entries:

```text
VBA-Release/Schema.bas
VBA-Release/VSchema.cls
VBA-Release/VValidationResult.cls
```

Prerelease tags such as `-rc.1` are outside the v1 workflow. Re-running the workflow for the same tag may update same-named assets on an existing GitHub Release with clobber behavior.

The Destination of `tools/release-stage.ps1` must be inside the repository, but it must not be the same as, a child of, or an ancestor of the repository root, `src`, `.git`, `.xlflow`, or `build`. These checks occur before an existing Destination is deleted, so a mistaken release-staging path cannot delete production source or the verification workbook.

The release gate must prove:

1. The three files import into an empty macro-enabled workbook.
2. VBE compilation succeeds.
3. The scalar smoke test succeeds without adding external references.
4. The imported non-document component set is exactly those three files.

The GitHub-hosted Excel-free `source-check` does not run the VBE compilation or behavioral tests above. Before a stable tag, a maintainer must collect Windows 64-bit Excel evidence for VBE compilation, behavioral tests, and release smoke according to `docs/release-checklist.md`. The release workflow does not automatically prove this local evidence, and source-check success must not be reported as VBE compilation passed.

The release provides `VBA-Release-vX.Y.Z.zip` and the corresponding `VBA-Release-vX.Y.Z.zip.sha256` under the same tag. The checksum file contains the ZIP SHA-256 and filename.

## 11. Compile contract

Maintain a compile-only fixture that uses every Public API. When changing a public name, argument type, return type, or fluent chain, update the fixture before the implementation and pass VBE compilation on Windows 64-bit Office.

As of 2026-09-21, the following API candidate passes VBE compilation on Windows 64-bit Office:

```vb
Schema.ObjectSchema() _
    .Field("name", Schema.Text().Min(1).Max(100)) _
    .Field("age", Schema.Number().WholeNumber().Min(0).OptionalField()) _
    .Field("enabled", Schema.Bool().Nullable()) _
    .Field("createdAt", Schema.DateTime()) _
    .Field("tags", Schema.ArrayOf(Schema.Text()).Length(3)) _
    .Strict()
```

## 12. Locale policy

Strict schemas validate the existing VBA value and do not parse localized textual representations of numbers, dates, or any other type. The same Variant input must produce the same validation result on every machine locale.

```text
Schema.Number().SafeParse("1.5")          -> invalid_type on every locale
Schema.DateTime().SafeParse("2024-05-01") -> invalid_type on every locale
```

The following rules keep validation semantics independent of the host locale:

- Type acceptance is decided by Variant-state classification (`IsNull`, `IsEmpty`, `IsError`, `IsObject`, `IsArray`) and `VarType`. User-provided Strings are never interpreted; `IsNumeric`, `Val`, and string-input conversion calls such as `CDbl(...)` or `CDate(...)` are not used to accept input.
- Conversion functions (`CDbl`, `CSng`, `CDec`, `CCur`, `CLng`, `CBool`, `CDate`) are applied only to values whose Variant type has already been confirmed. Numeric-to-numeric and Date-to-Date conversions do not consult the machine locale.
- `CStr` is applied only to String inputs, integral Long values, and late-bound member reads. It is never applied to a fractional number, because `CStr` renders the decimal separator per machine locale.
- `received` descriptors and `ErrorText` are locale-independent as defined in section 4: `Number(...)` uses `Str$` (always a `.` decimal separator), `Date(...)` uses explicit `yyyy-mm-dd` and `hh:nn:ss` `Format$` tokens, and `Error(<n>)` extracts the trailing numeric token from the Variant error text.
- String, field-name, and Literal/Enum comparisons use binary comparison (`StrComp(..., vbBinaryCompare)`).

If a future version adds an explicit coercion API, its locale policy must be defined as a separate public contract rather than silently inheriting host-locale parsing behavior.

## 13. v1 non-goals

The following capabilities are intentionally outside the v1 public contract. They are deferred design topics, not missing features; their absence is part of the stable v1 boundary.

- Implicit coercion of any kind
- An explicit coercion API (for example `Schema.Coerce.*`)
- `.Default()` or any other default-value injection
- Stripping unknown Object fields (`.Strict()` rejects them; validation never rewrites input)
- Arbitrary transformation or callback refinement (`.Transform()`, `.Refine()`, preprocess pipelines)
- `Partial`, `Pick`, `Omit`, or other Object-shape derivation helpers
- `Clone` or `Freeze`; schemas remain mutable and child schemas remain shared by reference as defined in section 8
- Int64/LongLong or Decimal Variant parsing and acceptance APIs
- `DateOnly` schemas and UTC/local timezone conversion
- Discriminated Union optimization
- A `Parse` API that raises on validation failure
- JSON parsing, HTTP requests, and worksheet reading
- Schema inference and code generation (for example OpenAPI generation)
