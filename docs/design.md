# VBA-Schema Design Specification

The normative specification for the public API, validation semantics, error contract, compatibility, and distribution boundary is `docs/specs/v1-contract.md`. This document explains the concept, structure, and implementation approach.

## 1. Overview

### 1.1 Project name

Working name:

**VBA-Schema**

A lightweight schema validation library for VBA.

### 1.2 Concept

VBA-Schema is a declarative runtime validation library for VBA inspired by TypeScript's Zod.

The goal is to define schemas and validate structure, types, and constraints for data such as the following.

* `Variant`
* `String`
* Numeric types
* `Boolean`
* `Date`
* `Collection`
* Arrays
* `Scripting.Dictionary`
* Dictionary / Collection structures produced by a JSON parser
* Values obtained from Excel cells or external APIs

Typical use cases:

```vb
Dim UserSchema As VSchema

Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text().Min(1).Max(100)) _
    .Field("age", Schema.Number().WholeNumber().Min(0)) _
    .Field("email", Schema.Text().Email()) _
    .Field("tags", Schema.ArrayOf(Schema.Text()).OptionalField())

Dim Result As VValidationResult
Set Result = UserSchema.SafeParse(UserData)

If Not Result.Success Then
    Debug.Print Result.ErrorText
End If
```

---

# 2. Design Goals

## 2.1 Primary goals

VBA-Schema prioritizes the following:

### Small distribution footprint

The distribution must contain no more than the following three modules.

```text
Schema.bas
VSchema.cls
VValidationResult.cls
```

**Never add a fourth VBA module.**

New schema types, constraints, error types, and similar features must also be implemented inside these three modules.

### No external references required

Do not require references to external libraries.

Use late binding where necessary.

For example:

```vb
CreateObject("Scripting.Dictionary")
CreateObject("VBScript.RegExp")
```

Users must not be required to:

```text
Microsoft Scripting Runtime
Microsoft VBScript Regular Expressions
```

This does not mean that late-bound runtime components are guaranteed to exist or that the library has zero dependencies. Environment wording follows `docs/specs/v1-contract.md`.

### Easy installation

An ordinary VBA user must be able to use the library by importing three files.

Neither a package manager nor xlflow may be required for usage.

### Fluent API

Schemas must be definable declaratively through fluent chains.

```vb
Schema.Text().Min(1).Max(100)
```

### Useful errors

The result must provide more than a Boolean:

* where the failure occurred,
* what was expected,
* what was actually received,
* and why it failed.

These details must be available to the caller.

---

# 3. Non-goals

The following are out of scope for v1.

## 3.1 Compile-time type system

This is not a static type-checking system for VBA.

VBA-Schema is a runtime validation library.

## 3.2 Complete Zod API compatibility

The Zod API is not reproduced verbatim.

In particular, v1 does not implement the following.

* `refine`
* `superRefine`
* Arbitrary callback validators
* `transform`
* preprocess pipeline
* discriminated union optimization
* recursive/lazy schema
* branded types
* Promise / async
* VBA type generation through schema inference

## 3.3 Class-per-schema architecture

The following designs are prohibited.

```text
VStringSchema.cls
VNumberSchema.cls
VObjectSchema.cls
VArraySchema.cls
VUnionSchema.cls
VValidationIssue.cls
```

Represent all schema kinds in `VSchema.cls`.

---

# 4. Hard Architecture Constraint

## Decision summary: Maximum Three VBA Modules

Record the rationale and compatibility trade-offs in `docs/adr/ADR-0001-small-distribution-and-portable-core.md`.

The distribution contains only:

```text
Schema.bas
VSchema.cls
VValidationResult.cls
```

Reasons:

1. In VBA, module count directly affects the installation experience.
2. Libraries with many files are difficult to add to an existing workbook.
3. Copying, removing, and updating the library should be easy.
4. The appeal of a small library should be preserved.
5. The installation experience should be as portable and easy to bring in as libraries such as VBA-JSON.

This constraint takes priority over internal architectural elegance.

To preserve code quality, private procedures inside `VSchema.cls` are still divided clearly by responsibility.

---

# 5. Module Architecture

```text
┌───────────────────────┐
│      Schema.bas       │
│                       │
│ Public factory API    │
└───────────┬───────────┘
            │ creates
            ▼
┌───────────────────────┐
│      VSchema.cls      │
│                       │
│ Schema definition     │
│ Constraint storage    │
│ Recursive validation  │
│ Type handling         │
└───────────┬───────────┘
            │ produces
            ▼
┌───────────────────────────┐
│ VValidationResult.cls     │
│                           │
│ Success                   │
│ Value                     │
│ Issues                    │
│ ErrorText                 │
└───────────────────────────┘
```

---

# 6. Public API

## 6.1 Schema.bas

`Schema.bas` exposes only the user-facing factory API.

It has no Public state.

### Required factories

```vb
Schema.AnyValue()
Schema.Text()
Schema.Number()
Schema.Bool()
Schema.DateTime()
Schema.ObjectSchema()
Schema.ArrayOf(ItemSchema)
Schema.Literal(Value)
Schema.EnumOf(Values)
Schema.UnionOf(Schemas)
```

These names are the fixed v1 candidates chosen to avoid reserved words and are verified by a compile-only fixture that uses every Public API. They passed VBE compilation on Windows 64-bit Office as of 2026-09-21. See `docs/adr/ADR-0002-vba-safe-api-and-error-boundary.md` for the rationale and `docs/specs/v1-contract.md` for signatures and fluent chains.

---

# 7. VSchema

## 7.1 Responsibility

`VSchema` is responsible for:

* storing the schema kind
* storing modifiers
* storing constraints
* storing child schemas
* storing Object fields
* running validation
* nested validation
* creating errors

---

# 8. Schema Kinds

Internally, one Enum represents the schema type.

```vb
Private Enum SchemaKind
    skAny = 0
    skString
    skNumber
    skBoolean
    skDate
    skObject
    skArray
    skLiteral
    skEnum
    skUnion
End Enum
```

The Enum is not exposed publicly.

Because a factory cannot call a Private initializer, use a Public but unsupported internal initialization API. Pass an internally encoded kind generated only by `Schema.bas`, not an ordinary schema-kind value.

Example:

```vb
Public Function InternalInitialize(ByVal KindCode As Long) As VSchema
```

This procedure may be called only once. A normal direct call or invalid encoded kind raises `vbObjectError + 2100`; reinitialization raises `vbObjectError + 2102`; and validation of a directly created, uninitialized schema raises an Err as programmer misuse. This is not a security boundary and is not part of the stable user-facing Public API.

Because VBA cannot access a Private member through another instance of the same class, the `Internal*` hooks needed for child-schema composition, path-aware nested validation, and schema-graph preflight are also Public. They are unsupported internal-only APIs and are not part of the stable user-facing API.

---

# 9. Schema State

`VSchema` broadly stores the following state.

```vb
Private mKind As Long

Private mOptional As Boolean
Private mNullable As Boolean

Private mHasMin As Boolean
Private mMinValue As Variant

Private mHasMax As Boolean
Private mMaxValue As Variant

Private mHasLength As Boolean
Private mLengthValue As Long

Private mIntegerOnly As Boolean
Private mEmail As Boolean

Private mHasPattern As Boolean
Private mPattern As String
Private mHasLiteral As Boolean
Private mLiteralValue As Variant
Private mEnumValues As Variant
Private mPatternRegExp As Object
Private mEmailRegExp As Object

Private mFields As Object
Private mFieldOrder As Collection
Private mItemSchema As VSchema

Private mUnionSchemas As Collection

Private mObjectMode As Long
```

The implementation need not match this list exactly, but **do not create a dedicated class for each constraint.**

---

# 10. Fluent Modifiers

## 10.1 OptionalField

```vb
Schema.Text().OptionalField()
```

Allow a missing field.

This has meaning only when the schema is used as an Object field.

Do not make it ambiguous whether `Empty` is Optional for standalone values.

In v1:

```text
Missing field ≠ Empty
```

use the following rule.

Allow Optional only when the key itself is absent from the Dictionary.

---

## 10.2 Nullable

```vb
Schema.Text().Nullable()
```

Allow `Null`.

---

## 10.3 Min

String:

```vb
Schema.Text().Min(3)
```

The minimum string length.

Number:

```vb
Schema.Number().Min(0)
```

The minimum numeric value.

Array:

It may apply to array element counts in the future, but it is allowed as a v1 implementation target.

---

## 10.4 Max

The same as `Min`.

---

## 10.5 Length

```vb
Schema.Text().Length(8)
```

An exact string length.

Use on Arrays may also be allowed.

---

## 10.6 WholeNumber

```vb
Schema.Number().WholeNumber()
```

Allow only integers.

The following are true:

```text
1
0
-10
```

The following are false:

```text
1.5
10.01
```

Allow a value that is mathematically integral even when its VBA type is `Double`.

---

## 10.7 Positive / Negative

Treat this as a future candidate that is not public in v1 and cannot be considered before v1.1.

If implemented:

```vb
.Positive()   ' > 0
.NonNegative() ' >= 0
.Negative()   ' < 0
.NonPositive() ' <= 0
```

It may be translated into existing `Min`/`Max` validation.

---

## 10.8 Pattern

```vb
Schema.Text().Pattern("^[A-Z]{3}-\d{4}$")
```

Use `VBScript.RegExp` through late binding. Fix `IgnoreCase=False`, `Global=False`, and `MultiLine=False`, and succeed only when the Expression matches the entire input.

Create the RegExp instance during validation or in a lazy cache rather than at schema construction. An invalid Expression raises an Err as programmer misuse at the first validation; a missing runtime component is an environment failure.

---

## 10.9 Email

```vb
Schema.Text().Email()
```

Email validation does not aim for complete RFC compliance. Use a simple ASCII form with one `@` separating local and domain and a total length of at most 254 UTF-16 code units; reject empty values, local-only values, multiple `@` characters, whitespace, Unicode, quoted local parts, comments, and IP literals.

The purpose is to detect common input mistakes, not to provide an excessively strict RFC implementation.

---

# 11. Object Schema

## 11.1 Definition

```vb
Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text()) _
    .Field("age", Schema.Number().OptionalField())
```

Internally, store field names and `VSchema` instances in a Dictionary.

```vb
Private mFields As Object
```

Store field declaration order in a separate `Collection` so validation Issue order does not depend on Dictionary enumeration order.

Construction:

```vb
Set mFields = CreateObject("Scripting.Dictionary")
```

---

## 11.2 Input

Object input officially supported in v1:

```text
Scripting.Dictionary
```

This includes late-bound Dictionaries.

Do not treat a `Collection` as a key-value Object.

Do not perform reflection on arbitrary VBA class instances.

---

# 12. Unknown Field Policy

Object schemas control how unknown fields are handled.

## Default

```text
Passthrough
```

Validation succeeds even when fields absent from the schema are present.

## Strict

```vb
Schema.ObjectSchema() _
    .Field("name", Schema.Text()) _
    .Strict()
```

An undefined field produces a validation error.

## Strip

Treat this as a future candidate that is not implemented in v1.

Changing the input data would make mutation semantics complex.

---

# 13. Array Schema

```vb
Schema.ArrayOf(Schema.Text())
```

Supported inputs:

* VBA native array
* `Collection`

When possible, also support Collections produced by JSON parsers.

Do not treat a Dictionary as an array.

---

# 14. Literal Schema

```vb
Schema.Literal("active")
```

Avoid the ambiguous coercion of VBA Variant comparison. Validate the scalar category of `Literal` and `EnumOf` candidates at construction, and snapshot the contents of the one-dimensional array for `EnumOf`.

Treat the following as distinct categories by default.

```text
"1"
1
True
```

Literal comparison also considers type.

---

# 15. Enum Schema

Recommended API:

```vb
Schema.EnumOf(Array("pending", "active", "disabled"))
```

Validation succeeds when the received value matches any candidate.

Comparison semantics are the same as Literal. Reject duplicate candidates as programmer misuse.

---

# 16. Union Schema

In v1, do not make an `Array()` containing schema Objects part of the public contract; use a `Collection`.

```vb
Dim Options As New Collection

Options.Add Schema.Text()
Options.Add Schema.Number()

Set S = Schema.UnionOf(Options)
```

`UnionOf` snapshots the branch order of the input `Collection` at construction. It does not clone each `VSchema` instance; it keeps references to child schemas, so child-builder changes after construction also affect the Union. An empty Collection, `Nothing`, a non-`VSchema` element, or an uninitialized `VSchema` raises `vbObjectError + 2100` as programmer misuse.

Union validation tries branches in order and succeeds when any branch succeeds. Do not flatten nested Unions; validate recursively while preserving their structure. When every branch fails, do not expose branch-specific Issue details; return one `invalid_union` Issue at the failed path. Programmer or environment errors inside a branch are propagated instead of being converted into validation failures.

---

# 17. Validation API

## 17.1 SafeParse

Basic API.

```vb
Dim Result As VValidationResult
Set Result = UserSchema.SafeParse(Data)
```

Do not raise a VBA runtime error for ordinary validation failure.

---

## 17.2 Parse

Treat this as a future candidate that is not implemented in v1.

```vb
Value = UserSchema.Parse(Data)
```

Raise `Err` on validation failure.

Because VBA Object/Variant return semantics are complex, the v1 validation API consists only of `SafeParse`. Reconsider `Parse` no earlier than v1.1.

---

# 18. VValidationResult

## 18.1 Public API

Provide at least the following.

```vb
Result.Success
Result.Value
Result.Issues
Result.ErrorText
```

### Success

```vb
Public Property Get Success() As Boolean
```

### Value

Return the original validated value.

Because v1 performs no transforms, this is normally the input value itself.

Do not deep-copy Objects either.

Return `Empty` for validation failure.

### Issues

Return a `Collection`.

Represent each Issue as a late-bound `Scripting.Dictionary`.

Do not create a dedicated `VValidationIssue.cls`.

Return snapshots so caller mutations cannot affect the internal Result.

---

# 19. Validation Issue Structure

Each Issue must contain at least the following.

```text
path
code
message
expected
received
```

Example:

```text
path     = "$.users[2].email"
code     = "invalid_email"
message  = "Invalid email address"
expected = "email"
received = "foo"
```

`received` is a safe String summary rather than an Object reference or the array itself. Issue types, order, and snapshot semantics follow `docs/specs/v1-contract.md`.

When a required field is absent, use `Missing` as `received`. This is distinct from a present `Empty` or `Null`, and `OptionalField` allows only missing fields.

In v1, the spec fixes the grammar of `received` and `ErrorText`. Implementations must not vary String escaping/truncation, locale-independent Number formatting, Date precision, or Issue line breaks.

Dictionary example:

```vb
Issue("path")
Issue("code")
Issue("message")
Issue("expected")
Issue("received")
```

---

# 20. Error Codes

Always provide a machine-readable code rather than relying only on an error-message string.

Initial codes:

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

Publish codes in the README as part of the public contract to avoid future breaking changes.

---

# 21. Path Format

Use the following format for nested error locations.

Every path starts with the root symbol `$`.

Object:

```text
$.user.name
```

Array:

```text
$.users[2]
```

Nested:

```text
$.users[2].address.zip
```

Field names that are not identifiers:

```text
$["field.with.dot"]
```

Arrays and Collections normalize enumeration order to zero-based logical indexes regardless of the actual `LBound` or one-based index.

v1 accepts only one-dimensional native arrays. A multidimensional array produces `invalid_array_rank`.

v1 does not introduce a dedicated path-segment class.

---

# 22. ErrorText

`VValidationResult.ErrorText` returns a human-readable representation.

Use the format `<path>: <message> (expected=<expected>, received=<received>)` and join Issues with `vbCrLf`. Do not add the complete raw input or secrets.

Example:

```text
$.users[2].email: Invalid email address (expected=email, received=String("foo"))
$.users[2].age: Number is below the minimum. (expected=Number >= 18, received=Number(16))
$.settings.timeout: The value has an invalid type. (expected=Number, received=String("slow"))
```

Fix the format with tests.

The Issue Collection is the official machine-readable API; `ErrorText` is a presentation API.

---

# 23. Type Semantics

Define this area precisely because VBA has type ambiguities.

## String

In general:

```vb
VarType(Value) = vbString
```

are the only accepted forms.

Do not automatically convert numbers to Strings.

---

## Number

Treat the following as Number:

```text
Byte
Integer
Long
Single
Double
Currency
```

Do not accept Decimal Variants or LongLong in v1 until the same contract can be verified on both 32-bit and 64-bit hosts.

Number type acceptance is independent of whether a value can be converted to Decimal. A finite Double outside the Decimal range is accepted by unconstrained Number validation. Keep constraint boundaries as Variants and do not unconditionally convert them to Double.

Run numeric comparison only for `Min`/`Max` constraints or `WholeNumber` checks. Use subtype-aware lossless widening and necessary round-trip checks. Never silently accept precision loss, overflow, NaN, or Infinity. Follow the fixed spec order and add only the first Issue for each node.

Do not treat Boolean as numeric.

Date is also numeric internally but is not treated as Number.

---

## Boolean

Only `vbBoolean`.

Do not automatically convert `0` or `-1` to Boolean.

---

## Date

Only `vbDate`.

Do not automatically parse date Strings.

---

## Null

Allow only when Nullable is specified.

---

## Empty

Keep it distinct from Null.

For ordinary schemas in v1, treat it as a validation failure.

Allow only `AnyValue()`.

---

## Error Variant

Allow it for `AnyValue()` and treat it as a validation failure for every other schema.

---

# 24. No Implicit Coercion

This is a core VBA-Schema design principle.

Never automatically convert the following:

```text
"123" → Number
123 → String
0 → Boolean
"2026-01-01" → Date
```

Reason:

Implicit conversion in a validation library hides invalid values at data boundaries.

If a coercion API is added in the future, it must be explicit.

Example:

```text
Schema.Coerce.Number()
```

This remains out of scope for v1.

---

# 25. Validation Algorithm

Conceptually:

```text
Validate(schema, value, path)

1. Null / Empty / Optional handling
2. schema kind dispatch
3. Basic type validation
4. schema-specific constraint
5. nested validation
6. Add to the Issue Collection
```

pseudo code:

```vb
Private Sub ValidateValue( _
    ByVal Value As Variant, _
    ByVal Path As String, _
    ByRef Issues As Collection)

    If HandleNullability(...) Then Exit Sub

    Select Case mKind
        Case skAny
            Exit Sub

        Case skString
            ValidateString Value, Path, Issues

        Case skNumber
            ValidateNumber Value, Path, Issues

        Case skObject
            ValidateObject Value, Path, Issues

        Case skArray
            ValidateArray Value, Path, Issues

        ' ...
    End Select
End Sub
```

Do not implement this as one enormous procedure.

Split validation into private procedures by kind.

---

# 26. Object Validation Algorithm

```text
ValidateObject(value, path):

1. Check `TypeName(value) = "Dictionary"` and the limited `Count`/`Exists`/`Keys` capabilities. On failure, add an `invalid_type` Issue.
2. Enumerate schema fields.
3. Enumerate `Keys` and use `StrComp(..., vbBinaryCompare)` to check whether each key exists in the input.
4. If absent, evaluate Optional.
5. If present, validate recursively with the child schema.
6. Add non-String input keys as `invalid_key` at the root Object path regardless of `.Strict()`.
7. For Strict, enumerate String input keys in binary ordinal order.
8. Add keys absent from the schema as `unknown_field`.
```

---

# 27. Array Validation Algorithm

```text
ValidateArray(value, path):

1. Determine whether the value is a VBA array or Collection; do not treat a Dictionary as an array.
2. Treat an uninitialized dynamic array as zero elements and a second or later dimension as `invalid_array_rank`.
3. Validate length constraints.
4. Run the child schema recursively for each element.
5. Add a zero-based logical `[index]` to the path.
```

Use `LBound` / `UBound` for native VBA arrays.

Take care not to raise a runtime error for empty or uninitialized dynamic arrays.

Implement a safe array-detection helper for this case. Clear the Err state used during probing inside the helper, then return to normal validation. Do not modify the input array or Collection.

---

# 28. Object Detection

Because VBA has no interface reflection, v1 explicitly targets Dictionaries.

The v1 check combines the following entry point with limited capability checks.

```vb
TypeName(Value) = "Dictionary"
```

If reading `Count`, `Exists`, or `Keys` fails, report `invalid_type` for the input. If creating the Dictionary used for Issues fails, propagate it as a runtime/environment error (`vbObjectError + 2200`).

Do not use exception-driven duck typing that accepts an Object and probes arbitrary members. Use only limited helpers after TypeName. Do not change the input Dictionary's `CompareMode`; key comparison is always binary.

Non-String keys are `invalid_key`; strict unknown String keys are `unknown_field`. Unknown fields pass through by default, and validation does not modify the input.

---

# 28.1 Schema graph preflight

Builders are mutable and schemas can be shared, so inspect the schema graph with DFS at the start of `SafeParse`. Store the active path in a `Collection` and compare schema identity with VBA `Is`. Raise `vbObjectError + 2103` as soon as a direct or indirect cycle is found, before value validation or Issue creation begins. Allow non-cyclic reuse of shared children. Do not use `ObjPtr`, the Windows API, or pointer-size dependencies.

---

# 29. Performance Requirements

It must be fast enough for ordinary VBA validation workloads.

Initial targets are measured on the same Windows 64-bit machine and workbook using the median of five runs after warm-up:

```text
1,000 scalar fields:
Less than 500 ms

10,000 scalar validations:
Less than 1,000 ms
```

Because environments vary widely, do not judge regressions by absolute time alone. Investigate a degradation greater than 25% against the established baseline for the same fixture.

Prioritize the following over micro-optimization.

* Do not use runtime errors as validation flow control.
* Avoid unnecessary RegExp recreation.
* Use Dictionary lookup effectively.
* Avoid unnecessary deep copies.
* Create error Objects only on failure.

---

# 30. Mutation Policy

Do not modify input in v1.

The following are prohibited.

```text
Removing unknown fields
string trim
type coercion
Injecting default values
Rebuilding Objects
```

Therefore:

```vb
Result.Value
```

normally returns the input value itself.

---

# 31. Error Handling

Distinguish validation failures from library bugs.

Use the complete classification and Err ranges in `docs/specs/v1-contract.md`.

## Validation failure

Do not raise an Err from `SafeParse` for ordinary validation failures.

Result:

```text
Success = False
Issues.Count > 0
```

## Programmer misuse

Example:

```vb
Schema.Text().Min(-1)
Schema.ArrayOf(Nothing)
.Field("", Nothing)
```

This is API misuse, so it may raise `Err.Raise`.

Define dedicated error-number ranges.

Example:

```vb
vbObjectError + 2100
```

Use `vbObjectError + 2200` through `vbObjectError + 2299` for missing runtime components and internal invariant violations, and do not convert them into validation failures even in `SafeParse`.

---

# 32. Fluent API Mutation Semantics

Builder methods normally mutate and return the same instance.

```vb
Public Function Min(ByVal Value As Variant) As VSchema
    mHasMin = True
    mMinValue = Value
    Set Min = Me
End Function
```

This makes the following possible:

```vb
Schema.Number().Min(0).Max(100).WholeNumber()
```

the fluent builder contract.

Do not use an immutable-schema design because it would increase allocation and implementation size in VBA.

---

# 33. Schema Reuse Warning

Because builders are mutable, explain the following behavior in the README.

```vb
Dim Base As VSchema
Set Base = Schema.Text()

Dim A As VSchema
Set A = Base.Min(1)
```

`A` and `Base` are the same Object.

An immutable clone API is out of scope for v1.

---

# 34. Documentation Example

The beginning of the README should emphasize low installation cost.

```text
Modern schema validation for VBA.

3 files.
No reference setup.
No compile-time dependencies.
```

Example:

```vb
Dim UserSchema As VSchema

Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text().Min(1)) _
    .Field("age", Schema.Number().WholeNumber().Min(0)) _
    .Field("email", Schema.Text().Email())

Dim Result As VValidationResult
Set Result = UserSchema.SafeParse(Data)

If Result.Success Then
    Debug.Print "Valid"
Else
    Debug.Print Result.ErrorText
End If
```

The VBE import payload contains only the three production files; provide README, LICENSE, CHANGELOG, sample workbooks, ZIP files, and other assets separately. A GitHub Release starts from a `vMAJOR.MINOR.PATCH` tag push, runs `source-check` first, and stores only the three files under `VBA-Release/` in `VBA-Release-vX.Y.Z.zip`. Prerelease tags are outside the current workflow. Staged VBA source is UTF-8 without a BOM and uses LF line endings. The Git tag and CHANGELOG are authoritative for versioning; do not add a version constant to VBA source.

---

# 35. Compatibility

Development and verification target:

```text
VBA7
Office 2016+
Windows 64-bit Office
```

Windows 32-bit Office and macOS Office are compatibility-conscious but unverified and unsupported because no verification environments are available. Do not describe them as supported or verified.

Windows is the primary target.

To avoid unnecessarily losing compatibility with unverified environments, prohibit core dependencies on the Windows API, pointer-size-specific declarations, and the Excel Object Model.

Late-bind `Scripting.Dictionary` and `VBScript.RegExp`, but do not guarantee their availability in unverified environments.

When a required runtime component is unavailable, report an environment failure rather than a validation failure.

---

# 36. Testing Strategy

Tests must be executable automatically through xlflow.

At minimum, provide the following categories.

```text
tests/
├── TestString.bas
├── TestNumber.bas
├── TestBoolean.bas
├── TestDate.bas
├── TestObject.bas
├── TestArray.bas
├── TestOptional.bas
├── TestNullable.bas
├── TestLiteral.bas
├── TestEnum.bas
├── TestUnion.bas
├── TestErrors.bas
└── TestIntegration.bas
```

The number of test modules is not counted against the three-module distribution constraint.

---

# 37. Required Test Cases

## Strings

```text
valid string
wrong type
Min boundary
Max boundary
Length
empty string
Pattern pass/fail
Email pass/fail
Null
Empty
length boundary accepts `3`, `3&`, `CByte(3)`, `3#`
length boundary rejects fraction / negative / overflow / non-numeric
```

## Numbers

```text
Byte
Integer
Long
Single
Double
Currency
minimum
maximum
integer
fraction
Boolean rejection
Date rejection
numeric string rejection
Number Min/Max boundary subtype coverage
```

## Object

```text
required field present
required field missing
optional field missing
nested object
multiple errors
strict unknown field
dictionary input
input Dictionary with `CompareMode = vbTextCompare` still uses binary field matching
wrong input type
```

## Array

```text
native array
Collection
empty array
nested array
wrong element
correct error index
```

## Literal

```text
correct value
wrong value
same textual representation but different type
```

## Union

```text
first branch success
later branch success
all branches fail
nested union
```

## Error path

```text
$
$.user.name
$.users[0]
$.users[2].email
$.orders[1].items[4].price
```

---

# 38. Regression Testing

Add a reproduction test for every bug fix.

For a fix that changes validation semantics, always check its impact on existing tests.

An AI agent must not change only production code when fixing a bug.

---

# 39. Code Quality Requirements

All modules:

```vb
Option Explicit
```

are required.

Avoid the following.

```text
Select
Activate
Broad use of `On Error Resume Next`
Implicit Variants
Public mutable field
unqualified Excel object references
Windows API
```

Direct dependence on the Excel Object Model is prohibited by default.

VBA-Schema is designed as a VBA runtime library rather than an Excel-only library.

---

# 40. xlflow Dogfooding

Use xlflow for development.

At minimum, CI or local development must run:

```text
xlflow test
xlflow lint
xlflow fmt
```

and pass these checks.

Use the VBE compile oracle when available.

Verify that API examples actually compile in Excel/VBE.

---

# 41. Repository Structure

Canonical layout:

```text
vba-schema/
├── build/
│   └── Book.xlsm
├── src/
│   ├── modules/
│   │   ├── Schema.bas
│   │   ├── Tests/
│   │   └── Xlflow/
│   ├── classes/
│   │   ├── VSchema.cls
│   │   └── VValidationResult.cls
│   └── workbook/
│
├── sample/
│   ├── 01-order-import/
│   │   ├── README.md
│   │   └── SampleOrderImport.bas
│   ├── 02-settings-validation/
│   │   ├── README.md
│   │   └── SampleSettingsValidation.bas
│   ├── 03-api-payload/
│   │   ├── README.md
│   │   └── SampleApiPayload.bas
│   └── README.md
│
├── docs/
│   ├── adr/
│   ├── specs/
│   └── design.md
│
├── dist/
│   └── VBA-Schema/
│       ├── Schema.bas
│       ├── VSchema.cls
│       └── VValidationResult.cls
│
├── README.md
├── LICENSE
└── THIRD_PARTY_NOTICES.md
```

If there are no dependencies, `THIRD_PARTY_NOTICES.md` is unnecessary.

`src/` is for development and tests, `build/Book.xlsm` is the tracked xlflow development fixture with the current source imported, `sample/` contains supporting user examples, and `dist/VBA-Schema/` is the user distribution containing only three files. Do not generate a three-component workbook directly from the complete xlflow source tree; create the release-staging project from the three-file allowlist. Do not mix `.bas` files from `sample/` into the distribution payload.
The release-staging Destination must not overlap the repository root, `src`, `.git`, `.xlflow`, or `build`; `tools/release-stage.ps1` rejects the overlap before deleting an existing Destination.

---

# 42. Implementation Phases

## Phase 0 — Contract and compile fixture

Implementation:

```text
Public API compile-only fixture
VBE compile oracle
Fix the error/path/type contract.
Design three-file release staging.
Provide an allowlist staging script or Task.
```

Acceptance criteria:

* `docs/specs/v1-contract.md` matches the Public signatures.
* Every factory and fluent chain compiles on Windows 64-bit Office.
* The boundary between development source, verification workbook, and three-file distribution is explicit.
* A script or Task can stage only three files and run the release gate.

---

## Phase 1 — Core infrastructure

Implementation:

```text
Schema.bas
VSchema.cls
VValidationResult.cls
```

Target schemas:

```text
AnyValue
Text
Number
Bool
DateTime
```

Target APIs:

```text
SafeParse
OptionalField
Nullable
Min
Max
Length
WholeNumber
```

Acceptance criteria:

* Compilation succeeds.
* Scalar validation tests succeed.
* error path `$`
* zero external references
* exactly 3 distributed modules

---

## Phase 2 — Object validation

Implementation:

```text
Object
Field
nested Object
Optional field
Strict
```

Acceptance criteria:

```vb
Schema.ObjectSchema() _
    .Field("name", Schema.Text()) _
    .Field("age", Schema.Number().OptionalField())
```

works as specified.

nested path:

```text
$.user.address.zip
```

is generated correctly.

---

## Phase 3 — Arrays

Implementation:

```text
ArrayOf
native VBA arrays
Collection
nested arrays
```

path:

```text
$.users[3].email
```

can be generated.

---

## Phase 4 — Value constraints

Implementation:

```text
Literal
Enum
Pattern
Email
```

---

## Phase 5 — Union

Implementation:

```text
UnionOf
```

Organize error output.

---

## Phase 6 — Documentation and hardening

Perform:

```text
README
API reference
Examples
edge cases
64-bit Office validation
32-bit Office compatibility review (validate if hardware is available)
macOS compatibility review (validate if hardware is available)
performance benchmark
```

Benchmarks are development support artifacts; `src/modules/Benchmarks` and
`tools/run-benchmark.ps1` are not included in the three-file import payload. Windows 64-bit
Excel fixture results are compared with the tracked baseline. macOS and Windows 32-bit
remain compatibility-conscious but unverified and unsupported. Detailed fixtures, counters,
thresholds, and session ownership are defined authoritatively by `docs/specs/benchmark-contract.md` and
`docs/adr/ADR-0008-runtime-benchmark-contract.md`. The benchmark
does not compare or update the baseline when `office_bitness` is not `x64`; it writes an unsupported
report and fails.

---

# 43. MVP Definition

Minimum v1.0 deliverables:

```text
AnyValue
Text
Number
Bool
DateTime
ObjectSchema
ArrayOf
Literal
EnumOf
UnionOf

OptionalField
Nullable

Min
Max
Length
WholeNumber
Pattern
Email

Field
Strict

SafeParse

Success
Value
Issues
ErrorText
```

---

# 44. v1.0 Exclusions

Explicitly excluded from v1.0.

```text
Transform
Refine
SuperRefine
Callback validator
Default values
Coercion
Strip unknown fields
Recursive schemas
Class object reflection
OpenAPI generation
JSON parsing
HTTP
Schema serialization
Code generation
```

Consider these only after the core API is stable.

---

# 45. Future Extensions

## OpenAPI integration

In the future:

```text
OpenAPI Schema
      ↓
generated VBA-Schema definition
      ↓
VBA-HTTP response
      ↓
runtime validation
```

will be possible.

Example:

```vb
Set User = UserSchema.SafeParse(Response.Json)
```

Keep an OpenAPI generator as a separate project or build-time tool so it does not enlarge VBA-Schema itself.

---

# 46. Design Principles for AI Agents

AI agents must follow these rules during implementation.

### 1. Do not add modules

Do not add new `.bas` / `.cls` / `.frm` files to production source.

### 2. Prefer private procedures over classes

If responsibilities must be separated:

```text
Private ValidateString
Private ValidateNumber
Private ValidateObject
Private ValidateArray
Private AddIssue
Private BuildPath
```

split procedures in this way.

### 3. Do not over-engineer

Do not introduce the following.

```text
dependency injection
interface hierarchy
visitor pattern
factory classes
constraint classes
error classes
schema subclasses
```

### 4. Tests before extensions

When adding a feature:

```text
1. failing test
2. minimal implementation
3. full tests
4. lint
5. compile validation
```

use this order by default.

### 5. Preserve strict semantics

Do not add implicit type coercion merely because it is convenient.

### 6. Public API stability matters

Treat Public method-name and error-code changes as breaking changes.

---

# 47. Definition of Done

v1.0 is complete only when all of the following are true.

* Distributed production source contains only three VBA components.
* No external reference setup is required.
* VBE compilation succeeds on Windows 64-bit Office.
* macOS / Windows 32-bit are explicitly described as compatibility-conscious, unverified, and unsupported.
* Text / Number / Bool / DateTime are supported.
* Object / nested Object are supported.
* Array / Collection are supported.
* OptionalField / Nullable are supported.
* Literal / Enum / Union are supported.
* Validation errors include paths.
* machine-readable error codes
* human-readable `ErrorText`
* Strict object validation
* Pattern / Email
* xlflow tests pass.
* xlflow lint passes.
* xlflow analyze passes.
* README contains installation, examples, and an API overview.
* Major Public APIs have regression tests.
* The distribution files can be imported manually and used.
* The release artifact contains only `Schema.bas`, `VSchema.cls`, and `VValidationResult.cls`.

---

# 48. Product Positioning

VBA-Schema's value is not that it completely reproduces Zod in VBA.

Its value is:

> Modern schema validation for VBA with almost zero installation cost.

to deliver this outcome.

The project must preserve these characteristics:

```text
3 files
no external reference setup
declarative schemas
strict runtime validation
structured errors
works with ordinary VBA projects
```

.

If a feature would remove these characteristics, do not add it to the core library.

**Small size is a feature, not an implementation detail.**
