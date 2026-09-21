# VBA-Schema

> **Declarative schema validation for VBA.**
> **3 files. No reference setup. No implicit coercion.**

VBA-Schema brings Zod-inspired runtime schema validation to ordinary VBA projects.

Instead of scattering `VarType`, `IsNull`, length checks, nested `If` statements, and ad-hoc error messages throughout your code, define the shape of your data once and validate it consistently.

```vb
Dim UserSchema As VSchema

Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text().Min(1).Max(100)) _
    .Field("age", Schema.Number().WholeNumber().Min(0)) _
    .Field("email", Schema.Text().Email()) _
    .Field("enabled", Schema.Bool().OptionalField()) _
    .Strict()
```

Then validate any value at runtime:

```vb
Dim result As VValidationResult
Set result = UserSchema.SafeParse(inputObject)

If result.Success Then
    Debug.Print "Valid"
Else
    Debug.Print result.ErrorText
End If
```

```text
$.email: Invalid email address. (expected=email, received=String("not-an-email"))
$.age: Number is below the minimum. (expected=Number >= 0, received=Number(-1))
```

**Inspired by Zod, designed for VBA — not a port of Zod.**

---

## Why VBA-Schema?

Data entering a VBA application rarely comes with guarantees.

It may come from an API, JSON parser, worksheet, CSV file, configuration file, database, another workbook, or a user-maintained macro.

Without a validation layer, code often ends up doing this everywhere:

```vb
If Not data.Exists("name") Then
    ' ...
End If

If VarType(data("name")) <> vbString Then
    ' ...
End If

If Len(data("name")) = 0 Then
    ' ...
End If

If Not data.Exists("age") Then
    ' ...
End If

If Not IsNumeric(data("age")) Then
    ' ...
End If
```

VBA-Schema moves those assumptions into a reusable schema:

```vb
Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text().Min(1)) _
    .Field("age", Schema.Number().WholeNumber().Min(0))
```

Your validation rules become explicit, composable, testable, and easy to read.

---

## Installation

VBA-Schema is intentionally distributed as only three VBA components.

1. Download the latest `VBA-Release-vX.Y.Z.zip` from GitHub Releases.
2. Import `Schema.bas`, `VSchema.cls`, and `VValidationResult.cls` into your VBA project.
3. Start defining schemas.

No additional reference setup is required.

```text
Schema.bas
VSchema.cls
VValidationResult.cls
```

That's the entire library.

---

## Quick Start

Create some input data:

```vb
Dim user As Object
Set user = CreateObject("Scripting.Dictionary")

user.Add "name", "Alice"
user.Add "age", 27
user.Add "email", "alice@example.com"
```

Define its schema:

```vb
Dim UserSchema As VSchema

Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text().Min(1).Max(100)) _
    .Field("age", Schema.Number().WholeNumber().Min(0)) _
    .Field("email", Schema.Text().Email()) _
    .Strict()
```

Validate it:

```vb
Dim result As VValidationResult
Set result = UserSchema.SafeParse(user)

If result.Success Then
    Debug.Print "User is valid"
Else
    Debug.Print result.ErrorText
End If
```

`SafeParse` does not raise an error for ordinary validation failures. It returns a `VValidationResult` containing the validation result and structured issues.

---

## Supported Schemas

| Schema  | Example                    | Accepts                                       |
| ------- | -------------------------- | --------------------------------------------- |
| Any     | `Schema.AnyValue()`        | Any Variant state                             |
| Text    | `Schema.Text()`            | `String`                                      |
| Number  | `Schema.Number()`          | Byte, Integer, Long, Single, Double, Currency |
| Boolean | `Schema.Bool()`            | `Boolean`                                     |
| Date    | `Schema.DateTime()`        | `Date`                                        |
| Object  | `Schema.ObjectSchema()`    | `Scripting.Dictionary`                        |
| Array   | `Schema.ArrayOf(...)`      | One-dimensional VBA arrays and `Collection`   |
| Literal | `Schema.Literal("active")` | One exact scalar value                        |
| Enum    | `Schema.EnumOf(...)`       | One of several scalar values                  |
| Union   | `Schema.UnionOf(...)`      | A value matching at least one schema          |

---

## Text Validation

Text schemas support length constraints, regular expressions, and email validation.

```vb
Dim CodeSchema As VSchema

Set CodeSchema = Schema.Text() _
    .Length(8) _
    .Pattern("[A-Z]{3}-[0-9]{4}")
```

```vb
Dim EmailSchema As VSchema
Set EmailSchema = Schema.Text().Email()
```

`Pattern` performs a case-sensitive full match using `VBScript.RegExp`.

`Email` is intentionally a practical ASCII email validator rather than a complete implementation of every RFC-valid email form.

---

## Number Validation

```vb
Dim AgeSchema As VSchema

Set AgeSchema = Schema.Number() _
    .WholeNumber() _
    .Min(0) _
    .Max(150)
```

VBA-Schema does not silently convert values.

```text
42       -> valid Number
42#      -> valid Number
"42"     -> invalid Number
True     -> invalid Number
#1/1/26# -> invalid Number
```

Validation is intended to detect incorrect data at system boundaries, not hide it through implicit conversion.

---

## Object Validation

Object schemas validate late-bound `Scripting.Dictionary` values.

```vb
Dim AddressSchema As VSchema

Set AddressSchema = Schema.ObjectSchema() _
    .Field("city", Schema.Text().Min(1)) _
    .Field("zip", Schema.Text().Pattern("[0-9]{5}"))

Dim UserSchema As VSchema

Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text()) _
    .Field("address", AddressSchema)
```

Nested validation automatically produces paths such as:

```text
$.address.zip
```

By default, unknown fields are allowed.

Use `.Strict()` when additional fields should be rejected:

```vb
Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text()) _
    .Strict()
```

An unexpected field then produces an `unknown_field` issue.

---

## Optional and Nullable Values

`OptionalField` and `Nullable` have intentionally different meanings.

```vb
Set SettingsSchema = Schema.ObjectSchema() _
    .Field("nickname", Schema.Text().OptionalField()) _
    .Field("description", Schema.Text().Nullable())
```

`OptionalField()` allows the field to be absent from an object.

`Nullable()` allows a present value to be `Null`.

A missing value, `Empty`, and `Null` are not treated as the same thing.

---

## Arrays and Collections

Use `ArrayOf` to validate sequences.

```vb
Dim TagsSchema As VSchema
Set TagsSchema = Schema.ArrayOf(Schema.Text().Min(1))
```

Both one-dimensional native VBA arrays and `Collection` objects are supported.

```vb
Dim tags As Collection
Set tags = New Collection

tags.Add "vba"
tags.Add "excel"
tags.Add "schema"

Dim result As VValidationResult
Set result = TagsSchema.SafeParse(tags)
```

Nested paths use zero-based logical indexes regardless of the original VBA array bounds.

```text
$.users[2].email
$.orders[1].items[4].price
```

Array schemas also support `Length`, `Min`, and `Max`.

```vb
Set TagsSchema = Schema.ArrayOf(Schema.Text()) _
    .Min(1) _
    .Max(10)
```

---

## Literal and Enum Schemas

Use `Literal` when a value must exactly match one scalar value.

```vb
Dim StatusSchema As VSchema
Set StatusSchema = Schema.Literal("active")
```

Literal comparison is type-aware.

```text
"1" <> 1
1 <> True
Date <> Double
```

Use `EnumOf` for a fixed set of allowed values:

```vb
Dim StatusSchema As VSchema

Set StatusSchema = Schema.EnumOf( _
    Array("pending", "active", "disabled") _
)
```

Enum candidates are copied when the schema is created, so changing the source array afterward does not alter the schema.

---

## Union Schemas

A union succeeds when any one branch matches.

```vb
Dim branches As Collection
Set branches = New Collection

branches.Add Schema.Text()
branches.Add Schema.Number()

Dim IdSchema As VSchema
Set IdSchema = Schema.UnionOf(branches)
```

Both of these are valid:

```vb
Set result = IdSchema.SafeParse("A-1024")
Set result = IdSchema.SafeParse(1024)
```

If every branch fails, VBA-Schema returns one `invalid_union` issue at that path instead of exposing a large collection of internal branch errors.

---

## Structured Validation Errors

`SafeParse` returns a `VValidationResult`.

| Property    | Description                         |
| ----------- | ----------------------------------- |
| `Success`   | `True` when validation succeeds     |
| `Value`     | Original input value on success     |
| `Issues`    | Structured validation issues        |
| `ErrorText` | Human-readable error representation |

Each issue is a `Scripting.Dictionary` containing:

```text
path
code
message
expected
received
```

Example:

```vb
Dim issues As Collection
Set issues = result.Issues

Dim issue As Object
Set issue = issues.Item(1)

Debug.Print issue("path")
Debug.Print issue("code")
Debug.Print issue("message")
Debug.Print issue("expected")
Debug.Print issue("received")
```

Example output:

```text
$.users[2].email
invalid_email
Invalid email address.
email
String("invalid")
```

Paths are deterministic and support nested objects, arrays, and field names that cannot use dot notation.

```text
$
$.user.name
$.users[0]
$["field.with.dot"]
```

`Issues` returns snapshots. Modifying a returned Collection or issue Dictionary does not modify the `VValidationResult`.

---

## Validation vs Programming Errors

VBA-Schema intentionally separates invalid input from invalid schema definitions.

This is a validation failure:

```vb
Set result = Schema.Number().SafeParse("hello")

Debug.Print result.Success
' False
```

This is a programming error:

```vb
Set schema = Schema.Number().Min(10).Max(5)
```

Invalid schema construction raises a VBA error instead of returning `Success = False`.

The same principle applies to missing runtime components and internal failures. `SafeParse` only converts actual validation failures into `VValidationResult`.

---

## No Implicit Coercion

VBA frequently performs automatic conversions. VBA-Schema deliberately does not.

```text
"123"       is not Number
123         is not Text
0           is not Boolean
"2026-01-01" is not Date
```

This makes schema validation useful at trust boundaries, where silently accepting the wrong data type can hide bugs.

---

## Common Use Cases

VBA-Schema works especially well at boundaries where VBA receives structured but untrusted data.

| Source              | Example                                                   |
| ------------------- | --------------------------------------------------------- |
| HTTP / REST API     | Validate parsed response objects                          |
| JSON                | Validate Dictionary / Collection structures after parsing |
| Configuration       | Validate application settings before use                  |
| Excel import        | Validate rows after mapping them into structured values   |
| CSV                 | Validate converted records                                |
| Databases           | Validate dynamically retrieved values                     |
| Inter-workbook data | Check assumptions before processing                       |
| User input          | Centralize complex validation rules                       |

For example, VBA-Schema can sit directly after an HTTP or JSON layer:

```text
HTTP / JSON
    ↓
Dictionary / Collection
    ↓
VBA-Schema
    ↓
trusted application logic
```

---

## Example: API Response Validation

```vb
Dim UserSchema As VSchema

Set UserSchema = Schema.ObjectSchema() _
    .Field("id", Schema.Number().WholeNumber()) _
    .Field("username", Schema.Text().Min(1)) _
    .Field("email", Schema.Text().Email()) _
    .Field("active", Schema.Bool()) _
    .Strict()

Dim result As VValidationResult
Set result = UserSchema.SafeParse(responseData)

If Not result.Success Then
    Debug.Print result.ErrorText
    Exit Sub
End If

' From here onward, application code can rely on the schema contract.
```

This is the main role of VBA-Schema:

> **Turn untrusted runtime data into an explicitly validated boundary.**

---

## Public API Overview

### Factories

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

### Modifiers

```vb
.OptionalField()
.Nullable()

.Min(...)
.Max(...)
.Length(...)

.WholeNumber()

.Pattern(...)
.Email()

.Field(...)
.Strict()
```

### Validation

```vb
Set result = schema.SafeParse(value)
```

The detailed and normative v1 behavior is documented in [`docs/specs/v1-contract.md`](docs/specs/v1-contract.md).

---

## Design Philosophy

VBA-Schema intentionally stays small.

The production library will not grow into a large hierarchy of schema classes, validation issue classes, constraint classes, adapters, and framework infrastructure.

The distribution boundary is deliberately fixed at:

```text
Schema.bas
VSchema.cls
VValidationResult.cls
```

Some features that would significantly increase complexity are intentionally outside the v1 scope, including transforms, arbitrary callback refinements, automatic coercion, class reflection, JSON parsing, OpenAPI generation, and schema code generation.

**Small size is a feature, not an implementation detail.**

---

## Compatibility

The currently verified environment is:

| Environment                 | Status                                    |
| --------------------------- | ----------------------------------------- |
| Windows 64-bit Office 2016+ | Verified                                  |
| Windows 32-bit Office 2016+ | Compatibility-conscious, not yet verified |
| macOS Office                | Compatibility-conscious, not yet verified |

The core does not depend on the Excel Object Model or Windows API.

`Scripting.Dictionary` is used for object validation and structured issues. `VBScript.RegExp` is used by `Pattern` and `Email`. Both are late-bound, so no VBA reference setup is required.

If a required runtime component is unavailable, VBA-Schema reports an environment error rather than treating it as invalid user data.

---

## Examples

Runnable examples are available in [`sample/`](sample/README.md).

| Example              | Demonstrates                                         |
| -------------------- | ---------------------------------------------------- |
| Order import         | Nested objects, arrays, Enum, Pattern, Email, Strict |
| Application settings | Enum, ranges, Pattern, OptionalField, Nullable       |
| API payload          | Union, Literal, nested objects, Collection           |

The samples focus on validation after external data has already been converted into VBA values.

VBA-Schema itself does not parse JSON, perform HTTP requests, or read worksheets.

---

## Development

VBA-Schema is developed and tested with [xlflow](https://github.com/harumiWeb/xlflow).

The repository includes automated linting, static analysis, public API compile fixtures, behavioral tests, release-payload verification, and Windows x64 benchmarks.

Development infrastructure and tests are not part of the three-file distribution.

See [`docs/`](docs/) for the full design, architecture decisions, and v1 contract.

---

## License

MIT License. See [`LICENSE`](LICENSE).

---

## Project Status

VBA-Schema is preparing for its first stable release.

The public v1 contract is being treated as a compatibility boundary, so API and validation semantics are intentionally being stabilized before `v1.0.0`.
