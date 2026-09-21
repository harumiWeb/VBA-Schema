# Application settings validation

This sample validates values read from settings sheets or forms before the application starts using them.

## Run

- Import: `SampleSettingsValidation.bas`
- Macro: `RunSettingsValidationSample`

When a setting is missing, `OptionalField` and `Nullable` express different meanings:

- `OptionalField`: allows the Dictionary key to be absent.
- `Nullable`: allows a present key whose value is `Null`.

## APIs used

- `ObjectSchema`, `Field`, `Strict`
- `EnumOf`, `Pattern`
- `Number().WholeNumber().Min().Max()`
- `ArrayOf(Text())`
- `OptionalField`, `Nullable`

The failing input demonstrates results such as:

```text
$.environment  invalid_enum
$.apiBaseUrl   invalid_pattern
$.port         too_small
$.tags[1]      invalid_type
$.debugMode    unknown_field
```
