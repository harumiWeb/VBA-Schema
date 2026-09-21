# Order and line-item import

This sample validates order data before import, after an Excel table or CSV has been converted into `Scripting.Dictionary` and `Collection` values.

## Run

- Import: `SampleOrderImport.bas`
- Macro: `RunOrderImportSample`

The sample validates one successful and one failing input and prints `ErrorText` to the Immediate Window.

## APIs used

- `ObjectSchema`, `Field`, `Strict`
- Nested `ObjectSchema` and `ArrayOf(Collection)`
- `Text().Pattern()`, `Text().Email()`, `DateTime()`
- `EnumOf`
- `Number().WholeNumber().Min().Max()`
- `OptionalField`

The failing input demonstrates paths and error codes such as:

```text
$.orderId           invalid_pattern
$.email             invalid_email
$.items[0].sku      invalid_pattern
$.items[0].quantity too_small
$.debug             unknown_field
```

Removing `Strict` also lets you observe the migration behavior when unknown fields are passed through.
