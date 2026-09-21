# User-facing samples

This directory contains minimal examples for adding VBA-Schema to an ordinary VBA project.

## Common setup and execution

1. Back up the target VBA project.
2. Import `Schema.bas`, `VSchema.cls`, and `VValidationResult.cls` into the VBE.
3. Import the `.bas` file from one of the directories below as a standard module.
4. Run the `Run...` macro documented in the corresponding README from the VBE Immediate Window.

The samples do not use the Excel Object Model, so results can be inspected from the Immediate Window without preparing a worksheet. Input fixtures are built in VBA as `Scripting.Dictionary` and `Collection` values representing converted Excel or JSON data.

## Examples

| Sample | Primary use case | Main APIs |
| --- | --- | --- |
| [01-order-import](01-order-import/README.md) | Validation before importing orders and line items | `ObjectSchema`, `ArrayOf`, `EnumOf`, `DateTime`, `Strict` |
| [02-settings-validation](02-settings-validation/README.md) | Validation of settings-sheet and form values | `EnumOf`, `Pattern`, `OptionalField`, `Nullable` |
| [03-api-payload](03-api-payload/README.md) | Validation of API-response-shaped data | nested objects, `UnionOf`, `Literal`, `ArrayOf` |

## Scope and compatibility

- Samples are not part of the three-file VBA-Schema import payload.
- `Scripting.Dictionary` and `VBScript.RegExp` are used through late binding. If they are unavailable, the result may be an environment error rather than an input validation failure.
- VBE compilation and runtime behavior are verified on Windows 64-bit Office. macOS Office and Windows 32-bit Office are compatibility-conscious but not verified or guaranteed.
- Fetching JSON, converting worksheet data into Dictionaries, and reading settings values are outside the samples' scope. The samples focus on validating already-converted VBA values.
