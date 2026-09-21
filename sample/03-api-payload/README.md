# API response-shaped payload validation

This sample validates a response after a JSON parser or another transport layer has converted it into VBA `Dictionary` and `Collection` values. It does not include JSON retrieval or parsing.

## Run

- Import: `SampleApiPayload.bas`
- Macro: `RunApiPayloadSample`

## APIs used

- Nested `ObjectSchema`
- `ArrayOf(Collection)`
- `UnionOf` (an email address or numeric ID)
- `Literal`, `Nullable`, `OptionalField`
- `Number().WholeNumber().Min()`, `Text().Length()`

The failing input shows exactly where the response is invalid:

```text
$.status         invalid_literal
$.contact        invalid_union
$.items[0].id    too_small
$.items[0].active invalid_type
$.traceId        unknown_field
```

Changing `contact` to an integer also demonstrates the second branch of the same Union succeeding.
