# ADR-0003: Dictionary Object detection and schema-graph cycle boundary

## Status

`accepted`

## Background

The v1 Object schema is limited to `Scripting.Dictionary` as the sole implementation for treating external input as key-value data. VBA has no general interface reflection, and a Dictionary's `CompareMode` may vary between inputs. Mutable builders can also create direct or indirect self-reference after a schema is shared, so a boundary is needed to prevent infinite recursion during validation.

## Decision

- Begin Object detection with `TypeName(value) = "Dictionary"` and perform limited capability checks for `Count`, `Exists`, and `Keys`.
- Treat a TypeName mismatch or failed capability check as an `invalid_type` validation Issue. If a runtime component required by the library itself, such as the Dictionary used to create an Issue, cannot be created, propagate `vbObjectError + 2200`.
- Match fields case-sensitively with key enumeration and `StrComp(..., vbBinaryCompare)`, independently of the input Dictionary's `CompareMode`. Do not modify the input Dictionary during validation.
- A non-String input key produces an `invalid_key` Issue regardless of `.Strict()`. The path is the Object node path, `received` is a safe descriptor, and multiple keys are ordered by binary ordinal descriptor order.
- Pass through String unknown fields by default; add an `unknown_field` Issue only for `.Strict()`. Use canonical paths and the `DescribeValue` grammar for their path and `received`, ordered by binary ordinal.
- Report a missing required field as `received = "Missing"`. This is a distinct public state from `Empty` and `Null`.
- Preflight the schema graph with DFS at the start of `SafeParse`. Store the active path in a `Collection` and compare schema identity with VBA `Is`. Raise `vbObjectError + 2103` for direct or indirect cycles before validation begins; allow non-cyclic reuse of shared children.
- Keep the `Internal*` hooks Public because VBA cannot access Private members through another instance of the same class. These hooks are needed for composition, nested validation, and cycle traversal, and are unsupported internal-only APIs outside the stable user-facing compatibility guarantee, like `InternalInitialize`.

## Consequences

- The same algorithms can run on macOS and Windows 32-bit Office without Windows API or pointer-size dependencies, but no support is guaranteed without device verification.
- Binary field matching prevents case-folding mistakes caused by Dictionary `CompareMode` differences. Enumerating keys for each lookup leaves room for later performance improvement on large Objects.
- Reporting non-String keys as explicit validation failures safely exposes an external contract violation, but users who previously relied on arbitrary Variant keys need migration guidance.
- `Internal*` hooks are an implementation boundary around VBA visibility constraints and must not be treated as future public APIs.

## Rationale

- Tests: `TestObject.bas` covers required/optional fields, nested paths, binary matching, strict ordering, non-String keys, and input non-mutation. `TestErrors.bas` covers builder misuse and direct/indirect cycle Err numbers.
- Code: `GetDictionary`, `TryFindDictionaryKey`, `ValidateObject`, and `EnsureAcyclicGraph` in `VSchema.cls` implement this decision.
- Related specs: `docs/specs/v1-contract.md`, `docs/design.md`

## Supersedes

- None

## Superseded by

- None
