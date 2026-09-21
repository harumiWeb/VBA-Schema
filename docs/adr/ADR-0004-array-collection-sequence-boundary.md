# ADR-0004: Array/Collection sequence boundary and safe array probing

## Status

`accepted`

## Background

Native VBA arrays can expose ambiguous states through Variant values: `LBound`, one versus multiple dimensions, uninitialized dynamic arrays, and typed Object arrays. A `Collection` returned by a JSON parser is treated as the same logical sequence, while a Dictionary is reserved for Object schemas and must not be implicitly converted into an array.

## Decision

- `ArrayOf(ItemSchema)` accepts one-dimensional native arrays and `Collection` objects, but does not treat `Scripting.Dictionary` or arbitrary classes as arrays.
- Treat an uninitialized dynamic array as a zero-element sequence. Return an `invalid_array_rank` Issue for multidimensional arrays instead of leaking a runtime error.
- Do not expose native array indexes or the one-based Collection index in public paths; normalize enumeration order to a zero-based logical index.
- Apply `Length`, `Min`, and `Max` to sequence length before recursively validating the item schema, and add only the first constraint Issue to the array node.
- Validate items depth-first in logical-index order without modifying the input array or Collection. Delegate Object, Nothing, and Error Variant elements to the ordinary child-schema contract.
- Probe array rank and bounds through limited helpers. Clear and restore the Err state used to detect an uninitialized array inside the helper. Do not use the Windows API, `ObjPtr`, or pointer-size-specific behavior.

## Consequences

- Arrays with non-zero lower bounds still produce stable Issue paths beginning at `$[0]`.
- An uninitialized array can be treated as an empty sequence without misclassifying `LBound`/`UBound` exceptions as validation failures.
- Large sequences perform late-bound access for each element, so performance targets are verified by the M7 benchmark.
- The structure uses only standard VBA facilities and is compatibility-conscious on macOS Office and Windows 32-bit Office, but remains unverified and unsupported there.

## Rationale

- Tests: `TestArray.bas` covers zero, one, and negative bounds; Collections; nested sequences; logical paths; length; uninitialized and multidimensional arrays; and typed Object/Nothing elements.
- Code: `GetNativeArrayState`, `TryGetPrimaryArrayBounds`, `HasSecondArrayDimension`, and `ValidateArray` in `VSchema.cls` implement this decision.
- Related specs: `docs/specs/v1-contract.md`, `docs/design.md`

## Supersedes

- None

## Superseded by

- None
