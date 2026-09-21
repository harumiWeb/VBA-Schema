# ADR-0002: VBA-safe API names and the error boundary

## Status

`accepted`

## Background

The original Zod-like API included names such as `Any`, `String`, `Boolean`, `Date`, `Object`, `Optional`, and `Integer`, which collide with VBA reserved words or type names. The Windows 64-bit Office VBE compile oracle showed that at least declarations using `Any`, `Optional`, and `Integer` as procedure names fail to compile.

If `SafeParse` converted every runtime error into a validation failure, users could not distinguish invalid input, schema construction mistakes, missing runtime components, and library bugs.

## Decision

- Use `AnyValue`, `Text`, `Number`, `Bool`, `DateTime`, `ObjectSchema`, `ArrayOf`, `Literal`, `EnumOf`, and `UnionOf` as v1 factory names.
- Use `OptionalField`, `Nullable`, `Min`, `Max`, `Length`, `WholeNumber`, `Pattern`, `Email`, `Field`, and `Strict` as v1 modifier names.
- Restrict Union inputs containing Objects to a `Collection` of `VSchema` instances rather than a Variant array.
- Adopt a public API addition or change only after a fixture containing every signature and fluent chain passes VBE compilation.
- Convert only validation failures into a Result through `SafeParse`.
- Raise schema-construction programmer misuse and runtime/environment failures in distinct, recognizable Err number ranges.
- Keep `InternalInitialize` Public because of VBA class initialization constraints, but classify it as an unsupported internal-only API. Only Schema factories may pass the internal encoded kind; ordinary direct calls, invalid kinds, and reinitialization are programmer misuse.
- Keep the `Internal*` hooks needed to compose, validate, and cycle-check child schemas Public because of VBA Private-member visibility constraints. Classify them like `InternalInitialize` as unsupported internal-only APIs outside the stable user-facing compatibility guarantee. ADR-0003 defines the detailed Object boundary.
- Separate Number type acceptance from comparison. Accept finite Byte, Integer, Long, Single, Double, and Currency values as Number even when they cannot be converted to Decimal. Use subtype-aware lossless widening, Currency comparison, and necessary round-trip checks only for `Min`/`Max` and `WholeNumber`; never silently round precision loss, overflow, NaN, or Infinity. Generate `ErrorText` and `received` with fixed grammar and keep constraint order fixed.
- Do not perform implicit type conversion. Type comparisons, including Literal and Enum comparisons, follow the category rules in `docs/specs/v1-contract.md`.

## Consequences

- The names are not as short as their Zod equivalents, but the fluent API remains compilable and stable in VBA.
- `SafeParse` still raises Err for environment failures and library bugs, so callers must add operational error handling where appropriate.
- Public API names are a breaking-change surface, making the compile fixture mandatory before implementation.
- Union definitions require preparing a `Collection`, but this removes ambiguous Object-array semantics from the public contract.

## Rationale

- Tests: `PublicApiCompile.bas` compiles every Public signature and fluent chain, while M2 focused tests verify scalar Results, Issue snapshots, and the error boundary. `TestNumber.bas` fixes the regression that finite Doubles outside the Decimal range are accepted as Number. Windows 64-bit Office VBE compilation and focused/full tests were run on 2026-09-21.
- Code: scalar core is implemented in `Schema.bas`, `VSchema.cls`, and `VValidationResult.cls`. Object, Array, Literal, Enum, Union, Pattern, and Email were implemented in later milestones.
- Related specs: `docs/specs/v1-contract.md`

## Supersedes

- None

## Superseded by

- None
