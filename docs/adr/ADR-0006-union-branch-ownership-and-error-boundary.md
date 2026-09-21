# ADR-0006: Union branch ownership and error boundary

## Status

`accepted`

## Background

`UnionOf` is a public API that treats multiple `VSchema` instances as one schema node. A VBA `Collection` is mutable; retaining the input Collection directly would allow external mutation after schema construction to change validation meaning. Child `VSchema` instances follow the same mutable-builder contract as existing Object and Array schemas, so introducing deep cloning would break shared references, cycle detection, and builder consistency.

If branch-specific failure details were exposed directly, the Issue contract would depend on branch count and internal structure. Failing to distinguish programmer misuse or runtime failure inside a branch from ordinary validation failure would hide schema-definition errors or missing environment components.

## Decision

- `UnionOf(ByVal Schemas As Collection)` snapshots the branch order and element set of a non-empty Collection during construction.
- Store snapshot elements as `VSchema` references without cloning child schemas. Changes made through a child builder after construction therefore affect Union validation.
- Treat `Nothing`, an empty Collection, non-`VSchema` elements, and uninitialized `VSchema` instances as programmer misuse and raise `vbObjectError + 2100`.
- Validate in branch order and succeed on the first successful branch. Do not flatten nested Unions; preserve the schema-graph structure for recursive validation.
- When every branch produces a validation failure, do not expose branch Issues. Add one `invalid_union` Issue (`expected=Union`) at the original failed path.
- Do not convert programmer misuse, runtime/environment failures, or internal invariant failures inside a branch into `invalid_union`; propagate them unchanged.

## Consequences

- Later `Add` or `Remove` operations on the Collection passed to Union do not affect the existing schema.
- Child fluent builders follow the existing shared-reference contract, so users must manage mutations made after sharing a child schema.
- The number and shape of validation Issues remain stable regardless of branch count; branch-specific diagnostics are not part of the v1 public result.
- Nested Union recursion is covered by schema-graph preflight, so cycles are detected as `vbObjectError + 2103`.
- If a branch lacks an environment-dependent component such as Dictionary or RegExp, Union does not hide the environment error.

## Rationale

- Tests: `src/modules/Tests/TestUnion.bas` verifies branch order, later-branch success, a single `invalid_union`, Collection snapshots, shared child references, nested Unions, special values, and branch programmer errors.
- Tests: `src/modules/Tests/TestErrors.bas` verifies `vbObjectError + 2100) for empty, Nothing, non-VSchema, and uninitialized branches.
- Code: `InternalSetUnion`, `ValidateUnion`, and `VisitSchemaGraph` in `src/classes/VSchema.cls` implement this boundary.
- Related specs: `docs/specs/v1-contract.md`, `docs/design.md`.

## Supersedes

- None

## Superseded by

- None
