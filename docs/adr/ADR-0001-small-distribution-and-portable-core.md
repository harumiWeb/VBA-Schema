# ADR-0001: Three-file distribution and a portable core

## Status

`accepted`

## Background

VBA-Schema is intended to be a runtime validation library that can be brought into an existing VBA project easily. The number of files and reference settings create adoption barriers, while VBA lacks conventional package boundaries and internal visibility. Excessive class decomposition would make distribution and updates harder.

Development and verification are performed on Windows 64-bit Office. macOS Office and Windows 32-bit Office are not currently available for verification, so the implementation may remain compatibility-conscious without claiming operational support.

## Decision

- Limit the production components in the distribution to `Schema.bas`, `VSchema.cls`, and `VValidationResult.cls`.
- Treat `dist/VBA-Schema/` as an import payload generated from those three files. It is a staging artifact separated from source and excluded from commits.
- Do not include tests, the xlflow execution infrastructure, samples, or generation helpers in the distribution.
- Do not require external reference settings; use late binding for runtime components where necessary.
- Fix `Scripting.Dictionary` as the required concrete type for the v1 public Issue representation. Do not introduce a portable fallback in v1; report an unavailable component as an environment failure.
- Do not add dependencies on the Windows API, Win32-specific declarations, or the Excel Object Model to the core.
- Treat Windows 64-bit Office as the verified support target.
- Describe macOS Office and Windows 32-bit Office as compatibility-conscious but unverified; do not describe them as supported.
- Do not weaken the verified environment's error contract or type safety to accommodate unverified environments.
- Describe `Scripting.Dictionary` and `VBScript.RegExp` as runtime facilities that need no reference setting, not as having no dependencies.

## Consequences

- Installation, removal, and updates consist of operations on three files.
- Responsibilities are concentrated in `VSchema.cls`, so the class needs clear private procedures and internal regions.
- Dedicated Issue classes and schema subtypes cannot be added, which limits type safety in internal representations.
- Reports from macOS or 32-bit Office can inform compatibility improvements, but no fix can be declared complete or supported until a reproduction environment is available.
- Missing late-bound runtime components are reported as environment errors instead of being hidden as validation failures.

## Rationale

- Tests: `docs/specs/v1-contract.md` defines the distribution smoke test and compatibility gate. `tools/release-smoke.ps1` verifies three-file import into a fresh workbook, VBE compilation, scalar smoke behavior, and three non-document components.
- Code: production source is located in `src/modules/Schema.bas`, `src/classes/VSchema.cls`, and `src/classes/VValidationResult.cls`; `dist/VBA-Schema/` is a non-committed payload generated from the allowlist.
- Related specs: `docs/specs/v1-contract.md`

## Supersedes

- None

## Superseded by

- None
