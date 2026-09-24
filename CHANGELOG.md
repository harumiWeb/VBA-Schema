# Changelog

This file tracks releases made from Git tags. No version constant is embedded in the source payload.

## Unreleased

- Froze the v1 validation-only contract: `docs/specs/v1-contract.md` now records explicit v1 non-goals (coercion, `.Default()`, unknown-field stripping, transformations, `Partial`/`Pick`/`Omit`, `Clone`/`Freeze`, Int64/Decimal parsing, timezone conversion, and code generation) and a locale policy stating that strict schemas never parse localized text.
- Added Null/Empty/Error Variant regression coverage for constrained Text, Number, Boolean, and DateTime schemas, Object field states, and array/Collection elements.
- Finished the public README for v1.0: documented schema mutability and shared child references, `Value` pass-through semantics, locale-independent validation, sample links, and verified-environment wording.

## v0.2.0 — 2026-09-21

- Fixed Number validation so finite Doubles outside the Decimal range are not rejected by type checking alone; numeric constraint comparison and `WholeNumber` detection are now subtype-aware.
- Added the Windows 64-bit Excel maintainer checklist for stable releases and a SHA-256 checksum asset for the release ZIP.
- Added Union schemas with branch Collection snapshots, nested Unions, and a single `invalid_union` issue.
- Added scalar/category boundaries and runtime error boundaries for Literal, Enum, Pattern, and Email.
- Added staging, verification, and release smoke coverage for the three-file import payload.
- Indexed large-input Object validation lookups and added boundary helpers and a production compatibility audit.
- Added Windows 64-bit-only runtime benchmarks, a tracked baseline, and Excel-free release staging and verification.
- Added the standard MIT license text to `LICENSE`.
- Added user-facing order, settings, and API-response samples together with sample-specific format, lint, and analyze checks.
- Added a workflow that runs `source-check` first on `vMAJOR.MINOR.PATCH` tag pushes and attaches a three-module `VBA-Release-vX.Y.Z.zip` to the GitHub Release.

## v1.0.0

Not released.
