# ADR-0007: Release assets, version authority, and CI boundary

## Status

`accepted`

## Background

VBA-Schema limits the production source imported into the VBE to three files. README, LICENSE, CHANGELOG, sample workbooks, and verification information are useful to users and release maintainers, but mixing them into the VBE import payload would make the installation contract unclear.

Compile and behavioral tests that require Excel/VBE cannot always be reproduced on GitHub-hosted runners. CI without Excel must not report a false compile success, and macOS and Windows 32-bit must be represented accurately as unverified.

## Decision

- Treat Windows 64-bit Excel VBE compilation, behavioral tests, and release smoke before a stable release as a maintainer gate separate from the Excel-free GitHub-hosted `source-check`; fix the procedure in `docs/release-checklist.md). Excel-free status does not replace this evidence.
- Generate both the release ZIP and a SHA-256 checksum asset containing the same ZIP filename, attach both to the same GitHub Release tag, and clobber both on retries.
- Fix the VBE import payload to exactly `Schema.bas`, `VSchema.cls`, and `VValidationResult.cls`.
- Provide README, LICENSE, CHANGELOG, sample workbooks, ZIP files, and other supporting assets separately from the three-file payload. Update the v1 contract and ADR-0001 before changing this boundary.
- Treat the Git tag and tracked `CHANGELOG.md` as the version authority. Do not add a version constant or separate VERSION file to production VBA source.
- Verify staged VBA source as UTF-8 without a BOM and with LF line endings. Use an explicit allowlist for release staging and pass a fresh-workbook smoke test that compiles without additional references.
- Make `tools/release-stage.ps1` fail before deletion when Destination overlaps the repository root, the production source tree (`src`), xlflow state (`.git`, `.xlflow`), or the verification workbook (`build`). Even when arbitrary repository paths are accepted, do not permit overlap with protected ancestors or descendants.
- Run Excel-free `lint`, `analyze`, format checks, and test discovery in GitHub-hosted CI. Treat Windows Excel/VBE compilation and behavioral tests as separate local or self-hosted checks, and include `vbe` or `excel` in their status names to make their scope explicit.
- Never report static CI without Excel as VBE compilation passed. Add macOS/Windows 32-bit matrix jobs only after obtaining device verification.
- Accept only stable `vMAJOR.MINOR.PATCH` tag pushes in the GitHub Release workflow, run the reusable `source-check` first, then create a `VBA-Release-vX.Y.Z.zip` containing only the three production files under `VBA-Release/`. Retries against an existing Release update same-named assets with clobber behavior.

## Consequences

- The import procedure remains three files, and supporting-document changes cannot alter the number of VBA components.
- Git tags are the single runtime-independent version source, so v1 does not provide an API for retrieving a version from the VBA project.
- CI detects Excel-free failures early, but VBE compilation still depends on a separate environment and cannot be proven by static CI alone.
- A checksum asset lets users verify that a downloaded release ZIP matches the published artifact. The checksum is generated per tag instead of being manually hard-coded for every source change.
- UTF-8/LF checks detect release-payload encoding differences early. VBE-host differences remain verified on Windows 64-bit and unverified on macOS/32-bit.
- Prerelease tags are explicitly out of scope; adding them later requires a separate decision covering tag validation, CHANGELOG, release notes, and prerelease display.

## Rationale

- Evidence: `docs/specs/v1-contract.md` and `tools/release-stage.ps1`/`tools/release-verify.ps1` define the three-file allowlist.
- Evidence: `tools/release-smoke.ps1` verifies fresh-workbook import, VBE compilation without additional references, scalar smoke behavior, and the non-document component count.
- Evidence: `docs/release-checklist.md` defines Windows 64-bit maintainer evidence that Excel-free CI cannot replace; `tools/package-release.ps1` and `.github/workflows/release.yml` verify the ZIP/checksum pairing.
- Evidence: DG-008/DG-009 and the M7 release/CI tasks in `tasks/todo.md` separate Excel-free checks from VBE checks.
- Related docs: `README.md`, `CHANGELOG.md`, `docs/design.md`.

## Supersedes

- None

## Superseded by

- None
