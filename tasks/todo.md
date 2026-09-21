# VBA-Schema v1 Development Roadmap

Last updated: 2026-09-21

This document is the execution roadmap for completing VBA-Schema v1 with the same design contract and verification criteria, even when work is handed to another contributor.

## 1. How to use this document

### 1.1 Read these first

Review these sources in order before starting work:

1. `AGENTS.md`
2. `tasks/todo.md` (this document)
3. `tasks/lessons.md`
4. `docs/specs/v1-contract.md`
5. Relevant ADRs in `docs/adr/`
6. `docs/design.md`
7. Existing source and focused tests
8. `xlflow.toml` and `Taskfile.yml`

The authority order is:

1. User request
2. `AGENTS.md` and repository safety rules
3. v1 specifications
4. Accepted ADRs
5. Current source and tests
6. This roadmap
7. Explanatory notes in this roadmap

When a contradiction is found, stop implementation and update the specification or ADR instead of silently absorbing the contradiction in code.

### 1.2 Checklist state

- `[ ]`: not started or no completion evidence
- `[x]`: completion criteria met and evidence recorded in the Progress Log
- `BLOCKED:`: progress cannot continue because of an external environment or unresolved Decision Gate

For each completed phase, record:

- Date
- Commit SHA or an explicit uncommitted status
- Changed files
- Commands run
- Test results
- Absolute paths of temporary workspaces used
- Unverified items
- Next starting point

Use task identifiers in the form `<Milestone>/<subheading>/<short-label>`, for example `M3/Tests/binary-compare-mode`. Do not use line numbers or checkbox numbers because edits change them.

### 1.3 Change rules

- Production distribution consists only of `Schema.bas`, `VSchema.cls`, and `VValidationResult.cls`.
- Tests, fixtures, release helpers, and xlflow helpers do not count against the three-file distribution boundary.
- Add a focused regression test in the same work unit as every production-code change.
- Update the specification before changing a Public API, error code, path format, or type semantic; create or supersede an ADR when necessary.
- Use the `adr-manager` skill when editing ADRs.
- Use the `xlflow` source-to-workbook proof loop when validating VBA source.
- Record xlflow defects under `docs/xlflow-issues/` and do not distort permanent VBA-Schema semantics to hide a tool workaround.
- Never describe macOS Office or Windows 32-bit Office as supported or verified without device evidence.

## 2. Current checkpoint

### 2.1 Repository state

- Production implementation covers scalar schemas, Object and nested Object validation, Array/Collection validation, Literal, Enum, Pattern, Email, and Union.
- Production files are `Schema.bas`, `VSchema.cls`, and `VValidationResult.cls`.
- Test discovery finds 61 tests in 24 source files; the current full Excel run has 60 passes and one intentional TODO.
- The tracked development workbook is `build/Book.xlsm`.
- The public v1 contract, ADRs, design specification, README, samples, release checklist, and CHANGELOG are maintained in English.
- The current branch is `harumiWeb/improve-vba-schema-todo`; local changes are intentionally uncommitted.
- The remote `origin/main` README revision `1ff6cfedb28702eb5f798d1b2794492cdc8d7153` has been incorporated into the working tree.

### 2.2 Confirmed evidence

- Windows 64-bit Office VBE compilation succeeds for the Public API compile fixture.
- The fixture covers factory declarations, fluent chains, typed arrays, Object and scalar Result assignment, and internal initialization boundaries.
- `PublicApiCompile.CompilePublicApi` runs successfully in a managed Windows x64 session.
- `rtk task verify` passes static lint, analysis, format, production hygiene, sample verification, test discovery, release safety, and benchmark-environment checks.
- `rtk task release-smoke` passes three-file import, VBE compilation without additional references, scalar smoke, and non-document component checks.
- `rtk task release-package TAG=v0.0.0` creates the ZIP and SHA-256 checksum asset.
- `build/Book.xlsm` has been refreshed from source, compiled, tested, saved, and added to the Git index.
- Final xlflow status after save reports inactive, clean, `save_required=false`, `recovery_required=false`, and `src_newer_than_workbook=false`.

### 2.3 Explicitly unverified

- A real GitHub-hosted Excel/VBE compile or behavioral-test job.
- A real stable tag push and GitHub Release upload.
- macOS Office and Windows 32-bit Office.
- Physical reproduction of unavailable Dictionary or RegExp runtimes.
- Automatic verification that no external reference was added after a clean remote checkout.
- A fresh clone performed from a remote commit containing the current uncommitted changes.

### 2.4 v1.0 release-candidate follow-ups

- [x] Separate Number type acceptance from numeric comparison so a finite Double outside the Decimal range is not rejected by type checking alone.
- [x] Add Number regression coverage for `CDbl(1E+100)`, numeric boundaries, and `WholeNumber`.
- [x] Document that Excel-free source-check does not prove VBE compilation or behavioral tests; add the maintainer release checklist.
- [x] Pin GitHub Actions third-party actions to full commit SHAs and generate and verify a release ZIP checksum asset.
- [x] Put the user value, three-file installation, and a short example first in the public README.
- [x] Synchronize the specification, ADRs, design document, CHANGELOG, and Progress Log with the implementation and release boundary.
- [x] Translate all Markdown documentation other than README.md into English.

### 2.5 Tracked development workbook

- [x] Track `build/Book.xlsm` as a binary development fixture and synchronize `.gitignore`, `.gitattributes`, repository trees, and provisioning documentation.
- [x] Push the current source into the workbook, run VBE compilation and behavioral tests, save it, and confirm a clean xlflow status.

## 3. Milestone dependency map

```text
M0 contract and gates
  -> M1 harness, compile fixture, and release staging
    -> M2 scalar core and Result
      -> M3 Object validation
        -> M4 Array and Collection validation
          -> M5 Literal, Enum, Pattern, and Email
            -> M6 Union
              -> M7 hardening, documentation, CI, and release
                -> M8 independent review and completion
```

Do not start a later milestone until the previous milestone's exit gate is satisfied.

## 4. Decision Gates

Unresolved decisions must not become hidden assumptions in production code. Decide them before the relevant phase.

### DG-001 — Portable Issue representation

- Status: accepted.
- Use late-bound `Scripting.Dictionary` for Issue and snapshot records.
- Do not add a portable fallback in v1.
- A missing component required by the library's own Issue generation is an environment failure.

### DG-002 — Numeric comparison

- Status: accepted.
- Accept Byte, Integer, Long, Single, Double, and Currency as Number.
- Keep Number acceptance independent from Decimal conversion.
- Use safe subtype-aware comparison only when evaluating numeric constraints or `WholeNumber`.
- Reject unproven precision loss, overflow, NaN, and Infinity.

### DG-003 — ErrorText and received descriptors

- Status: accepted.
- Use deterministic paths, codes, messages, expected values, and safe received descriptors.
- Snapshot Issues and keep `ErrorText` deterministic.
- Do not expose raw Object references, input lifetimes, secrets, or host-locale formatting.

### DG-004 — Schema cycle detection

- Status: accepted.
- Run DFS preflight immediately before validation.
- Track the active path in a Collection and compare schema identity with VBA `Is`.
- Reject direct and indirect cycles with `vbObjectError + 2103`; allow non-cyclic shared children.

### DG-005 — Dictionary detection and key matching

- Status: accepted.
- Start with `TypeName(value) = "Dictionary"` and limited checks for `Count`, `Exists`, and `Keys`.
- Use binary, case-sensitive field matching independent of input `CompareMode`.
- Report non-String keys as `invalid_key`; report strict unknown String keys as `unknown_field`.

### DG-006 — Native array inspection

- Status: accepted.
- Accept one-dimensional native arrays and Collections.
- Treat an uninitialized dynamic array as empty and multidimensional arrays as `invalid_array_rank`.
- Normalize paths to zero-based logical indexes and protect the Err state during bounds probing.

### DG-007 — Pattern and Email

- Status: accepted.
- Late-bind `VBScript.RegExp` with fixed flags and full-input matching.
- Compile lazily; invalid expressions are programmer misuse and missing RegExp is an environment failure.
- Email is a practical ASCII validator, not full RFC compliance.

### DG-008 — Release artifact

- Status: accepted.
- The release ZIP contains exactly `Schema.bas`, `VSchema.cls`, and `VValidationResult.cls`.
- Release staging uses an explicit allowlist and verifies encoding, headers, and component count.
- Provide the ZIP and its SHA-256 checksum asset under the same stable tag.

### DG-009 — CI and compile ownership

- Status: accepted.
- GitHub-hosted source-check runs Excel-free lint, analysis, format, test discovery, safety, release, and benchmark-environment checks.
- Windows x64 VBE compilation, behavioral tests, and release smoke are a maintainer gate documented in `docs/release-checklist.md`.
- Excel-free CI must never be reported as VBE compilation passed.

### DG-010 — InternalInitialize exposure

- Status: accepted.
- Keep `InternalInitialize` Public only because of VBA class initialization constraints.
- Treat it and the other `Internal*` hooks as unsupported internal-only APIs.
- Reject direct normal-kind calls and reinitialization with the documented programmer-misuse errors.

### DG-011 — Constraint composition

- Status: accepted.
- Builders mutate and return the same instance.
- Detect duplicate or contradictory constraints at builder time.
- Keep modifier applicability and fixed constraint order in the v1 contract.

### DG-012 — Schema input ownership

- Status: accepted.
- Snapshot Enum candidates and Union branch Collections at construction.
- Share child `VSchema` references rather than cloning them.
- Make post-construction child-builder mutation visible through the shared reference.

## 5. M0 — Design baseline and decision gates

### Tasks

- [x] Define the three-file boundary.
- [x] Define the v1 Public API and VBA-safe names.
- [x] Define Issue, path, type, error, compatibility, and distribution contracts.
- [x] Record DG-001 through DG-012 in the specification and ADRs.
- [x] Define the source-to-workbook proof loop and Excel-free CI boundary.

### Verification

- [x] Read and cross-check the specification, ADRs, design, source tree, and tests.
- [x] Confirm that the release payload is separate from development and test modules.
- [x] Record contradictions and tool limitations under `docs/xlflow-issues/`.

### Exit gate

- [x] The v1 contract and design baseline are explicit.
- [x] No unresolved gate blocks the next milestone.
- [x] Public API names have a compile fixture.

## 6. M1 — Development, test, and release harness

### Tasks

- [x] Add Taskfile wrappers for source verification, workbook provisioning, session tests, release staging, release verification, and release smoke.
- [x] Keep `build/Book.xlsm` as a tracked development fixture; provision only when it is missing.
- [x] Refuse to overwrite a tracked or user-owned workbook.
- [x] Add the three-file release allowlist and ZIP packaging.
- [x] Add source encoding, line-ending, class-header, and component-count checks.
- [x] Add the Public API compile fixture.

### Exit gate

- [x] A clean development workbook can be used immediately.
- [x] A missing workbook can be rebuilt from an isolated xlflow scaffold.
- [x] Existing workbooks are never overwritten by provisioning.
- [x] Release staging contains exactly three production components.
- [x] The compile fixture covers the complete Public API surface.

## 7. M2 — Scalar core and Result contract

### Tests first

- [x] AnyValue, Text, Number, Bool, and DateTime.
- [x] OptionalField and Nullable.
- [x] Result success/failure properties and Issue snapshots.
- [x] Programmer misuse and runtime/environment error boundaries.
- [x] Number subtype, boundary, WholeNumber, NaN, Infinity, and Decimal-range regressions.

### Production implementation

- [x] Implement Schema factories and VSchema scalar validation.
- [x] Implement VValidationResult snapshots and deterministic ErrorText.
- [x] Keep Object return values and scalar return values consistent with VBA assignment rules.
- [x] Avoid implicit coercion.

### Exit gate

- [x] Focused scalar tests pass in Windows x64 Excel.
- [x] Public API compilation passes.
- [x] Result snapshots are isolated from caller mutation.
- [x] Three-file release smoke passes.

## 8. M3 — Object validation

### Tests first

- [x] Required, OptionalField, Nullable, nested paths, Strict, field order, and binary matching.
- [x] Case differences and input Dictionary CompareMode.
- [x] Non-String keys, Nothing, input non-mutation, and direct/indirect cycles.
- [x] Missing runtime component behavior.

### Implementation

- [x] Implement ObjectSchema, Field, Strict, Dictionary capability checks, and canonical paths.
- [x] Implement deterministic unknown-field and invalid-key Issues.
- [x] Add DFS schema-graph preflight.

### Exit gate

- [x] Object focused tests and release smoke pass.
- [x] Input Dictionaries are not modified.
- [x] Cycle errors occur before validation begins.

## 9. M4 — Array and Collection validation

### Tests first

- [x] Native arrays with zero, one, negative, and non-zero bounds.
- [x] Collections, nested sequences, logical paths, length constraints, and special elements.
- [x] Uninitialized and multidimensional arrays.
- [x] Typed Object arrays and Nothing elements.

### Implementation

- [x] Implement ArrayOf and item-schema recursion.
- [x] Protect LBound/UBound probes and normalize logical indexes.
- [x] Apply array length constraints before item validation.

### Exit gate

- [x] Array focused tests and release smoke pass.
- [x] Inputs remain unchanged.
- [x] Array and Object semantics remain distinct.

## 10. M5 — Literal, Enum, Pattern, and Email

### Tasks

- [x] Implement scalar category comparison and Literal.
- [x] Snapshot Enum candidates and reject invalid or duplicate definitions.
- [x] Implement lazy RegExp creation and fixed Pattern flags.
- [x] Implement practical ASCII Email validation.
- [x] Add programmer misuse and unavailable-runtime tests.

### Exit gate

- [x] Category, snapshot, duplicate, full-match, Pattern-flag, and Email boundary tests pass.
- [x] RegExp failures remain separate from validation failures.
- [x] Scope and RFC limitations are documented in English.

## 11. M6 — Union

### Tasks

- [x] Snapshot non-empty branch Collections.
- [x] Share child schema references without cloning.
- [x] Preserve nested Union structure.
- [x] Return one `invalid_union` Issue on all-branch validation failure.
- [x] Propagate programmer and environment errors from branches.

### Exit gate

- [x] Union focused tests and full behavioral tests pass.
- [x] Collection mutation after construction does not change branch membership.
- [x] Nested cycle detection remains active.

## 12. M7 — Hardening, documentation, CI, and release

### Tasks

- [x] Run full lint, analysis, format, hygiene, sample, release-safety, and benchmark-environment checks.
- [x] Add runtime benchmark fixtures, counters, thresholds, and a Windows x64 baseline.
- [x] Separate Excel-free source-check from local VBE/behavioral evidence.
- [x] Pin GitHub Actions and validate release checksum assets.
- [x] Add the English public README from `origin/main`.
- [x] Translate all remaining Markdown documentation into English.
- [x] Record compatibility limitations and the maintainer release checklist.

### Exit gate

- [x] No blocking findings remain in the independent review.
- [x] Release staging and release smoke pass.
- [x] Windows x64 benchmark gates pass under the documented contract.
- [x] Public documentation, samples, specification, ADRs, and development notes are English.

## 13. M8 — Independent final review and completion

### Review

- [x] Run an independent review in a separate worktree.
- [x] Validate the implementation, contract, tests, release boundary, and CI claims.
- [x] Resolve valid blocking and high-severity findings.
- [x] Keep unsupported observations separate from confirmed findings.

### Final verification commands

```powershell
rtk task verify
rtk task samples-verify
rtk xlflow doctor --json
rtk xlflow session start --json
rtk xlflow push --session --no-save --json
rtk xlflow test --session --no-save --json
rtk xlflow run PublicApiCompile.CompilePublicApi --diagnostic --headless --session --no-save --json
rtk xlflow save --session --json
rtk xlflow status --json
rtk xlflow session stop --json
rtk task release-smoke
rtk task release-stage
rtk task release-verify
rtk task release-package TAG=v1.0.0
```

Do not claim that GitHub-hosted Excel-free checks prove VBE compilation or behavioral tests.

### User-facing samples

- [x] Maintain independent order, settings, and API-response samples.
- [x] Keep sample source outside the three-file import payload.
- [x] Verify sample format, lint, analyze, import, compile, and macro execution.

### v1 completion gate

- [x] Production distribution contains exactly three VBA components.
- [x] No additional reference setup is required.
- [x] Windows 64-bit VBE compilation passes.
- [x] macOS and Windows 32-bit are explicitly unverified and unsupported.
- [x] All v1 schema types and modifiers are implemented and tested.
- [x] Validation errors contain deterministic paths and structured codes.
- [x] xlflow test, lint, and analyze pass.
- [x] The README contains installation, examples, and API documentation.
- [x] The release artifact contains only the three production files.
- [ ] A stable GitHub tag and Release have been published and read back from GitHub.

## 14. Deferred beyond v1

- Transforms and arbitrary callback refinements.
- Automatic coercion.
- Class reflection.
- JSON parsing, HTTP requests, and worksheet conversion.
- OpenAPI generation and schema code generation.
- A large hierarchy of schema, Issue, constraint, or adapter classes.
- Portable replacements for Dictionary and RegExp until target-host evidence exists.
- Prerelease tag workflow.
- A self-hosted Excel runner in GitHub Actions.
- A `Parse` API that raises on validation failure.
- Positive and Negative modifiers.

## 15. Handoff protocol

Before handing off work:

1. Update this document and the Progress Log.
2. Run `rtk git status --short`.
3. Run `rtk xlflow status --json`.
4. Record commands, results, absolute temporary-workspace paths, and unverified items.
5. State the next concrete command.
6. Do not delete or discard user-owned workbooks.
7. Do not use destructive Git commands without explicit authorization.

A handoff must not rely on the Progress Log alone; the next contributor must re-check Git status, specifications, and xlflow status.

## 16. Progress Log

New entries are added at the top.

### 2026-09-21 — English documentation and main README integration

- Status: Integrated the English README from remote `origin/main` and translated all other Markdown documentation, including public samples, CHANGELOG, specifications, design, ADRs, release checklist, xlflow issue notes, lessons, and roadmap.
- Changed: `README.md`, `CHANGELOG.md`, `sample/**/README.md`, `docs/**/*.md`, `tasks/lessons.md`, and `tasks/todo.md`.
- Verification target: No Japanese characters remain in Markdown files other than the README source content, which is also English.
- Unverified: The documentation-only changes do not require Excel execution. GitHub remote checks, stable Release publication, macOS, and Windows 32-bit remain unverified.

### 2026-09-21 — v1.0 release-candidate implementation

- Status: Number Decimal-range acceptance, Excel-free release boundary, SHA-pinned Actions, checksum assets, and tracked development workbook are implemented.
- Number: Number type acceptance is independent of Decimal conversion. Finite Doubles such as `CDbl(1E+100)` are accepted; `Min`, `Max`, and `WholeNumber` use safe subtype-aware comparison.
- Workbook: `build/Book.xlsm` was refreshed from source, compiled, behaviorally tested, saved, and added to the Git index.
- Verification: `rtk task verify`, `rtk task release-smoke`, `rtk task release-package TAG=v0.0.0`, `rtk actionlint`, and `rtk git diff --check` pass. The full Excel run reports 60 passes and one intentional TODO out of 61 discovered tests.
- Release ZIP SHA-256: `80574fa28cff0f30539549397891a90562bf89559a9154f037dc46e988b78c0b`.
- Workbook SHA-256: `507a1ab5dae93ddff6c0a8f9dd5bc262e28a4f64604ed1bd67194d5c80c08449`.
- Unverified: GitHub-hosted Excel/VBE execution, stable tag and Release publication, macOS, Windows 32-bit, and unavailable Dictionary/RegExp runtimes.

### 2026-09-21 — M8 independent review

- Status: Two independent review passes completed in separate worktrees with no blocking finding.
- Verification: Static checks, source lint, analysis, production hygiene, test discovery, release staging/verification, release smoke, benchmark, full behavioral tests, and Public API compilation passed under their documented environments.
- Unverified: Remote GitHub workflow execution and unsupported Office environments.

### 2026-09-21 — M7 hardening and release boundary

- Status: Runtime benchmark contract, release/CI boundary, sample projects, release checksum handling, and compatibility documentation were added.
- Verification: Windows x64 runtime benchmark, tracked baseline comparison, release smoke, release staging, release verification, and Excel-free source checks passed.
- Unverified: GitHub-hosted Excel/VBE execution, macOS, Windows 32-bit, unavailable Dictionary/RegExp runtimes.

### 2026-09-21 — M6 Union

- Status: Union production implementation, branch ownership, nested validation, single `invalid_union` boundary, focused/full proof loop, Public API compilation, and release smoke completed.
- Verification: Union branch order, later-branch success, all-failure behavior, snapshots, shared references, nested Union, special values, and branch errors are covered.

### 2026-09-21 — M5 Literal, Enum, Pattern, and Email

- Status: Scalar category comparison, Enum snapshots, Pattern flags/full match, Email boundaries, lazy RegExp creation, and programmer-misuse errors completed.
- Verification: Focused tests, full tests, and release smoke passed; unsupported RegExp runtime environments remain unverified.

### 2026-09-21 — M4 Array and Collection

- Status: Native arrays, Collections, logical indexes, length constraints, nested validation, uninitialized arrays, multidimensional rejection, and special elements completed.
- Verification: Focused/full tests and release smoke passed; alternate-host behavior remains unverified.

### 2026-09-21 — M3 Object validation

- Status: Object fields, Strict, binary matching, canonical paths, unknown/non-String key Issues, nested validation, and schema cycle preflight completed.
- Verification: Object tests, error tests, release smoke, and focused proof loop passed.

### 2026-09-21 — M1 harness and M2 scalar core

- Status: Three-file staging, compile fixture, release smoke, scalar validation, Result snapshots, error boundaries, and focused/full proof loop completed.
- Verification: Windows x64 VBE compile and managed-session tests passed; the intentional sample TODO remains.
- Unverified: Remote CI Excel/VBE execution and unsupported Office hosts.

### 2026-09-21 — M0 baseline

- Status: Initial design baseline, decision gates, three-file payload ownership, and implementation roadmap established.
- Verification: Specification, ADRs, design, source tree, compile fixture, and release staging policy were cross-checked.
