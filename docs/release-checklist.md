# v1.0 stable release checklist

This checklist complements the GitHub-hosted, Excel-free `source-check`. A successful `source-check` does not prove that VBE compilation or Excel behavioral tests succeeded. Before pushing a stable tag, a maintainer with Windows 64-bit Excel must collect the following evidence.

## Required evidence

- The working tree contains no unintended changes.
- `rtk task verify`, `rtk task samples-verify`, `rtk task release-stage`, and `rtk task release-verify` succeed.
- VBE compilation succeeds in a managed xlflow session on Windows 64-bit Excel after the source is pushed.
- `rtk xlflow test --session --no-save --json` runs in the same session and every test except the intentional TODO passes.
- `rtk task release-smoke` verifies three-file import into a fresh workbook, compilation without additional references, and the scalar smoke test.
- `rtk xlflow status --json` reports a clean session with `recovery_required` set to `false`.

## Suggested sequence

```powershell
rtk task verify
rtk task samples-verify
rtk xlflow doctor --json
rtk xlflow session start --json
rtk xlflow push --session --no-save --json
rtk xlflow test --session --no-save --json
rtk xlflow run PublicApiCompile.CompilePublicApi --diagnostic --headless --session --no-save --json
rtk task release-smoke
rtk xlflow save --session --json
rtk xlflow status --json
rtk xlflow session stop --json
rtk task release-package TAG=v1.0.0
```

This sequence updates and saves the tracked `build/Book.xlsm` development workbook. For a verification run that must not save, omit `save` and finish with `session stop --discard`.

If a user-owned workbook is already open, do not run `session start`; use `session attach` only under the user's explicit control. Do not save or close a user-owned workbook when verification is complete.

## Release asset check

`task release-package` generates these two files:

- `artifacts/release/VBA-Release-v1.0.0.zip`
- `artifacts/release/VBA-Release-v1.0.0.zip.sha256`

Verify the checksum file before uploading the release. Attach the ZIP and checksum file to the same GitHub tag. macOS, Windows 32-bit, and Excel/VBE compilation on GitHub-hosted runners remain unverified for this v1.0 gate.
