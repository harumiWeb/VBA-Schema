# xlflow issue: fmt output triggers parser recovery in XlflowAssert

## Status

Reproduced. The VBA-Schema repository works around it by excluding the existing xlflow helper from the formatter target.

## Environment

- Repository: `VBA-Schema`
- OS: Windows x64
- xlflow bridge: `xlflow-excel-bridge` x64, version `1.0.0`, commit `dev`
- Source: `src/modules/Xlflow/XlflowAssert.bas`

## Reproduction

1. Run this before applying the formatter.

   ```powershell
   rtk xlflow lint --json
   ```

   The result is success.

2. Run:

   ```powershell
   rtk xlflow fmt --write
   ```

   Ten files under `src/modules`, `src/workbook`, and tests are formatted.

3. Run the same lint command again.

   One finding appears at `src/modules/Xlflow/XlflowAssert.bas:1`:

   ```text
   VB014 parser recovery detected; inspect the reported source context before pushing to Excel.
   ```

   `rtk git diff --ignore-all-space -- src/modules/Xlflow/XlflowAssert.bas` shows no meaningful difference; the issue is reproduced after formatter-only whitespace and line-ending changes. Restoring the source before formatting makes lint pass again.

## Impact

Including existing xlflow helpers in the bulk formatter target causes source lint to fail and produces parser recovery unrelated to VBA-Schema production work.

## Workaround in this repository

`tools/check-format.ps1` targets only `Schema.bas`, `VSchema.cls`, `VValidationResult.cls`, and focused tests. Existing `src/modules/Xlflow` and scaffold workbook modules are excluded from the formatter gate. If the same issue occurs in new VBA-Schema source, report it as a separate issue instead of hiding it behind this workaround.

## Evidence

- `rtk xlflow analyze --json` succeeds both before and after formatting.
- `rtk xlflow test list --json` discovers the same five scaffold tests before and after formatting.
- Environment: workbook bootstrap was verified at `C:\temp\vba-schema-new-probe-20260921`.
