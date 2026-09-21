# xlflow issue: fmt output triggers parser recovery in XlflowAssert

## Status

再現済み。VBA-Schema側では既存のxlflow helperをformatter対象から除外して回避する。

## Environment

- Repository: `VBA-Schema`
- OS: Windows x64
- xlflow bridge: `xlflow-excel-bridge` x64, version `1.0.0`, commit `dev`
- Source: `src/modules/Xlflow/XlflowAssert.bas`

## Reproduction

1. formatter適用前に次を実行する。

   ```powershell
   rtk xlflow lint --json
   ```

   結果はsuccess。

2. 次を実行する。

   ```powershell
   rtk xlflow fmt --write
   ```

   `src/modules`、`src/workbook`、testsの10ファイルが整形される。

3. 同じlintを再実行する。

   ```powershell
   rtk xlflow lint --json
   ```

   `src/modules/Xlflow/XlflowAssert.bas:1`で次の1件が発生する。

   ```text
   VB014 parser recovery detected; inspect the reported source context before pushing to Excel.
   ```

`rtk git diff --ignore-all-space -- src/modules/Xlflow/XlflowAssert.bas`では意味のある差分がなく、formatter-onlyの空白・改行変更後に再現する。整形前のsourceへ戻すとlintはsuccessへ戻る。

## Impact

既存のxlflow helperを一括formatter対象にすると、source lintが失敗し、VBA-Schema固有のproduction作業と無関係なparser recoveryが発生する。

## Workaround in this repository

`tools/check-format.ps1`は`Schema.bas`、`VSchema.cls`、`VValidationResult.cls`、focused testだけを対象にする。既存の`src/modules/Xlflow`、scaffold workbook modulesはformatter gateへ含めない。新しいVBA-Schema sourceで同じ問題が再現した場合は、この回避策で隠さず別issueとして切り出す。

## Evidence

- `rtk xlflow analyze --json`はformatter前後ともsuccess。
- `rtk xlflow test list --json`はformatter前後ともscaffold 5件を検出。
- 環境: `C:\temp\vba-schema-new-probe-20260921`でworkbook bootstrapを検証。
