# ADR-0007: Release assets, version authority, and CI boundary

## Status

`accepted`

## Background

VBA-Schemaは利用者がVBEへimportするproduction sourceを3ファイルに限定している。一方、README、LICENSE、CHANGELOG、sample workbook、検証情報は利用者やrelease管理に必要だが、VBE import payloadへ混ぜると導入契約が不明確になる。

また、Excel/VBEを必要とするcompile・behavioral testはGitHub-hosted runnerで常に再現できない。Excelを持たないCIがcompile成功と誤表示しないこと、macOSとWindows 32-bitを未検証として正確に扱うことが必要である。

## Decision

- VBEへimportするrelease payloadは `Schema.bas`、`VSchema.cls`、`VValidationResult.cls` の3ファイルだけに固定する。
- README、LICENSE、CHANGELOG、sample workbook、ZIPなどの補助物は3-file payloadとは別のrepositoryまたはrelease assetとして提供する。同一payloadへ同梱する変更は、先にv1 contractとADR-0001を更新する。
- versionの正はGit tagとtracked `CHANGELOG.md`とする。production VBA sourceへversion定数や別のVERSIONファイルは追加しない。
- staged VBA sourceはUTF-8（BOMなし）、LF改行として検証する。release stagingは明示allowlistを使い、追加のreference設定なしにcompileできるfresh workbook smokeを通す。
- `tools/release-stage.ps1`は、Destinationがrepository root、production source tree（`src`）、xlflow管理領域（`.git`、`.xlflow`）、または検証用workbook（`build`）と重なる場合、削除処理より前に失敗させる。任意のrepository内パスを受け付ける場合でも、保護対象の子孫・祖先との重複を許可しない。
- GitHub-hosted CIはExcel不要の`lint`、`analyze`、format check、test discoveryを実行する。Windows Excel/VBE compileとbehavioral testsは別のlocal/self-hosted checkとして扱い、status名にも`vbe`または`excel`を含めて検証範囲を明示する。
- Excelなしのstatic CIをVBE compile passedとは表示しない。macOS/Windows 32-bit用matrixは実機検証を取得した時点で別jobとして追加する。
- GitHub Release workflowは`vMAJOR.MINOR.PATCH`のstable tag pushだけを受理し、再利用可能な`source-check`を先行実行する。成功後、`VBA-Release-vX.Y.Z.zip`を作成し、`VBA-Release/`直下へproduction 3ファイルだけを格納する。既存Releaseへの再実行は同名assetをclobberして更新する。

## Consequences

- import手順は常に3ファイルで維持され、補助資料の更新がVBA component数を変えない。
- versionのsingle sourceはGit tagとなるため、VBA project内から実行時にversionを取得するAPIはv1では提供しない。
- CIはExcel不要の失敗を早く検出できる一方、VBE compileの合否は別の実行環境に依存し、static CIだけでは証明できない。
- UTF-8/LFの検証によりrelease payloadのencoding差異を早期に検出できる。VBE import hostの実機差異はWindows 64-bit検証済み、macOS/32-bit未検証として残る。
- prerelease tagは明示的に対象外となるため、将来対応する場合はtag検証、CHANGELOG、Release note、pre-release表示を別の設計判断として更新する必要がある。

## Rationale

- Evidence: `docs/specs/v1-contract.md` と `tools/release-stage.ps1` / `tools/release-verify.ps1` が3-file allowlistを定義する。
- Evidence: `tools/release-smoke.ps1` はfresh workbook import、追加reference設定なしのVBE compile、scalar smoke、non-document component数を検証する。
- Evidence: `tasks/todo.md` のDG-008/DG-009とM7 release/CI tasksが、Excel-free検査とVBE検査を分離する。
- Related docs: `README.md`、`CHANGELOG.md`、`docs/design.md`。

## Supersedes

- None

## Superseded by

- None
