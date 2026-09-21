# Changelog

このファイルはGit tagによるreleaseと対応します。source payloadへversion定数は埋め込みません。

## Unreleased

- Union schema（branch Collection snapshot、nested Union、単一 `invalid_union` Issue）を追加。
- Literal、Enum、Pattern、Emailのscalar/category境界とruntime error境界を追加。
- 3ファイルimport payloadのstaging、検証、release smokeを追加。
- Object validationのlarge-input lookupをindex化し、boundary helperとproduction compatibility auditを追加。
- Windows 64-bit専用runtime benchmark、tracked baseline、Excel-free release staging/verificationを追加。
- `LICENSE`へ標準MITライセンス本文を追加。
- 利用者向けの注文・設定・APIレスポンスsampleと、sample専用のformat/lint/analyze検証を追加。
- `vMAJOR.MINOR.PATCH` tag pushでsource-checkを先行実行し、3モジュールだけを含む`VBA-Release-vX.Y.Z.zip`をGitHub Releaseへ添付するworkflowを追加。

## v1.0.0

未リリース。
