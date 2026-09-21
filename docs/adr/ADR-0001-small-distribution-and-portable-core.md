# ADR-0001: 3モジュール配布と移植可能なコア

## Status

`accepted`

## Background

VBA-Schemaは、既存のVBAプロジェクトへ容易に持ち込めるランタイムバリデーションライブラリを目指す。導入時のファイル数と参照設定は利用障壁になる一方、VBAには一般的なパッケージ境界や内部可視性がなく、過度なクラス分割は配布と更新を難しくする。

開発と検証はWindows 64-bit Officeで行う。macOS OfficeとWindows 32-bit Officeは手元に検証環境がないため、互換性を意識した実装はできても、現時点で動作保証はできない。

## Decision

- 配布対象のproduction componentは`Schema.bas`、`VSchema.cls`、`VValidationResult.cls`の3ファイルに限定する。
- `dist/VBA-Schema/`は上記3ファイルから生成するimport payloadであり、sourceと分離したcommit対象外のstaging生成物とする。
- テスト、xlflow実行基盤、サンプル、生成補助コードは配布対象に含めない。
- 外部参照設定を要求せず、利用可能な実行時コンポーネントはlate bindingする。
- Issueの公開表現は`Scripting.Dictionary`をv1の必須concrete typeとして固定する。portable fallbackはv1へ導入せず、component unavailable時はenvironment failureとして扱う。
- Windows API、Win32固有の型宣言、Excel Object Modelへのコア依存を追加しない。
- 検証済みのサポート対象はWindows 64-bit Officeとする。
- macOS OfficeとWindows 32-bit Officeは「互換性を意識するが未検証」と表示し、サポート済みとは表示しない。
- 未検証環境の互換性のために、検証済み環境のエラー契約や型安全性を弱めない。
- `Scripting.Dictionary`および`VBScript.RegExp`は参照設定不要でも実行時機能であるため、「依存なし」ではなく「外部参照設定不要」と表現する。

## Consequences

- 導入、削除、更新が3ファイルの操作で完結する。
- `VSchema.cls`へ責務が集中するため、private procedureと明確な内部領域による分割が必要になる。
- 専用Issue classやschema subtypeを追加できず、内部表現の型安全性には限界がある。
- macOSまたは32-bit Officeでの不具合報告は互換性改善の対象にはなるが、再現環境を得るまでは修正完了や対応済みを保証できない。
- late-bound runtime componentが存在しない場合は、validation failureとして隠さず、実行環境エラーとして扱う。

## Rationale

- Tests: `docs/specs/v1-contract.md`に配布smoke testとcompatibility gateを定義し、`tools/release-smoke.ps1`でfresh workbookへの3-file import、VBE compile、scalar smoke、non-document component 3件を検証する。
- Code: production sourceは`src/modules/Schema.bas`、`src/classes/VSchema.cls`、`src/classes/VValidationResult.cls`に配置し、`dist/VBA-Schema/`はallowlistから生成するcommit対象外payloadとする。
- Related specs: `docs/specs/v1-contract.md`

## Supersedes

- None

## Superseded by

- None
