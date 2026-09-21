# ADR-0006: Union branch ownership and error boundary

## Status

`accepted`

## Background

`UnionOf` は複数の `VSchema` を一つの schema node として扱う公開 API である。VBA の `Collection` は可変であり、入力 Collection をそのまま保持すると、schema 構築後の外部 mutation が validation の意味を変更する。一方、child `VSchema` は既存の Object/Array と同じ mutable builder 契約を持つため、deep clone を導入すると共有・cycle 検出・builder の一貫性が崩れる。

また、branchごとの失敗詳細をそのまま公開すると、branch数や内部構造に依存して Issue 契約が不安定になる。branch内部の programmer misuse や runtime failure を通常の validation failure と区別しない場合、利用者の schema 定義ミスや実行環境欠落を隠してしまう。

## Decision

- `UnionOf(ByVal Schemas As Collection)` は、構築時に非空 Collection の branch順と要素集合を snapshot する。
- snapshot の要素は `VSchema` の参照として保持し、child schema自体はcloneしない。したがって、構築後の child builder変更は Union の validation に反映される。
- `Nothing`、空 Collection、`VSchema` 以外の要素、未初期化 `VSchema` は programmer misuse として `vbObjectError + 2100` を送出する。
- validation は branch順に実行し、最初に成功した branch で成功する。nested Union は flatten せず、schema graph の構造を保持して再帰的に検証する。
- 全branchが validation failure の場合、branch内部の Issue は外部結果へ持ち出さず、元の失敗 path に一件の `invalid_union` Issue（expected=`Union`）を追加する。
- branch内部で発生した programmer misuse、runtime/environment failure、internal invariant failure は `invalid_union` に変換せず、そのまま送出する。

## Consequences

- Union に渡した Collection の後続 `Add`/`Remove` は既存 schema に影響しない。
- child `VSchema` の fluent builder は既存の共有参照契約どおり影響するため、childを共有した後の mutation は利用者が管理する必要がある。
- validation failure の Issue 数と形は branch 数に依存せず安定する。一方、branchごとの診断詳細は v1 の公開結果には含まれない。
- nested Union の再帰は schema graph preflight の対象となるため、cycle は `vbObjectError + 2103` として検出される。
- branchが利用する Dictionary/RegExp などの環境依存 component が欠落した場合、Union は環境エラーを隠さない。

## Rationale

- Tests: `src/modules/Tests/TestUnion.bas` は branch順、後続branch成功、単一 `invalid_union`、Collection snapshot、child参照共有、nested Union、special value、branch programmer error を検証する。
- Tests: `src/modules/Tests/TestErrors.bas` は空/Nothing/non-VSchema/未初期化 branch の `vbObjectError + 2100` を検証する。
- Code: `src/classes/VSchema.cls` の `InternalSetUnion`、`ValidateUnion`、`VisitSchemaGraph` がこの境界を実装する。
- Related specs: `docs/specs/v1-contract.md`、`docs/design.md`。

## Supersedes

- None

## Superseded by

- None
