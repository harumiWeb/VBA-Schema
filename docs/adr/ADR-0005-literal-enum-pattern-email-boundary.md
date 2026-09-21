# ADR-0005: Literal・Enum・Pattern・Email の値制約境界

## Status

`accepted`

## Background

`Literal` と `EnumOf` は Variant の暗黙 coercion に依存すると、String、Number、Boolean、Date、Null、Empty の境界が曖昧になる。`EnumOf` の入力配列を共有すると、schema 構築後の外部 mutation によって validation の意味が変わる。

`Pattern` と `Email` は正規表現 runtime を必要とする。VBA-Schema は Windows 64-bit Office を検証対象とし、macOS と Windows 32-bit は互換性を意識するが未検証・非保証とするため、正規表現の設定、失敗境界、Email の対象範囲を固定する必要がある。

## Decision

- `Literal` の候補は String、Number、Boolean、Date、Null、Empty の scalar category に限定する。Error Variant、Object、Array は schema 構築時に `vbObjectError + 2100` とする。
- `EnumOf` は一つの一次元 native array を受け取り、構築時に候補値を正規化して内部 snapshot する。typed scalar array は受理し、未初期化配列、多次元配列、空配列、Error/Object/Array 要素は `vbObjectError + 2100` とする。
- `EnumOf` の候補重複は無意味な定義として `vbObjectError + 2101` とする。String は binary comparison、Number は DG-002 の lossless comparison、他の category は同一 category の値比較を使う。
- `Pattern` は `VBScript.RegExp` を late binding し、`IgnoreCase=False`、`Global=False`、`MultiLine=False` に固定する。Expression は入力全体に一致する場合だけ成功する。RegExp は validation 時に lazy compile する。
- Pattern expression の compile 失敗は validation Issue へ変換せず、最初の validation 時に `vbObjectError + 2100` とする。RegExp component の生成失敗は `vbObjectError + 2200` とする。
- `Email` は VBScript.RegExp による実用的な ASCII 簡易形式とし、local と domain を `@` で一つだけ区切り、全体を 254 UTF-16 code unit 以下に制限する。local の一般的な ASCII atom と単一ドット区切り、domain の ASCII label とドット区切りを受理し、空、local-only、複数 `@`、空白、Unicode、quoted local、comment、IP literal、RFC の全ての拡張は受理対象外とする。
- `Pattern` と `Email` を併用した場合の評価順は既存 contract どおり Pattern、Email とする。いずれも Text schema にだけ適用できる。
- Enum の child schema や Union の Collection ownership はこのADRの対象外とし、Union 部分は M6 の ADR-0006 で別途決める。

## Consequences

- Enum 定義後の入力配列 mutation は validation 結果へ影響しない。
- candidate の category 境界と duplicate 定義が deterministic になり、`"1"` と `1`、Boolean と Number、Null と Empty を混同しない。
- Pattern は検索一致ではなく全体一致であるため、部分文字列検索には明示的な expression が必要になる。
- Email は典型的な入力ミス検出を目的とし、RFC 完全準拠を提供しない。README で範囲を明記する。
- VBScript.RegExp が利用できない host では validation failure ではなく environment failure が発生する。macOS/32-bit の実機証拠は別途必要である。

## Rationale

- Tests: M5 focused tests で category、snapshot、duplicate、Pattern flags、full-match、Email 境界を固定する。
- Code: `VSchema.cls` に scalar comparison、Enum snapshot、lazy RegExp creation を集約し、外部参照設定を追加しない。
- Related specs: `docs/specs/v1-contract.md`、`docs/design.md`

## Supersedes

- None

## Superseded by

- None
