# ADR-0003: Dictionary Object判定とschema graphの循環境界

## Status

`accepted`

## Background

v1のObject schemaは、外部入力をkey-valueとして扱うための唯一の実装対象を`Scripting.Dictionary`に限定する。VBAには一般的なinterface reflectionがなく、Dictionaryの`CompareMode`も入力ごとに異なり得る。また、mutable builderではschemaを共有した後に自己参照・間接循環が発生し得るため、validation中に無限再帰を起こさない境界が必要になる。

## Decision

- Object入力は`TypeName(value) = "Dictionary"`を入口とし、`Count`、`Exists`、`Keys`の限定的な能力検査を行う。
- TypeName不一致または能力検査失敗は`invalid_type` validation issueとする。VBA-Schema自身がIssue用Dictionaryを生成できないなど、libraryが要求する実行時componentの生成失敗は`vbObjectError + 2200`として伝播する。
- field照合は入力Dictionaryの`CompareMode`へ依存せず、`Keys`列挙と`StrComp(..., vbBinaryCompare)`でcase-sensitiveに行う。validation中に入力Dictionaryを変更しない。
- String以外の入力keyは`.Strict()`の有無にかかわらず`invalid_key` issueとする。pathはobject nodeのpath、`received`は安全なdescriptorとし、複数keyはdescriptorのbinary ordinal順に並べる。
- Stringのunknown fieldはdefaultでpassthroughとし、`.Strict()`時だけ`unknown_field` issueを追加する。unknown fieldのpathと`received`はcanonical path／`DescribeValue` grammarを使用し、binary ordinal順に並べる。
- required fieldの欠落は`received = "Missing"`とする。これは`Empty`や`Null`と異なる状態として公開grammarへ追加する。
- `SafeParse`開始時にschema graphをDFSでpreflightする。active pathは`Collection`に保持し、schema identityはVBAの`Is`比較で判定する。direct/indirect cycleはvalidation開始前に`vbObjectError + 2103`を投げ、shared childの非循環再利用は許可する。
- VBAでは同一classの別instanceのPrivate memberを参照できないため、composition・nested validation・cycle traversalに必要な`Internal*` hookをPublicで提供する。これらは`InternalInitialize`と同じunsupported internal-only APIであり、安定したuser-facing compatibility guaranteeの対象外とする。

## Consequences

- macOSおよびWindows 32-bit OfficeでもWindows APIやpointer-size依存なしに同じアルゴリズムを利用できるが、実機検証なしのため対応保証はしない。
- Dictionaryの`CompareMode`差異による大文字小文字の誤一致を防げる一方、field lookupごとにKeysを列挙するため、大規模Objectでは後続の性能測定で改善余地を確認する。
- non-string keyをpassthroughせず明示的なvalidation failureにするため、外部入力の契約違反を安全に報告できるが、従来の任意Variant keyを許可する利用者には移行上の注意が必要になる。
- `Internal*` hookはVBAの可視性制約を回避する実装境界であり、将来の公開APIとして扱わない。

## Rationale

- Tests: `TestObject.bas`でrequired／optional、nested path、binary matching、strict ordering、non-string key、input non-mutationを検証し、`TestErrors.bas`でbuilder misuseとdirect/indirect cycleのErr番号を検証する。
- Code: `VSchema.cls`の`GetDictionary`、`TryFindDictionaryKey`、`ValidateObject`、`EnsureAcyclicGraph`が本決定を実装する。
- Related specs: `docs/specs/v1-contract.md`、`docs/design.md`

## Supersedes

- None

## Superseded by

- None
