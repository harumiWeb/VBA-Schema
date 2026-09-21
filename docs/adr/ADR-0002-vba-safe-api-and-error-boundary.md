# ADR-0002: VBA予約語を避けたAPIと失敗境界

## Status

`accepted`

## Background

当初のZod風APIには`Any`、`String`、`Boolean`、`Date`、`Object`、`Optional`、`Integer`など、VBAの予約語または型名と衝突する名前が含まれていた。Windows 64-bit OfficeのVBE compile oracleでは、少なくとも`Any`、`Optional`、`Integer`をprocedure名にした宣言がコンパイルに失敗した。

また、`SafeParse`がすべてのruntime errorをvalidation failureへ変換すると、入力不正、schema構築ミス、runtime component不足、library bugを利用者が区別できなくなる。

## Decision

- v1のfactory名は`AnyValue`、`Text`、`Number`、`Bool`、`DateTime`、`ObjectSchema`、`ArrayOf`、`Literal`、`EnumOf`、`UnionOf`とする。
- v1のmodifier名は`OptionalField`、`Nullable`、`Min`、`Max`、`Length`、`WholeNumber`、`Pattern`、`Email`、`Field`、`Strict`とする。
- objectを含むUnion入力はVariant配列ではなく、`VSchema`を含む`Collection`へ限定する。
- Public APIの追加・変更は、全signatureとfluent chainを含むfixtureをVBEでcompileしてから採用する。
- `SafeParse`がResultへ変換するのはvalidation failureだけとする。
- schema構築のprogrammer misuseとruntime/environment failureは、それぞれ区別可能なErr番号範囲で投げる。
- `InternalInitialize`はVBAのclass初期化制約のためPublicで残すが、unsupported internal-only APIとする。Schema factoryだけが内部encoded kindを渡せ、通常の直接呼び出し、不正kind、再初期化はprogrammer misuseとして拒否する。
- VBAのPrivate member可視性制約を越えてchild schemaを合成・検証・循環検査する`Internal*` hookもPublicで残すが、`InternalInitialize`と同じunsupported internal-only APIとし、安定したuser-facing compatibility guaranteeの対象外とする。詳細なObject境界はADR-0003で定める。
- Numberの比較は型対応のlossless wideningとround-trip確認を優先し、precision loss・overflow・NaN・Infinityを黙って丸めない。ErrorTextと`received`は固定文法で生成し、constraint評価順も固定する。
- 暗黙の型変換を行わず、Literal/Enumを含む型比較は`docs/specs/v1-contract.md`のcategory規則に従う。

## Consequences

- Zodと同じ短い名前にはならないが、VBAでコンパイル可能なfluent APIを安定して提供できる。
- `SafeParse`でもenvironment failureやlibrary bugはErrを投げるため、呼び出し側が必要に応じて運用上のerror handlingを行う必要がある。
- Public API名はbreaking-change surfaceになるため、実装前のcompile fixtureが必須になる。
- Union定義には`Collection`の準備が必要になるが、object arrayの曖昧なsemanticsを公開契約から排除できる。

## Rationale

- Tests: `PublicApiCompile.bas`で全Public signatureとfluent chainをcompileし、M2 focused testsでscalar Result、Issue snapshot、Error boundaryを検証する。2026-09-21にWindows 64-bit OfficeのVBE compileとfocused/full testを実行済み。
- Code: `Schema.bas`、`VSchema.cls`、`VValidationResult.cls`にscalar coreを実装済み。Object/Array/Literal/Enum/Union/Pattern/Emailは後続milestoneで実装する。
- Related specs: `docs/specs/v1-contract.md`

## Supersedes

- None

## Superseded by

- None
