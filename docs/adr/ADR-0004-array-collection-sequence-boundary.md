# ADR-0004: Array／Collection sequence境界と安全な配列probe

## Status

`accepted`

## Background

VBAのnative arrayは`LBound`、一次元／多次元、未初期化dynamic array、typed object arrayなどの状態がVariant経由で曖昧になりやすい。JSON parserが返す`Collection`も同じ論理sequenceとして扱う一方、DictionaryはObject schema専用であり、arrayへ暗黙変換してはならない。

## Decision

- `ArrayOf(ItemSchema)`は一次元native arrayと`Collection`を受理し、`Scripting.Dictionary`や任意classをarrayとして扱わない。
- 未初期化dynamic arrayは0要素sequenceとして扱う。多次元arrayは`invalid_array_rank` issueを返し、runtime errorへ漏らさない。
- native arrayの実indexやCollectionの1始まりindexは公開pathへ持ち込まず、列挙順を0始まりのlogical indexへ正規化する。
- `Length`、`Min`、`Max`は要素schemaの再帰検証より先にsequence lengthへ適用し、最初のconstraint issueだけをarray nodeへ追加する。
- 要素検証はlogical index順でdepth-firstに行い、input array／Collectionを変更しない。Object、Nothing、Error Variantを含む要素はchild schemaの通常契約へ委譲する。
- 配列のrank／bounds取得は限定的なhelperでprobeし、未初期化判定に必要な`LBound`／`UBound`のErr stateをhelper内でclearして復元する。Windows API、`ObjPtr`、pointer-size依存は使用しない。

## Consequences

- 0始まり以外のarrayも`$[0]`から安定したIssue pathを得られる。
- 未初期化arrayの`LBound`／`UBound`例外をvalidation failureへ誤変換せず、空sequenceとして扱える。
- 大きなsequenceでは要素ごとのlate-bound accessが発生するため、性能目標はM7 benchmarkで確認する。
- macOS OfficeとWindows 32-bit Officeでも同じVBA標準機能だけで動作する構造だが、実機未検証・非保証である。

## Rationale

- Tests: `TestArray.bas`でzero／one／negative bound、Collection、nested sequence、logical path、length、uninitialized、multidimensional、typed object／Nothingを検証する。
- Code: `VSchema.cls`の`GetNativeArrayState`、`TryGetPrimaryArrayBounds`、`HasSecondArrayDimension`、`ValidateArray`が本決定を実装する。
- Related specs: `docs/specs/v1-contract.md`、`docs/design.md`

## Supersedes

- None

## Superseded by

- None
