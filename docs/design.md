# VBA-Schema Design Specification

公開API、validation semantics、error contract、compatibility、配布境界の規範仕様は`docs/specs/v1-contract.md`とする。本書は構想、構造、実装方針を説明する。

## 1. Overview

### 1.1 Project name

仮称:

**VBA-Schema**

A lightweight schema validation library for VBA.

### 1.2 Concept

VBA-Schemaは、TypeScriptのZodに着想を得た、VBA向けの宣言的なランタイムバリデーションライブラリである。

以下のようなデータに対して、スキーマを定義し、構造・型・制約を検証できることを目的とする。

* `Variant`
* `String`
* 数値型
* `Boolean`
* `Date`
* `Collection`
* 配列
* `Scripting.Dictionary`
* JSONパーサによって生成されたDictionary / Collection構造
* Excelセルや外部APIから取得した値

典型的な利用例:

```vb
Dim UserSchema As VSchema

Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text().Min(1).Max(100)) _
    .Field("age", Schema.Number().WholeNumber().Min(0)) _
    .Field("email", Schema.Text().Email()) _
    .Field("tags", Schema.ArrayOf(Schema.Text()).OptionalField())

Dim Result As VValidationResult
Set Result = UserSchema.SafeParse(UserData)

If Not Result.Success Then
    Debug.Print Result.ErrorText
End If
```

---

# 2. Design Goals

## 2.1 Primary goals

VBA-Schemaは以下を最優先する。

### Small distribution footprint

配布物は必ず以下の3モジュール以内とする。

```text
Schema.bas
VSchema.cls
VValidationResult.cls
```

**4個目のVBAモジュールを追加してはならない。**

新しいschema type、constraint、error type等を追加する場合も、この3モジュール内部で実装する。

### No external references required

外部ライブラリへの参照設定を要求しない。

必要に応じてlate bindingを使用する。

例:

```vb
CreateObject("Scripting.Dictionary")
CreateObject("VBScript.RegExp")
```

ユーザーに以下を要求してはならない。

```text
Microsoft Scripting Runtime
Microsoft VBScript Regular Expressions
```

late-bound runtime componentまで存在しないという意味の「依存ゼロ」ではない。対応環境の表現は`docs/specs/v1-contract.md`に従う。

### Easy installation

通常のVBAユーザーが3ファイルをインポートするだけで利用できること。

Package managerやxlflowを必須にしてはならない。

### Fluent API

スキーマを宣言的かつ連鎖的に定義できること。

```vb
Schema.Text().Min(1).Max(100)
```

### Useful errors

単なるBooleanではなく、

* どこで
* 何を期待して
* 実際に何を受け取り
* なぜ失敗したか

を取得できること。

---

# 3. Non-goals

v1では以下を対象外とする。

## 3.1 Compile-time type system

VBAの静的型検査機構を構築するものではない。

VBA-Schemaはruntime validation libraryである。

## 3.2 Zod API完全互換

ZodのAPIをそのまま再現しない。

特に以下はv1では実装しない。

* `refine`
* `superRefine`
* 任意callback validator
* `transform`
* preprocess pipeline
* discriminated union optimization
* recursive/lazy schema
* branded types
* Promise / async
* schema inferenceによるVBA型生成

## 3.3 Class-per-schema architecture

以下のような設計は禁止する。

```text
VStringSchema.cls
VNumberSchema.cls
VObjectSchema.cls
VArraySchema.cls
VUnionSchema.cls
VValidationIssue.cls
```

すべて`VSchema.cls`で表現する。

---

# 4. Hard Architecture Constraint

## Decision summary: Maximum Three VBA Modules

この判断の理由と互換性上のtrade-offは`docs/adr/ADR-0001-small-distribution-and-portable-core.md`に記録する。

配布物は以下のみとする。

```text
Schema.bas
VSchema.cls
VValidationResult.cls
```

理由:

1. VBAではライブラリ導入時にモジュール数がUXへ直接影響する。
2. ファイル数が多いライブラリは既存Workbookへ導入しづらい。
3. コピー・削除・アップデートが容易であることを重視する。
4. 小規模ライブラリとしての訴求力を維持する。
5. VBA-JSON等と同様の「簡単に持ち込める」導入体験を目指す。

この制約は内部設計の美しさより優先される。

ただし、コード品質を保つため`VSchema.cls`内部ではprivate procedureを責務ごとに明確に分割する。

---

# 5. Module Architecture

```text
┌───────────────────────┐
│      Schema.bas       │
│                       │
│ Public factory API    │
└───────────┬───────────┘
            │ creates
            ▼
┌───────────────────────┐
│      VSchema.cls      │
│                       │
│ Schema definition     │
│ Constraint storage    │
│ Recursive validation  │
│ Type handling         │
└───────────┬───────────┘
            │ produces
            ▼
┌───────────────────────────┐
│ VValidationResult.cls     │
│                           │
│ Success                   │
│ Value                     │
│ Issues                    │
│ ErrorText                 │
└───────────────────────────┘
```

---

# 6. Public API

## 6.1 Schema.bas

`Schema.bas`はユーザー向けFactory APIのみを提供する。

Public stateを持たない。

### Required factories

```vb
Schema.AnyValue()
Schema.Text()
Schema.Number()
Schema.Bool()
Schema.DateTime()
Schema.ObjectSchema()
Schema.ArrayOf(ItemSchema)
Schema.Literal(Value)
Schema.EnumOf(Values)
Schema.UnionOf(Schemas)
```

これらの名称は予約語を避けたv1の固定候補であり、全Public APIを使用するcompile-only fixtureで検証する。2026-09-21時点でWindows 64-bit OfficeのVBE compileを通過している。判断理由は`docs/adr/ADR-0002-vba-safe-api-and-error-boundary.md`、詳細なsignatureとchainは`docs/specs/v1-contract.md`を参照する。

---

# 7. VSchema

## 7.1 Responsibility

`VSchema`は以下を担当する。

* schema kind保持
* modifier保持
* constraint保持
* child schema保持
* object fields保持
* validation実行
* nested validation
* error生成

---

# 8. Schema Kinds

内部的には一つのEnumでschema typeを表現する。

```vb
Private Enum SchemaKind
    skAny = 0
    skString
    skNumber
    skBoolean
    skDate
    skObject
    skArray
    skLiteral
    skEnum
    skUnion
End Enum
```

Enumは外部公開しない。

factoryからprivate initializerを呼べないため、Publicだがunsupportedな内部初期化APIを使用する。通常のschema kind値ではなく、Schema.basだけが生成する内部encoded kindを渡す。

例:

```vb
Public Function InternalInitialize(ByVal KindCode As Long) As VSchema
```

このprocedureは一度だけ呼び出せるものとし、通常の直接呼び出し・encoded kindの不正値は`vbObjectError + 2100`、再初期化は`vbObjectError + 2102`、直接生成した未初期化schemaのvalidationはprogrammer misuseとしてErrを投げる。これはsecurity boundaryではなく、利用者向けの安定Public APIには含めない。

VBAでは同じclassの別instanceのPrivate memberを参照できないため、child schemaの合成・path付きnested validation・schema graph preflightに必要な`Internal*` hookもPublicで提供する。これらもunsupported internal-only APIであり、利用者向けの安定Public APIには含めない。

---

# 9. Schema State

`VSchema`は概ね以下のstateを保持する。

```vb
Private mKind As Long

Private mOptional As Boolean
Private mNullable As Boolean

Private mHasMin As Boolean
Private mMinValue As Variant

Private mHasMax As Boolean
Private mMaxValue As Variant

Private mHasLength As Boolean
Private mLengthValue As Long

Private mIntegerOnly As Boolean
Private mEmail As Boolean

Private mHasPattern As Boolean
Private mPattern As String
Private mHasLiteral As Boolean
Private mLiteralValue As Variant
Private mEnumValues As Variant
Private mPatternRegExp As Object
Private mEmailRegExp As Object

Private mFields As Object
Private mFieldOrder As Collection
Private mItemSchema As VSchema

Private mUnionSchemas As Collection

Private mObjectMode As Long
```

必ずしも上記と完全一致する必要はないが、**constraintごとの専用クラスを作成してはならない。**

---

# 10. Fluent Modifiers

## 10.1 OptionalField

```vb
Schema.Text().OptionalField()
```

Missing fieldを許可する。

Object fieldとして使用された場合に意味を持つ。

単体値に対して`Empty`をOptionalとして扱うかどうかは曖昧にしない。

v1では:

```text
Missing field ≠ Empty
```

とする。

Dictionaryにキー自体が存在しない場合のみOptionalとして許可する。

---

## 10.2 Nullable

```vb
Schema.Text().Nullable()
```

`Null`を許可する。

---

## 10.3 Min

String:

```vb
Schema.Text().Min(3)
```

文字列長の最小値。

Number:

```vb
Schema.Number().Min(0)
```

数値の最小値。

Array:

将来的に配列要素数へ適用可能だが、v1では実装対象としてよい。

---

## 10.4 Max

`Min`と同様。

---

## 10.5 Length

```vb
Schema.Text().Length(8)
```

厳密な文字列長。

Arrayへの利用も許可してよい。

---

## 10.6 WholeNumber

```vb
Schema.Number().WholeNumber()
```

整数のみ許可する。

以下はtrue:

```text
1
0
-10
```

以下はfalse:

```text
1.5
10.01
```

VBA型が`Double`でも数学的に整数であれば許可する。

---

## 10.7 Positive / Negative

v1では公開しない将来候補とし、最短でもv1.1以降とする。

実装する場合:

```vb
.Positive()   ' > 0
.NonNegative() ' >= 0
.Negative()   ' < 0
.NonPositive() ' <= 0
```

既存の`Min/Max` validationへ変換してもよい。

---

## 10.8 Pattern

```vb
Schema.Text().Pattern("^[A-Z]{3}-\d{4}$")
```

`VBScript.RegExp`をlate bindingで使用する。`IgnoreCase=False`、`Global=False`、`MultiLine=False`に固定し、Expressionは入力全体への一致だけを成功とする。

RegExp instanceはschema生成時ではなくvalidation時またはlazy cacheで生成する。不正なExpressionは最初のvalidation時にprogrammer misuseとしてErrを投げ、runtime component不足はenvironment failureとする。

---

## 10.9 Email

```vb
Schema.Text().Email()
```

Email validationはRFC完全準拠を目指さない。localとdomainを一つの`@`で区切るASCII簡易形式、全体長254 UTF-16 code unit以下とし、空、local-only、複数`@`、空白、Unicode、quoted local、comment、IP literalは受理しない。

目的は典型的な入力ミス検出であり、過剰に厳格なRFC実装を提供することではない。

---

# 11. Object Schema

## 11.1 Definition

```vb
Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text()) _
    .Field("age", Schema.Number().OptionalField())
```

内部ではfield nameと`VSchema`をDictionaryに保持する。

```vb
Private mFields As Object
```

validation issueの順序をDictionary列挙順に依存させないため、field追加順は別の`Collection`にも保持する。

生成:

```vb
Set mFields = CreateObject("Scripting.Dictionary")
```

---

## 11.2 Input

v1で正式サポートするObject input:

```text
Scripting.Dictionary
```

late-bound Dictionaryも含む。

`Collection`をkey-value objectとして扱わない。

一般VBA class instanceのreflectionは実施しない。

---

# 12. Unknown Field Policy

Objectはunknown fieldの扱いを制御する。

## Default

```text
Passthrough
```

schemaに存在しないfieldがあってもvalidation成功。

## Strict

```vb
Schema.ObjectSchema() _
    .Field("name", Schema.Text()) _
    .Strict()
```

未定義fieldが存在すればvalidation error。

## Strip

v1では実装しない将来候補とする。

なぜなら入力データ自体を書き換えるmutation semanticsが複雑になるため。

---

# 13. Array Schema

```vb
Schema.ArrayOf(Schema.Text())
```

対応対象:

* VBA native array
* `Collection`

可能であればJSON parserが生成するCollectionにも対応する。

Dictionaryをarrayとして扱ってはならない。

---

# 14. Literal Schema

```vb
Schema.Literal("active")
```

比較にはVBAのVariant比較による曖昧なcoercionを避ける。`Literal`と`EnumOf`の候補は構築時にscalar categoryを検証し、`EnumOf`は一次元配列の内容をsnapshotする。

以下は原則異なるものとして扱う。

```text
"1"
1
True
```

Literal比較では型も考慮する。

---

# 15. Enum Schema

推奨API:

```vb
Schema.EnumOf(Array("pending", "active", "disabled"))
```

受信値が候補のどれかと一致すれば成功。

comparison semanticsはLiteralと同じ。候補の重複はprogrammer misuseとして拒否する。

---

# 16. Union Schema

v1ではschema objectを含む`Array()`を公開契約にせず、`Collection`を使用する。

```vb
Dim Options As New Collection

Options.Add Schema.Text()
Options.Add Schema.Number()

Set S = Schema.UnionOf(Options)
```

`UnionOf`は入力`Collection`のbranch順を構築時にsnapshotする。ただし各`VSchema` instance自体はcloneせず、child schemaへの参照を保持するため、構築後のchild builder変更はUnionにも反映される。空のCollection、`Nothing`、`VSchema`以外、未初期化`VSchema`はprogrammer misuseとして`vbObjectError + 2100`を送出する。

Union validationはbranch順に試し、一つでも成功すれば成功する。nested Unionはflattenせず、構造を保持したまま再帰的に検証する。全branch失敗時はbranchごとの詳細Issueを外部へ漏らさず、失敗pathに一件の`invalid_union` Issueを返す。branch内部のprogrammer/environment errorはvalidation failureへ変換せず、そのまま送出する。

---

# 17. Validation API

## 17.1 SafeParse

基本API。

```vb
Dim Result As VValidationResult
Set Result = UserSchema.SafeParse(Data)
```

失敗時にもVBA runtime errorを投げない。

---

## 17.2 Parse

v1では実装しない将来候補とする。

```vb
Value = UserSchema.Parse(Data)
```

validation failure時に`Err.Raise`する。

VBAではObject/Variant return semanticsが複雑になるため、v1のvalidation APIは`SafeParse`だけとする。`Parse`は最短でもv1.1以降に再検討する。

---

# 18. VValidationResult

## 18.1 Public API

最低限以下を提供する。

```vb
Result.Success
Result.Value
Result.Issues
Result.ErrorText
```

### Success

```vb
Public Property Get Success() As Boolean
```

### Value

元のvalidated valueを返す。

v1ではtransformを行わないため、基本的にinput valueそのもの。

Objectの場合もdeep copyを行わない。

validation failure時は`Empty`を返す。

### Issues

`Collection`を返す。

各issueはlate-bound `Scripting.Dictionary`として表現する。

専用`VValidationIssue.cls`は作成しない。

返却値はsnapshotとし、呼び出し側の変更がResult内部へ影響してはならない。

---

# 19. Validation Issue Structure

各issueは最低限以下を持つ。

```text
path
code
message
expected
received
```

例:

```text
path     = "$.users[2].email"
code     = "invalid_email"
message  = "Invalid email address"
expected = "email"
received = "foo"
```

`received`はObject参照や配列そのものではなく、安全なString要約とする。Issueの型、順序、snapshot semanticsは`docs/specs/v1-contract.md`に従う。

required fieldが存在しない場合の`received`は`Missing`とする。present `Empty`や`Null`とは異なる状態であり、`OptionalField`はmissingだけを許可する。

v1では`received`と`ErrorText`の文法をspecで固定する。Stringのescape/truncate、Numberのlocale非依存表現、Dateの秒精度、Issue行の改行を実装ごとに変えてはならない。

Dictionary例:

```vb
Issue("path")
Issue("code")
Issue("message")
Issue("expected")
Issue("received")
```

---

# 20. Error Codes

error message文字列ではなく、machine-readableなcodeを必ず持たせる。

初期code:

```text
required
invalid_type
too_small
too_big
invalid_length
invalid_integer
invalid_pattern
invalid_email
invalid_literal
invalid_enum
invalid_union
invalid_key
unknown_field
invalid_array_rank
```

codeは将来的なbreaking changeを避けるためREADMEで公開仕様とする。

---

# 21. Path Format

nested error locationは以下の形式に統一する。

すべてのpathはroot記号`$`から開始する。

Object:

```text
$.user.name
```

Array:

```text
$.users[2]
```

Nested:

```text
$.users[2].address.zip
```

identifier形式でないfield名:

```text
$["field.with.dot"]
```

ArrayとCollectionは実際の`LBound`や1始まりindexにかかわらず、列挙順を0始まりの論理indexへ正規化する。

v1では一次元native arrayだけを受理する。多次元配列は`invalid_array_rank`とする。

v1ではpath segment専用クラスを持たない。

---

# 22. ErrorText

`VValidationResult.ErrorText`は人間向け表示を返す。

形式は`<path>: <message> (expected=<expected>, received=<received>)`とし、Issue間は`vbCrLf`で連結する。raw input全体やsecretを追加表示しない。

例:

```text
$.users[2].email: Invalid email address (expected=email, received=String("foo"))
$.users[2].age: Number is below the minimum. (expected=Number >= 18, received=Number(16))
$.settings.timeout: The value has an invalid type. (expected=Number, received=String("slow"))
```

フォーマットはテストで固定する。

ただしIssue collectionが正式なmachine-readable APIであり、`ErrorText`はpresentation APIと位置づける。

---

# 23. Type Semantics

ここはVBA特有の曖昧さがあるため厳密に定義する。

## String

基本的には:

```vb
VarType(Value) = vbString
```

のみ許可する。

数値から文字列への自動変換を行わない。

---

## Number

以下をNumberとして扱う。

```text
Byte
Integer
Long
Single
Double
Currency
```

Decimal VariantとLongLongは、32-bit/64-bitで同じ契約を検証できるまでv1では受理しない。

constraint boundaryはVariantで保持し、無条件にDoubleへ変換しない。

numeric comparisonは型対応のlossless wideningとround-trip確認を行い、precision loss、overflow、NaN、Infinityを黙って成功扱いしない。constraint評価順はspecの固定順に従い、nodeごとに最初の1件だけをIssueへ追加する。

Booleanは数値として扱わない。

Dateも内部的には数値だがNumberとして扱わない。

---

## Boolean

`vbBoolean`のみ。

`0/-1`をBooleanへ自動変換しない。

---

## Date

`vbDate`のみ。

日付文字列を自動parseしない。

---

## Null

Nullable指定時のみ許可。

---

## Empty

Nullとは区別する。

v1では通常のschemaに対してvalidation failureとする。

`AnyValue()`のみ許可。

---

## Error Variant

`AnyValue()`では許可し、それ以外ではvalidation failure。

---

# 24. No Implicit Coercion

VBA-Schemaの重要な設計原則。

以下を自動変換してはならない。

```text
"123" → Number
123 → String
0 → Boolean
"2026-01-01" → Date
```

理由:

validation libraryが暗黙変換を行うと、データ境界での不正値を隠してしまうため。

将来的にcoercion APIを提供する場合は明示APIとする。

例:

```text
Schema.Coerce.Number()
```

ただしv1対象外。

---

# 25. Validation Algorithm

概念的には以下。

```text
Validate(schema, value, path)

1. Null / Empty / Optional処理
2. schema kind dispatch
3. 基本型validation
4. schema-specific constraint
5. nested validation
6. issue collectionへ追加
```

pseudo code:

```vb
Private Sub ValidateValue( _
    ByVal Value As Variant, _
    ByVal Path As String, _
    ByRef Issues As Collection)

    If HandleNullability(...) Then Exit Sub

    Select Case mKind
        Case skAny
            Exit Sub

        Case skString
            ValidateString Value, Path, Issues

        Case skNumber
            ValidateNumber Value, Path, Issues

        Case skObject
            ValidateObject Value, Path, Issues

        Case skArray
            ValidateArray Value, Path, Issues

        ' ...
    End Select
End Sub
```

巨大な1 procedureにしてはならない。

kindごとにprivate validation procedureへ分割する。

---

# 26. Object Validation Algorithm

```text
ValidateObject(value, path):

1. `TypeName(value) = "Dictionary"`と`Count`／`Exists`／`Keys`の限定capabilityを確認する。失敗時は`invalid_type` issue
2. schema fieldsを列挙
3. `Keys`列挙と`StrComp(..., vbBinaryCompare)`でinputにkeyが存在するか確認
4. 無ければOptional判定
5. あればchild schemaをrecursive validation
6. String以外のinput keyは`.Strict()`にかかわらず`invalid_key`としてroot object pathへ追加
7. Strictの場合はString input keysをbinary ordinal順に列挙
8. schemaに存在しないkeyを`unknown_field`として追加
```

---

# 27. Array Validation Algorithm

```text
ValidateArray(value, path):

1. VBA arrayかCollectionか判定（Dictionaryはarrayとして扱わない）
2. 未初期化dynamic arrayは0要素、二次元以上は`invalid_array_rank`
3. length constraintを検証
4. 各要素に対してchild schemaを再帰的に実行
5. pathに0始まりlogical `[index]`を追加
```

native VBA arrayでは`LBound` / `UBound`を使用。

空配列や未初期化dynamic arrayでruntime errorを起こさないよう注意する。

このケース専用の安全なarray detection helperを実装し、probe中のErr stateをhelper内でclearしてから通常のvalidationへ戻す。input array／Collectionは変更しない。

---

# 28. Object Detection

VBAにはinterface reflectionがないため、v1ではDictionaryを明示的に対象とする。

v1の判定は次の入口と限定capability検査を組み合わせる。

```vb
TypeName(Value) = "Dictionary"
```

`Count`、`Exists`、`Keys`の読み取りに失敗した場合は入力の`invalid_type`とする。Issue用Dictionaryの生成自体が失敗する場合はruntime/environment error (`vbObjectError + 2200`)として伝播する。

`Object`を受け取り任意memberを試すexception-driven duck typingは行わず、TypeName後の限定helperだけで検査する。Input Dictionaryの`CompareMode`は変更せず、キー比較は常にbinaryとする。

String以外のkeyは`invalid_key`、strict unknown String keyは`unknown_field`とする。unknown fieldはpassthroughが既定であり、validation中にinputを変更しない。

---

# 28.1 Schema graph preflight

builderはmutableでschemaを共有できるため、`SafeParse`開始時にschema graphをDFSで検査する。active pathは`Collection`に保持し、schema identityはVBAの`Is`比較で判定する。direct／indirect cycleを検出した時点で`vbObjectError + 2103`を投げ、value validationやIssue生成を開始しない。shared childの非循環再利用は許可する。`ObjPtr`、Windows API、pointer-size依存は使用しない。

---

# 29. Performance Requirements

一般的なVBA validation用途で十分高速であること。

Windows 64-bitの同一machine、同一workbook、warm-up後5回のmedianで測定する初期目標:

```text
1,000 scalar fields:
500 ms未満

10,000 scalar validations:
1,000 ms未満
```

環境差が大きいため絶対時間だけで回帰判定せず、同一fixtureの確立済みbaselineに対する25%超の悪化も調査対象とする。

micro-optimizationより以下を優先する。

* runtime errorをvalidation flowに使わない
* RegExpの不要な再生成を避ける
* Dictionary lookupを活用
* 不要なdeep copyをしない
* error object生成を失敗時だけ行う

---

# 30. Mutation Policy

v1では入力を変更しない。

以下は禁止。

```text
unknown field削除
string trim
type coercion
default値注入
object再構築
```

したがって:

```vb
Result.Value
```

は原則として入力値そのものを返す。

---

# 31. Error Handling

validation failureとlibrary bugを区別する。

完全な分類とErr番号範囲は`docs/specs/v1-contract.md`に従う。

## Validation failure

`SafeParse`ではErrを投げない。

Result:

```text
Success = False
Issues.Count > 0
```

## Programmer misuse

例:

```vb
Schema.Text().Min(-1)
Schema.ArrayOf(Nothing)
.Field("", Nothing)
```

これはAPI misuseであるため`Err.Raise`してよい。

Error numberは独自範囲を定義する。

例:

```vb
vbObjectError + 2100
```

必要なruntime componentが存在しない場合や内部invariant違反は`vbObjectError + 2200`から`vbObjectError + 2299`を使用し、`SafeParse`でもvalidation failureへ変換しない。

---

# 32. Fluent API Mutation Semantics

builder methodは基本的に同一instanceを変更し、自身を返す。

```vb
Public Function Min(ByVal Value As Variant) As VSchema
    mHasMin = True
    mMinValue = Value
    Set Min = Me
End Function
```

これにより:

```vb
Schema.Number().Min(0).Max(100).WholeNumber()
```

を成立させる。

schema immutable設計はVBAではallocationと実装量を増やすため採用しない。

---

# 33. Schema Reuse Warning

mutable builderであるため、以下の挙動をREADMEで説明する。

```vb
Dim Base As VSchema
Set Base = Schema.Text()

Dim A As VSchema
Set A = Base.Min(1)
```

`A`と`Base`は同一object。

immutable clone APIはv1対象外。

---

# 34. Documentation Example

README冒頭は導入コストの低さを強調する。

```text
Modern schema validation for VBA.

3 files.
No reference setup.
No compile-time dependencies.
```

Example:

```vb
Dim UserSchema As VSchema

Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text().Min(1)) _
    .Field("age", Schema.Number().WholeNumber().Min(0)) _
    .Field("email", Schema.Text().Email())

Dim Result As VValidationResult
Set Result = UserSchema.SafeParse(Data)

If Result.Success Then
    Debug.Print "Valid"
Else
    Debug.Print Result.ErrorText
End If
```

VBE import payloadはproduction 3ファイルだけとし、README、LICENSE、CHANGELOG、sample workbook、ZIPなどは別のrelease/repository assetとして提供する。staged VBA sourceのencodingはUTF-8（BOMなし）、改行コードはLFに固定する。versionの正はGit tagとCHANGELOGとし、VBA sourceへversion定数は追加しない。

---

# 35. Compatibility

開発・検証対象:

```text
VBA7
Office 2016+
Windows 64-bit Office
```

Windows 32-bit OfficeとmacOS Officeはcompatibility-consciousとするが、検証環境がないため未検証・非保証とする。「対応済み」「サポート済み」と表現してはならない。

Windowsをprimary targetとする。

未検証環境との互換性を不必要に失わないよう、Windows API、pointer-size依存宣言、Excel Object Modelへのコア依存は禁止する。

`Scripting.Dictionary`と`VBScript.RegExp`はlate bindingするが、未検証環境で利用可能であることは保証しない。

必要なruntime componentが利用できない場合はvalidation failureではなくenvironment failureとする。

---

# 36. Testing Strategy

テストはxlflowを使用して自動実行可能にする。

最低限以下のカテゴリを用意する。

```text
tests/
├── TestString.bas
├── TestNumber.bas
├── TestBoolean.bas
├── TestDate.bas
├── TestObject.bas
├── TestArray.bas
├── TestOptional.bas
├── TestNullable.bas
├── TestLiteral.bas
├── TestEnum.bas
├── TestUnion.bas
├── TestErrors.bas
└── TestIntegration.bas
```

テストモジュール数は配布物の3モジュール制約には含めない。

---

# 37. Required Test Cases

## Strings

```text
valid string
wrong type
Min boundary
Max boundary
Length
empty string
Pattern pass/fail
Email pass/fail
Null
Empty
length boundary accepts `3`, `3&`, `CByte(3)`, `3#`
length boundary rejects fraction / negative / overflow / non-numeric
```

## Numbers

```text
Byte
Integer
Long
Single
Double
Currency
minimum
maximum
integer
fraction
Boolean rejection
Date rejection
numeric string rejection
Number Min/Max boundary subtype coverage
```

## Object

```text
required field present
required field missing
optional field missing
nested object
multiple errors
strict unknown field
dictionary input
input Dictionary with `CompareMode = vbTextCompare` still uses binary field matching
wrong input type
```

## Array

```text
native array
Collection
empty array
nested array
wrong element
correct error index
```

## Literal

```text
correct value
wrong value
same textual representation but different type
```

## Union

```text
first branch success
later branch success
all branches fail
nested union
```

## Error path

```text
$
$.user.name
$.users[0]
$.users[2].email
$.orders[1].items[4].price
```

---

# 38. Regression Testing

すべてのbug fixには再現テストを追加する。

validation semanticsを変更する修正は必ず既存テストへの影響を確認する。

AIエージェントはbug修正時にproduction codeだけ変更してはならない。

---

# 39. Code Quality Requirements

全module:

```vb
Option Explicit
```

必須。

以下を避ける。

```text
Select
Activate
On Error Resume Next の広範囲利用
暗黙Variant
Public mutable field
unqualified Excel object references
Windows API
```

Excel object modelそのものへの依存は原則禁止する。

VBA-SchemaはExcel専用ではなく、VBA runtime libraryとして設計する。

---

# 40. xlflow Dogfooding

開発にはxlflowを使用する。

CIまたはローカル開発で最低限:

```text
xlflow test
xlflow lint
xlflow fmt
```

を通す。

可能であればVBE compile oracleも利用する。

APIサンプルについては実際にExcel/VBEでcompileできることを検証する。

---

# 41. Repository Structure

正規配置:

```text
vba-schema/
├── src/
│   ├── modules/
│   │   ├── Schema.bas
│   │   ├── Tests/
│   │   └── Xlflow/
│   ├── classes/
│   │   ├── VSchema.cls
│   │   └── VValidationResult.cls
│   └── workbook/
│
├── sample/
│   ├── 01-order-import/
│   │   ├── README.md
│   │   └── SampleOrderImport.bas
│   ├── 02-settings-validation/
│   │   ├── README.md
│   │   └── SampleSettingsValidation.bas
│   ├── 03-api-payload/
│   │   ├── README.md
│   │   └── SampleApiPayload.bas
│   └── README.md
│
├── docs/
│   ├── adr/
│   ├── specs/
│   └── design.md
│
├── dist/
│   └── VBA-Schema/
│       ├── Schema.bas
│       ├── VSchema.cls
│       └── VValidationResult.cls
│
├── README.md
├── LICENSE
└── THIRD_PARTY_NOTICES.md
```

依存物がなければ`THIRD_PARTY_NOTICES.md`は不要。

`src/`は開発・テスト用、`sample/`は利用者向け補助サンプル、`dist/VBA-Schema/`は3ファイルだけを含むユーザー配布用、`.xlsm`はxlflow検証用成果物として区別する。現行のxlflow source tree全体から直接3-component workbookを生成せず、3ファイルのallowlistからrelease staging projectを作る。`sample/`の`.bas`は配布payloadへ混在させない。
release stagingのDestinationはrepository root、`src`、`.git`、`.xlflow`、`build`と重複してはならず、`tools/release-stage.ps1`は既存Destinationの削除より前にこの重複を拒否する。

---

# 42. Implementation Phases

## Phase 0 — Contract and compile fixture

実施:

```text
Public API compile-only fixture
VBE compile oracle
error/path/type contract固定
3-file release staging設計
allowlist staging scriptまたはTask
```

Acceptance criteria:

* `docs/specs/v1-contract.md`とPublic signatureが一致する
* 全factoryとfluent chainがWindows 64-bit Officeでcompile成功する
* 開発source、検証workbook、3-file distributionの境界が明示されている
* 3ファイルだけをstagingしrelease gateを実行できるscriptまたはTaskがある

---

## Phase 1 — Core infrastructure

実装:

```text
Schema.bas
VSchema.cls
VValidationResult.cls
```

対象schema:

```text
AnyValue
Text
Number
Bool
DateTime
```

対象API:

```text
SafeParse
OptionalField
Nullable
Min
Max
Length
WholeNumber
```

Acceptance criteria:

* compile成功
* scalar validation tests成功
* error path `$`
* zero external references
* exactly 3 distributed modules

---

## Phase 2 — Object validation

実装:

```text
Object
Field
nested Object
Optional field
Strict
```

Acceptance criteria:

```vb
Schema.ObjectSchema() _
    .Field("name", Schema.Text()) _
    .Field("age", Schema.Number().OptionalField())
```

が動作する。

nested path:

```text
$.user.address.zip
```

が正しく生成される。

---

## Phase 3 — Arrays

実装:

```text
ArrayOf
native VBA arrays
Collection
nested arrays
```

path:

```text
$.users[3].email
```

を生成できる。

---

## Phase 4 — Value constraints

実装:

```text
Literal
Enum
Pattern
Email
```

---

## Phase 5 — Union

実装:

```text
UnionOf
```

error outputを整理する。

---

## Phase 6 — Documentation and hardening

実施:

```text
README
API reference
Examples
edge cases
64-bit Office validation
32-bit Office compatibility review（実機が利用可能ならvalidation）
macOS compatibility review（実機が利用可能ならvalidation）
performance benchmark
```

ベンチマークは開発用の補助物であり、`src/modules/Benchmarks` と
`tools/run-benchmark.ps1` は3-file import payloadへ含めない。Windows 64-bit
Excelのfixture結果をtracked baselineと比較し、macOS／Windows 32-bitは
対応可能な実装を維持するが未検証・非保証とする。詳細なfixture、counter、
threshold、session ownershipは`docs/specs/benchmark-contract.md`と
`docs/adr/ADR-0008-runtime-benchmark-contract.md`を正とする。ベンチマークは
`office_bitness`が`x64`でない場合にbaseline比較・更新を行わず、unsupported
reportを残して失敗する。

---

# 43. MVP Definition

v1.0で最低限提供するもの:

```text
AnyValue
Text
Number
Bool
DateTime
ObjectSchema
ArrayOf
Literal
EnumOf
UnionOf

OptionalField
Nullable

Min
Max
Length
WholeNumber
Pattern
Email

Field
Strict

SafeParse

Success
Value
Issues
ErrorText
```

---

# 44. v1.0 Exclusions

明示的にv1.0へ入れない。

```text
Transform
Refine
SuperRefine
Callback validator
Default values
Coercion
Strip unknown fields
Recursive schemas
Class object reflection
OpenAPI generation
JSON parsing
HTTP
Schema serialization
Code generation
```

これらはコアAPIが安定した後に検討する。

---

# 45. Future Extensions

## OpenAPI integration

将来的に:

```text
OpenAPI Schema
      ↓
generated VBA-Schema definition
      ↓
VBA-HTTP response
      ↓
runtime validation
```

を実現可能。

例:

```vb
Set User = UserSchema.SafeParse(Response.Json)
```

ただしOpenAPI generatorは別プロジェクトまたはbuild-time toolingとし、VBA-Schema本体を肥大化させない。

---

# 46. Design Principles for AI Agents

AIエージェントは実装時に以下を厳守する。

### 1. Do not add modules

新規`.bas` / `.cls` / `.frm`をproduction sourceへ追加しない。

### 2. Prefer private procedures over classes

責務分離が必要なら:

```text
Private ValidateString
Private ValidateNumber
Private ValidateObject
Private ValidateArray
Private AddIssue
Private BuildPath
```

のように分割する。

### 3. Do not over-engineer

以下を導入しない。

```text
dependency injection
interface hierarchy
visitor pattern
factory classes
constraint classes
error classes
schema subclasses
```

### 4. Tests before extensions

新機能追加時:

```text
1. failing test
2. minimal implementation
3. full tests
4. lint
5. compile validation
```

の順を基本とする。

### 5. Preserve strict semantics

便利だからという理由で暗黙type coercionを追加しない。

### 6. Public API stability matters

Public method名やerror code変更はbreaking changeとして扱う。

---

# 47. Definition of Done

v1.0は以下をすべて満たした時点で完成とする。

* 配布対象production sourceが3 VBA componentsのみ
* External reference設定不要
* Windows 64-bit OfficeでVBE compile成功
* macOS / Windows 32-bitはcompatibility-conscious、未検証、非保証と明記
* Text / Number / Bool / DateTime対応
* Object / nested Object対応
* Array / Collection対応
* OptionalField / Nullable対応
* Literal / Enum / Union対応
* path付きvalidation errors
* machine-readable error codes
* human-readable `ErrorText`
* Strict object validation
* Pattern / Email
* xlflow test成功
* xlflow lint成功
* xlflow analyze成功
* READMEにinstallation / examples / API overviewあり
* 主要Public APIに回帰テストあり
* 配布ファイルを手動インポートして利用可能
* release artifactが`Schema.bas` / `VSchema.cls` / `VValidationResult.cls`の3ファイルだけ

---

# 48. Product Positioning

VBA-Schemaの価値は「VBAでZodを完全再現したこと」ではない。

価値は、

> Modern schema validation for VBA with almost zero installation cost.

を実現することにある。

プロジェクトとして維持すべき特徴は:

```text
3 files
no external reference setup
declarative schemas
strict runtime validation
structured errors
works with ordinary VBA projects
```

である。

機能追加によってこの特徴が失われる場合、その機能は本体へ追加しない。

**Small size is a feature, not an implementation detail.**
