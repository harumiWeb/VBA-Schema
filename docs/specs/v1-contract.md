# VBA-Schema v1 Contract

## 1. Status and authority

この文書はVBA-Schema v1の公開API、validation semantics、error contract、compatibility、配布境界を定義する。`docs/design.md`の例や概念説明と矛盾する場合、この文書を優先する。

## 2. Public API

### 2.1 Factory API

`Schema.bas`は次のfactoryを公開する。

```vb
Public Function AnyValue() As VSchema
Public Function Text() As VSchema
Public Function Number() As VSchema
Public Function Bool() As VSchema
Public Function DateTime() As VSchema
Public Function ObjectSchema() As VSchema
Public Function ArrayOf(ByVal ItemSchema As VSchema) As VSchema
Public Function Literal(ByVal ExpectedValue As Variant) As VSchema
Public Function EnumOf(ByVal Values As Variant) As VSchema
Public Function UnionOf(ByVal Schemas As Collection) As VSchema
```

- `ArrayOf`は`VSchema`を受け取る。
- `EnumOf`はscalar valueを含む一次元native arrayを受け取り、構築時に候補をsnapshotする。typed scalar arrayは受理する。
- `UnionOf`は`VSchema`だけを含む非空`Collection`を受け取る。objectを含む`Array()`はv1の公開契約にしない。Collectionのbranch順と要素集合は構築時にsnapshotするが、各child `VSchema` instanceはcloneせず参照を共有する。
- `ArrayOf(Nothing)`、`UnionOf(Nothing)`、未初期化・多次元・空のEnum、空のUnion、不正なEnum要素、重複するEnum候補、未初期化`VSchema`または`VSchema`以外のUnion要素はprogrammer misuseとして`Err.Raise`する。

### 2.2 Fluent API

`VSchema`は次のbuilderを公開する。

```vb
Public Function OptionalField() As VSchema
Public Function Nullable() As VSchema
Public Function Min(ByVal Boundary As Variant) As VSchema
Public Function Max(ByVal Boundary As Variant) As VSchema
Public Function Length(ByVal RequiredLength As Variant) As VSchema
Public Function WholeNumber() As VSchema
Public Function Pattern(ByVal Expression As String) As VSchema
Public Function Email() As VSchema
Public Function Field(ByVal FieldName As String, ByVal FieldSchema As VSchema) As VSchema
Public Function Strict() As VSchema
```

`OptionalField`はObject fieldが存在しない場合だけを許可する。`Empty`、`Null`、単体値には作用しない。`WholeNumber`は数学的に整数であるNumberを要求する。

各builderは同一instanceを変更して返す。後勝ちの上書きは行わず、同じconstraintの重複指定、`Min > Max`、負のLength、schema kindに適用できないmodifierはprogrammer misuseとする。

制約の検証順はschema nodeごとに固定する。型・Null/Emptyの判定後、`Length`、`Min`、`Max`、`WholeNumber`、`Pattern`、`Email`の順で評価し、最初に失敗した制約1件だけをそのnodeのIssueとする。Objectはfieldごとにこの規則を適用し、全体のIssue順はfield declaration orderを維持する。

`Length(n).Min(m)`は`m > n`、`Length(n).Max(m)`は`m < n`になる時点でprogrammer misuseとする。逆順のchainでも同じ条件をbuilder時に検出する。`Min`/`Max`と`WholeNumber`の組み合わせはNumberに限り許可する。`Pattern`と`Email`の併用はTextに限り許可し、上記の固定順で評価する。

適用可能な組み合わせは次のとおり。

| Modifier | Schema kind |
| --- | --- |
| `OptionalField` | すべて。Object fieldとして使用した場合だけ意味を持つ |
| `Nullable` | すべて |
| `Min`, `Max` | Text、Number、Array |
| `Length` | Text、Array |
| `WholeNumber` | Number |
| `Pattern`, `Email` | Text |
| `Field`, `Strict` | ObjectSchema |

### 2.3 Internal initialization (unsupported)

`VSchema`はVBAのclass初期化制約により、次のprocedureをPublicとして保持する。

```vb
Public Function InternalInitialize(ByVal KindCode As Long) As VSchema
```

これはv1の利用者向け安定APIではない。`Schema.bas`のfactoryだけが内部encoded kindを渡し、通常のschema kind値を直接渡す呼び出しは`vbObjectError + 2100`として拒否する。encoded kindの不正値と再初期化もprogrammer misuseとし、factory以外からの利用をサポート対象にしない。これはVBAの公開可視性による制約であり、security boundaryではない。

同じVBA classの別instanceのPrivate memberを参照できない制約を補うため、child schemaの合成、path付きnested validation、schema graph preflightに必要な`Internal*` hookもPublicとして保持する。これらもunsupported internal-only APIであり、利用者向け安定APIやsecurity boundaryではない。

### 2.4 Validation API

```vb
Public Function SafeParse(ByVal InputValue As Variant) As VValidationResult

Public Property Get Success() As Boolean
Public Property Get Value() As Variant
Public Property Get Issues() As Collection
Public Property Get ErrorText() As String
```

`VSchema`と`VValidationResult`を返すFunction、およびObjectを含む`Value`はVBAのObject assignment規則に従って`Set`で受け取る。scalarの`Value`は通常代入で受け取る。

`SafeParse`はvalidation failureではErrを投げない。成功時と失敗時のcontractは次のとおり。

| Property | Success | Failure |
| --- | --- | --- |
| `Success` | `True` | `False` |
| `Value` | 入力値。Objectは同じ参照 | `Empty` |
| `Issues` | 空のCollection | 1件以上のsnapshot Collection |
| `ErrorText` | 空文字列 | Issuesから決定的に生成した文字列 |

`Issues`は呼び出しごとにsnapshotを返す。呼び出し側が返却CollectionやIssue Dictionaryを変更しても、Result内部の状態と`ErrorText`は変化しない。

`ErrorText`は次の固定文法を使用する。

```text
<path>: <message> (expected=<expected>, received=<received>)
```

複数Issueは`vbCrLf`で連結し、先頭・末尾に余分な改行を置かない。`expected`と`received`はIssue snapshot内の文字列をそのまま使用し、raw input全体やsecretを追加出力しない。

## 3. Error boundary

### 3.1 Validation failure

入力値がschemaを満たさない場合は`Success=False`を返す。型不一致、required field欠落、constraint違反、strict objectのunknown fieldが該当する。

### 3.2 Programmer misuse

不正なschema構築は`vbObjectError + 2100`から`vbObjectError + 2199`の範囲で`Err.Raise`する。例:

- `ArrayOf(Nothing)`
- 空のfield名または`Field(name, Nothing)`
- 同じfield名の重複
- schema kindに適用できないmodifier
- 矛盾または重複するconstraint
- 空のEnumまたはUnion
- 循環するschema graph（`vbObjectError + 2103`）

初期割当は次のとおりとする。

```text
vbObjectError + 2100  invalid argument
vbObjectError + 2101  invalid constraint
vbObjectError + 2102  invalid schema composition
vbObjectError + 2103  cyclic schema
```

### 3.3 Runtime or environment failure

必要なruntime componentが利用できない場合や内部invariantが破られた場合は、`vbObjectError + 2200`から`vbObjectError + 2299`の範囲でErrを投げる。`SafeParse`はこれらをvalidation issueへ変換しない。

```text
vbObjectError + 2200  required runtime component unavailable
vbObjectError + 2299  internal invariant failure
```

`InternalInitialize`への通常の直接呼び出し、不正kind、factory encoded kindの破損は`vbObjectError + 2100`とする。初期化済みinstanceの再初期化は`vbObjectError + 2102`とする。

## 4. Issue contract

各Issueはlate-bound `Scripting.Dictionary`のsnapshotで、次のString keyを必ず持つ。

| Key | Type | Contract |
| --- | --- | --- |
| `path` | String | canonical path |
| `code` | String | 公開error code |
| `message` | String | 人間向けメッセージ |
| `expected` | String | 期待条件の安定した説明 |
| `received` | String | 入力値の安全な型・値要約 |

`received`は次の構造化文法を使用し、Object参照や配列そのものを公開しない。

```text
String("<escaped>")
Number(<invariant-number>)
Boolean(True|False)
Date(<yyyy-mm-ddThh:nn:ss>)
Null
Empty
Error(<number>)
Missing
Object(<TypeName>)
Array(<TypeName>)
```

String値はbackslash、double quote、control characterをescapeし、escape後の全体を80 UTF-16 code unit以内に収める。超過時は内容をmarker `...`込みで切り詰める。Numberはlocale非依存で小数点を`.`とし、桁区切りを使わない。Dateは秒未満を切り捨てて`yyyy-mm-ddThh:nn:ss`とする。これによりIssueが入力Objectの寿命、mutation、host localeへ依存することを防ぐ。

Issue順は決定的でなければならない。

1. depth-first
2. Object fieldはschemaへの追加順
3. Arrayは論理index順
4. Objectのnon-string keyはdescriptorのbinary ordinal順
5. strict unknown fieldはbinary ordinal順

v1の公開error codeは次のとおり。

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

## 5. Canonical path

- rootは常に`$`とする。
- identifier形式`[A-Za-z_][A-Za-z0-9_]*`のfieldは`$.user.name`形式とする。
- それ以外のfieldは`$["a.b"]`形式とする。backslash、double quote、改行、復帰、tabはそれぞれ`\\`、`\"`、`\n`、`\r`、`\t`へescapeし、その他のcontrol characterは`\uXXXX`とする。
- ArrayとCollectionは基底indexに関係なく、列挙順を0始まりの論理indexへ正規化する。
- v1で受理するnative arrayは一次元だけとする。
- 多次元配列は`invalid_array_rank` issueとする。
- 未初期化dynamic arrayは要素数0のarrayとして扱う。
- `ArrayOf`のlength constraint（`Length`、`Min`、`Max`）は要素検証より先に評価し、最初のconstraint issueだけをarray nodeへ追加する。
- native arrayのrank／bounds probeは限定helper内で行い、未初期化判定に伴う`LBound`／`UBound` Errをvalidation外へ漏らさない。入力array／Collectionは変更しない。

例:

```text
$
$.user.name
$.users[0]
$.users[2].email
$["field.with.dot"]
```

## 6. Object field semantics

- field名比較は`vbBinaryCompare`相当のcase-sensitive比較とする。
- 入力Dictionaryの`CompareMode`には依存せず、`Keys`列挙と`StrComp(..., vbBinaryCompare)`でfield照合を行う。
- 同じfield名の重複登録はprogrammer misuseとする。
- 入力Objectはv1では`Scripting.Dictionary`に限定する。
- Object判定は`TypeName(value) = "Dictionary"`を入口とし、`Count`、`Exists`、`Keys`の限定的な能力検査を行う。TypeName不一致または能力検査失敗は`invalid_type` issueとする。
- `Collection`や任意のclass instanceをkey-value Objectとして扱わない。
- String以外のinput keyは`.Strict()`の有無にかかわらず`invalid_key` issueとする。pathはobject nodeのpath、`expected`は`String key`、`received`は安全なdescriptorとし、複数keyはdescriptorのbinary ordinal順とする。
- unknown fieldのdefaultはpassthrough、`.Strict()`指定時は`unknown_field` issueとする。
- strict unknown fieldのpathはcanonical path、`received`は`DescribeValue`のString descriptorとし、binary ordinal順とする。
- required fieldが存在しない場合の`received`は`Missing`とする。これはpresent `Empty`、`Null`とは異なる。
- validationは入力Dictionaryを変更しない。

## 7. Type and comparison semantics

- `AnyValue`はNull、Empty、Error Variant、Nothingを含むすべてのVariant状態を受理する。
- 暗黙のString/Number/Boolean/Date変換を行わない。
- Numberとして受理する型はByte、Integer、Long、Single、Double、Currencyとする。
- Decimal VariantとLongLongは、32-bit/64-bitで同じ契約を検証できるまでv1では受理しない。
- Number constraintの境界値はVariantとして保持し、無条件にDoubleへ変換しない。
- Numberの`Min`/`Max`は受理対象のnumeric Variantだけを境界値として許可する。
- Text/Arrayの`Min`/`Max`と`Length`は、数学的に整数で0以上かつLong範囲内のnumeric Variantを許可する。内部保存時にLongへ正規化する。
- Null、Empty、Error Variant、Object、Array、Boolean、Date、StringをNumber boundaryへ渡した場合は`invalid argument`とする。
- 小数、Long範囲外、負数、非numeric値をText/Array boundaryまたは`Length`へ渡した場合は`invalid argument`とする。
- NaN、Infinity、overflowを発生させる比較は受理しない。
- Byte/Integer/Long同士は整数比較、SingleはDoubleへlossless widening、Currency同士はCurrency比較を優先する。異なるnumeric categoryを比較する場合は、比較前後のround-tripで値が変化しないと確認できる場合だけ共通表現へ変換する。overflowまたはprecision lossを証明できない場合は`vbObjectError + 2100`として拒否し、無条件の`CDbl`を行わない。
- NaN、Infinity、または比較結果を確定できないDouble/Currency境界は受理しない。
- `WholeNumber`はNumber値が数学的に整数の場合に成功する。
- Literal/EnumはString、Number、Boolean、Date、Null、Emptyのcategoryを区別する。Error Variant、Object、Arrayは定義候補として受理しない。
- Number category内では数値subtypeが異なっても値が等しければ一致する。
- String比較とfield名比較はbinary comparisonとする。
- `EnumOf`は構築時に候補値をsnapshotし、元の配列を後から変更してもschemaの意味を変更しない。候補重複は`vbObjectError + 2101`とする。
- StringのLengthはVBAの`Len`と同じUTF-16 code unit数とする。

`UnionOf`はbranch順に検証し、最初の成功で成功する。nested Unionはflattenせず構造を保持する。全branchがvalidation failureになった場合、内部branch Issueは破棄し、失敗pathに一件の`invalid_union` Issue（expected=`Union`）を追加する。branch内部で発生したprogrammer misuse、runtime/environment failure、internal invariant failureはUnionのvalidation failureへ変換せず、そのまま送出する。

`Pattern`は`VBScript.RegExp`をlate bindingし、`IgnoreCase=False`、`Global=False`、`MultiLine=False`に固定する。Expressionは入力全体に一致した場合だけ成功する。RegExp instanceはlazy compileし、不正なregular expressionは最初のvalidation時に`vbObjectError + 2100`としてErrを投げる。RegExp runtime自体が利用できない場合は`vbObjectError + 2200`のenvironment failureとする。

`Email`はPatternと同じRegExp runtimeを使う実用的なASCII簡易形式である。localとdomainを一つの`@`で区切り、全体長254 UTF-16 code unit以下、空白・Unicode・quoted local・comment・IP literalを受理対象外とする。RFC完全準拠を保証しない。`Pattern`と`Email`を併用した場合はPattern、Emailの順で評価する。

## 8. Schema graph and mutation

- builderはmutableであり、schemaを共有した後の変更はすべての参照先へ反映される。
- `Field`と`ArrayOf`はchild schemaをcloneしない。
- 自己参照または間接循環するschema graphはv1の対象外であり、`SafeParse`開始時のDFS preflightでvalidation開始前に`vbObjectError + 2103`として拒否する。active pathはCollection、schema identityはVBAの`Is`比較で判定し、shared childの非循環再利用は許可する。

## 9. Compatibility contract

| Environment | Status |
| --- | --- |
| Windows 64-bit Office 2016+ | 開発・検証対象 |
| Windows 32-bit Office 2016+ | compatibility-conscious、未検証、非保証 |
| macOS Office | compatibility-conscious、未検証、非保証 |

互換性を意識するため、Windows API、pointer-size依存宣言、Excel Object Modelへのコア依存を禁止する。`Scripting.Dictionary`と`VBScript.RegExp`はlate bindingするが、対象環境で利用可能であることを保証しない。必要なcomponentが存在しない場合はruntime/environment failureとする。

ドキュメントとrelease noteでは、未検証環境を「対応済み」「サポート済み」と表現してはならない。

## 10. Distribution contract

production sourceの正規配置は次の3ファイルとする。

```text
src/modules/Schema.bas
src/classes/VSchema.cls
src/classes/VValidationResult.cls
```

release artifactにはこの3ファイルだけを含める。`src/modules/Tests`、`src/modules/Xlflow`、`App.bas`、`Main.bas`、`Ui.bas`、workbook document moduleは含めない。

`src/`は開発・テスト用、`dist/VBA-Schema/`は3ファイルallowlistから生成するcommit対象外のimport payload、`.xlsm`はxlflow検証用成果物とする。現行のxlflow source tree全体をrelease artifactとして扱わない。release時は3ファイルの明示allowlistからstaging projectを作成する。

staged VBA sourceはUTF-8（BOMなし）かつLF改行でなければならない。README、LICENSE、CHANGELOG、sample workbook、ZIPなどの補助assetは3-file import payloadとは別に提供し、payloadのcomponent数を増やさない。

release gateは次を満たす必要がある。

1. 3ファイルだけを空のmacro-enabled workbookへimportできる
2. VBE compileが成功する
3. 外部参照設定を追加せずscalar smoke testが成功する
4. workbook document moduleを除くimport対象componentがこの3件と一致する

## 11. Compile contract

すべてのPublic APIを使用するcompile-only fixtureを保持する。Public名、引数型、戻り値型、fluent chainを変更する場合は、実装より先にfixtureを更新し、Windows 64-bit OfficeのVBE compileを通す。

2026-09-21時点で、次のAPI候補はWindows 64-bit OfficeのVBE compileを通過している。

```vb
Schema.ObjectSchema() _
    .Field("name", Schema.Text().Min(1).Max(100)) _
    .Field("age", Schema.Number().WholeNumber().Min(0).OptionalField()) _
    .Field("enabled", Schema.Bool().Nullable()) _
    .Field("createdAt", Schema.DateTime()) _
    .Field("tags", Schema.ArrayOf(Schema.Text()).Length(3)) _
    .Strict()
```
