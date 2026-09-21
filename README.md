# VBA-Schema

VBAで宣言的なruntime validationを行う、軽量なschema libraryです。TypeScriptのZodに近いbuilder APIを、通常のVBA projectへ少ない導入作業で追加できます。

## 特徴

- 配布・import対象は `Schema.bas`、`VSchema.cls`、`VValidationResult.cls` の3ファイル
- 追加のVBA referenceやExcel Object Modelへの依存なし
- `SafeParse` が `Success`、`Value`、`Issues`、`ErrorText` を返す
- `ObjectSchema`、`ArrayOf`、`Literal`、`EnumOf`、`UnionOf`、`Pattern`、`Email` を提供
- builderはmutableです。child schemaや同じschema instanceを共有した後の変更は、共有先にも反映されます

## 対応状況

Windows 64-bit OfficeでVBE compileとruntime testsを検証しています。macOS OfficeとWindows 32-bit Officeは互換性を意識した実装ですが、手元に実機がないため未検証であり、対応済み・保証済みとは表記しません。

`ObjectSchema`はlate-bound `Scripting.Dictionary`、`Pattern`/`Email`はlate-bound `VBScript.RegExp`を使用します。これらのruntime componentが利用できない環境では、validation failureではなくenvironment error（`vbObjectError + 2200`）を送出します。

## インストール

1. GitHub releaseの3-file import payload、またはrepositoryの `dist/VBA-Schema` から次の3ファイルを取得します。
2. VBEで対象VBA projectを開き、`File > Import File...` から `.bas` と `.cls` を順にimportします。
3. 追加のreference設定は不要です。

更新時は同じ3ファイルを再importし、古い同名componentを置き換えます。削除時は `Schema` module、`VSchema` class、`VValidationResult` classをVBEから削除します。作業前にprojectのバックアップを作成してください。

## 最小例

```vb
Dim branches As Collection
Set branches = New Collection
branches.Add Schema.Text().Email()
branches.Add Schema.Number().WholeNumber()

Dim UserSchema As VSchema
Set UserSchema = Schema.ObjectSchema() _
    .Field("name", Schema.Text().Length(1).Max(100)) _
    .Field("contact", Schema.UnionOf(branches))

Dim inputObject As Object
Set inputObject = CreateObject("Scripting.Dictionary")
inputObject.Add "name", "Alice"
inputObject.Add "contact", "alice@example.com"

Dim result As VValidationResult
Set result = UserSchema.SafeParse(inputObject)
If result.Success Then
    Debug.Print "valid"
Else
    Debug.Print result.ErrorText
End If
```

`UnionOf`はbranchのCollectionを構築時にsnapshotしますが、各child `VSchema`はcloneせず参照共有します。全branchが失敗した場合は、失敗pathに一件の `invalid_union` Issueを返します。

## Public API

Factoryは `AnyValue`、`Text`、`Number`、`Bool`、`DateTime`、`ObjectSchema`、`ArrayOf`、`Literal`、`EnumOf`、`UnionOf` です。modifierは `OptionalField`、`Nullable`、`Min`、`Max`、`Length`、`WholeNumber`、`Pattern`、`Email`、`Field`、`Strict` です。

validation failureは `SafeParse` の `Success=False` と `Issues` で扱います。Issueは `path`、`code`、`message`、`expected`、`received` のDictionary snapshotです。pathは `$` から始まり、object fieldは `$.field`、array elementは `$[0]` の形式です。

主なerror codeは `required`、`invalid_type`、`too_small`、`too_big`、`invalid_length`、`invalid_integer`、`invalid_pattern`、`invalid_email`、`invalid_literal`、`invalid_enum`、`invalid_union`、`invalid_key`、`unknown_field`、`invalid_array_rank` です。

schema定義の誤り（例: `ArrayOf(Nothing)`、空のUnion、重複modifier）は `vbObjectError + 2100` から `+2199` のErrとして送出します。schema graphのcycleは `vbObjectError + 2103`、runtime component不足は `vbObjectError + 2200` です。これらをvalidation failureへ変換しないことで、入力不正と定義・環境の問題を区別できます。

## PatternとEmail

`Pattern`は `VBScript.RegExp` のcase-sensitive full matchです。`IgnoreCase=False`、`Global=False`、`MultiLine=False`で固定し、expressionは最初のvalidation時にlazy compileします。

`Email`は典型的な入力ミスを検出するASCII簡易形式です。全体254 UTF-16 code unit以下、一つの `@`、ASCII local/domainを要求します。Unicode、quoted local、comment、IP literalを含むRFCの全形式には対応せず、RFC完全準拠を保証しません。

## 開発と検証

Windowsでのsource checksは次で実行できます。

```powershell
rtk task verify
rtk task release-smoke
rtk task benchmark
```

VBEを使うbehavioral testsとcompile fixtureはxlflowのmanaged Excel sessionで実行します。release payloadは `rtk task release-stage` で生成し、`rtk task release-verify` で3ファイルallowlist、encoding、class headerを検証します。
`rtk task benchmark`はWindows 64-bit Excel専用の性能検証で、専用managed sessionを開始・破棄し、結果を `artifacts/benchmarks` に保存します。実行結果の`office_bitness`が`x64`でない場合は、x64 baselineとの比較・更新を行わずunsupported reportを残して失敗します。macOSとWindows 32-bit Excelの性能値は未検証です。

## License

MIT Licenseです。詳細は [LICENSE](LICENSE) を参照してください。
