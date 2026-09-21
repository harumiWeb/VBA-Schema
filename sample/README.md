# 利用者向けサンプル

このディレクトリには、VBA-Schemaを通常のVBAプロジェクトへ組み込むときの最小構成サンプルを置いています。

## 共通の実行手順

1. 対象VBAプロジェクトをバックアップする。
2. `Schema.bas`、`VSchema.cls`、`VValidationResult.cls`をVBEへimportする。
3. 次のいずれかのディレクトリにある`.bas`を標準Moduleとしてimportする。
4. VBEのImmediate Windowで、各READMEに記載された`Run...` macroを実行する。

サンプルはExcel Object Modelを使用しないため、ワークシートを準備せずにImmediate Windowで結果を確認できます。入力fixtureは、Excel表やJSONパーサーの結果を表す`Scripting.Dictionary`／`Collection`をVBA内で組み立てています。

## 一覧

| サンプル | 主な用途 | 主に見るAPI |
| --- | --- | --- |
| [01-order-import](01-order-import/README.md) | 注文・明細の取込前検証 | `ObjectSchema`、`ArrayOf`、`EnumOf`、`DateTime`、`Strict` |
| [02-settings-validation](02-settings-validation/README.md) | 設定シートやフォーム値の検証 | `EnumOf`、`Pattern`、`OptionalField`、`Nullable` |
| [03-api-payload](03-api-payload/README.md) | APIレスポンス相当の検証 | nested object、`UnionOf`、`Literal`、`ArrayOf` |

## 対応範囲

- サンプルはVBA-Schema本体の3ファイルimport payloadには含まれません。
- `Scripting.Dictionary`と`VBScript.RegExp`はlate bindingで使用します。利用できない環境では、入力エラーではなくenvironment errorになる場合があります。
- Windows 64-bit OfficeでVBE compileとruntime検証を行っています。macOS OfficeとWindows 32-bit Officeは対応を意識した実装ですが未検証・非保証です。
- JSONの取得、Excel表からDictionaryへの変換、設定値の読み取り自体はサンプルの責務外です。サンプルは変換後のVBA値をどのように検証するかに集中しています。
