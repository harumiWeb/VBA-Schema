# アプリケーション設定の検証

設定シートやフォームから読み取った値を、利用開始前に検証するサンプルです。

## 実行

- import: `SampleSettingsValidation.bas`
- macro: `RunSettingsValidationSample`

設定項目が不足している場合、`OptionalField`と`Nullable`で意味を分けられます。

- `OptionalField`: Dictionaryにキーが存在しない場合だけ許可
- `Nullable`: キーが存在し、値が`Null`の場合を許可

## 使用しているAPI

- `ObjectSchema`、`Field`、`Strict`
- `EnumOf`、`Pattern`
- `Number().WholeNumber().Min().Max()`
- `ArrayOf(Text())`
- `OptionalField`、`Nullable`

失敗例では次のような結果を確認できます。

```text
$.environment  invalid_enum
$.apiBaseUrl   invalid_pattern
$.port         too_small
$.tags[1]      invalid_type
$.debugMode    unknown_field
```
