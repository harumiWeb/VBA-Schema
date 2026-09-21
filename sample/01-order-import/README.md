# 注文・明細インポート

Excel表やCSVを`Scripting.Dictionary`と`Collection`へ変換した後、注文データを取り込む前に検証するサンプルです。

## 実行

- import: `SampleOrderImport.bas`
- macro: `RunOrderImportSample`

成功例と失敗例を1件ずつ検証し、Immediate Windowへ`ErrorText`を出力します。

## 使用しているAPI

- `ObjectSchema`、`Field`、`Strict`
- nested `ObjectSchema`と`ArrayOf(Collection)`
- `Text().Pattern()`、`Text().Email()`、`DateTime()`
- `EnumOf`
- `Number().WholeNumber().Min().Max()`
- `OptionalField`

失敗例では、例えば次のpath/codeが確認できます。

```text
$.orderId       invalid_pattern
$.email         invalid_email
$.items[0].sku  invalid_pattern
$.items[0].quantity  too_small
$.debug         unknown_field
```

`Strict`を外すと未知フィールドをpassthroughできるため、移行時の挙動差も確認できます。
