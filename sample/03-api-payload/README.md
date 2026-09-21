# APIレスポンス相当の検証

JSONパーサーなどで取得したレスポンスを、VBAの`Dictionary`と`Collection`へ変換した後に検証するサンプルです。JSONの取得・解析機能そのものは含めません。

## 実行

- import: `SampleApiPayload.bas`
- macro: `RunApiPayloadSample`

## 使用しているAPI

- nested `ObjectSchema`
- `ArrayOf(Collection)`
- `UnionOf`（メールアドレスまたは数値ID）
- `Literal`、`Nullable`、`OptionalField`
- `Number().WholeNumber().Min()`、`Text().Length()`

失敗例では、レスポンスのどこが不正かをpath付きで確認できます。

```text
$.status       invalid_literal
$.contact      invalid_union
$.items[0].id  too_small
$.items[0].active invalid_type
$.traceId      unknown_field
```

`contact`を整数へ変更すると、同じUnionの2番目のbranchが成功する挙動も試せます。
