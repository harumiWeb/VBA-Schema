# Lessons learned

- VBA class moduleのPrivate field／procedureは、同じclassの別instanceを通じて参照できない。nested schema処理では`Internal*` hookをunsupported internal-only APIとして明示し、VBE compile proofで確認する。
- `Scripting.Dictionary`の入力判定はTypeName後の限定capability checkに留め、入力側の能力不足をenvironment failureへ変換しない。library自身のIssue生成component不足だけをenvironment errorとして扱う。
- fresh managed sessionやfresh workbook probeでは`xlflow push --fast`を使わない。共有push-state cacheがsource unchangedと誤判定するため、session所有者が`push --session --no-save`で明示的にimportする。
- CIで固定CLIを使う場合は、`GITHUB_PATH`の解決順序に依存せず、専用`GOBIN`の絶対パスを`GITHUB_ENV`へ出して各stepから直接呼び出す。`version --json`で固定commitを検証し、Goモジュールキャッシュ対象（`go.sum`）がないリポジトリでは`setup-go`のcacheを無効化する。
