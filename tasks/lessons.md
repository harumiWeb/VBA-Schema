# Lessons learned

- VBA class moduleのPrivate field／procedureは、同じclassの別instanceを通じて参照できない。nested schema処理では`Internal*` hookをunsupported internal-only APIとして明示し、VBE compile proofで確認する。
- `Scripting.Dictionary`の入力判定はTypeName後の限定capability checkに留め、入力側の能力不足をenvironment failureへ変換しない。library自身のIssue生成component不足だけをenvironment errorとして扱う。
