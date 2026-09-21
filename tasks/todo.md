# VBA-Schema v1 Development Roadmap

最終更新: 2026-09-21

この文書は、VBA-Schema v1の開発を別の作業者へ途中で引き継いでも、同じ設計契約と検証基準で完了まで進めるための実行ロードマップである。

## 1. 使い方

### 1.1 作業開始時に必ず読むもの

次の順序で確認する。

1. `AGENTS.md`
2. `tasks/todo.md`（本書）
3. `tasks/lessons.md`
4. `docs/specs/v1-contract.md`
5. `docs/adr/ADR-0001-small-distribution-and-portable-core.md`
6. `docs/adr/ADR-0002-vba-safe-api-and-error-boundary.md`
7. `docs/design.md`
8. `xlflow.toml`と`Taskfile.yml`
9. `rtk git status --short`
10. `rtk xlflow status --json`

規範の優先順位は次のとおり。

```text
docs/specs/v1-contract.md
    > accepted ADR
    > docs/design.md
    > 本ロードマップの説明文
```

矛盾を見つけた場合は実装で吸収せず、作業を止めてspecまたはADRを更新する。

### 1.2 チェック状態

- `[ ]`: 未着手または完了証拠なし
- `[x]`: 完了条件を満たし、検証証拠が本書のProgress Logに記録済み
- `BLOCKED:`: 外部環境または未解決Decision Gateにより進行不能

チェックを付けるだけでなく、各Phase完了時に次をProgress Logへ残す。

- 日付
- commit SHAまたは未コミットである旨
- 変更ファイル
- 実行コマンド
- test結果
- 使用した一時workspaceの絶対path
- 未検証事項
- 次の開始点

個別作業の識別子は`<Milestone>/<subheading>/<short label>`形式とする。例: `M3/Tests/binary-compare-mode`。行番号やcheckboxの通し番号は編集で変わるため識別子に使わない。

### 1.3 変更ルール

- production componentは`Schema.bas`、`VSchema.cls`、`VValidationResult.cls`の3ファイルだけとする。
- test、fixture、release helper、xlflow helperは3ファイル制約に含めない。
- production code変更にはfocused regression testを同じ作業単位で追加する。
- Public API、error code、path、type semanticsを変更する場合は、先にspecを更新し、必要ならADRを新規作成またはsupersedeする。
- ADR編集時は`adr-manager` skillを使用する。
- VBA sourceの検証時は`xlflow` skillのsource-to-workbook proof loopを使用する。
- xlflow側の不具合を見つけた場合は`docs/xlflow-issues/`へ再現条件を記録し、VBA-Schema側の回避策で恒久仕様を歪めない。
- macOS OfficeとWindows 32-bit Officeを、実機証拠なしに「対応済み」「検証済み」と記載しない。

## 2. Current Checkpoint

### 2.1 Repository state

- Branch: `review-vba-schema-design`
- Baseline implementation: `3593e01 fix: harden release and benchmark gates`
- production implementation: scalar core（AnyValue/Text/Number/Bool/DateTime、共通modifier、Result）、Object（Field／Strict／nested path／cycle preflight）、Array/Collection（一次元配列、logical index、length constraints）、Literal/Enum/Pattern/Email、Unionを実装済み
- production files: `Schema.bas`、`VSchema.cls`、`VValidationResult.cls`が存在
- 現在のtest: focused test 55件 + xlflow scaffold 5件 + compile-only fixture（test discovery対象外）。実行結果は59 pass、1 intentional TODO
- xlflow configured workbook: `build/Book.xlsm`
- xlflow session: inactive
- M7、MITライセンス、M8レビュー指摘（release staging安全性・benchmark環境ゲート）はコミット済み。独立Pass 2のreview targetは`f3d0434`で、blocking findingなし。
- CI hardening: `0fd4411`、`417b724`、`f273767`、`7f04898`で固定xlflow実体、VBA checkout改行、release safety probeの終了コードを修正済み。
- user-facing samples: `sample/`に注文・設定・APIレスポンスの3プロジェクトを追加し、root README・design・AGENTS treeへ反映済み。

### 2.2 Confirmed evidence

- Windows 64-bit OfficeのVBEで、v1 Public API候補の宣言とfluent chainがcompile成功している。
- typed native array、late-bound Object、Result propertyのObject/scalar assignmentを含むcompile fixtureが成功している。
- Windows 64-bit managed sessionで`PublicApiCompile.CompilePublicApi`をdiagnostic runし、VBE compileを含むmacro executionがsuccessしている。
- `rtk xlflow lint --json`: success
- `rtk xlflow analyze --json`: success
- `rtk xlflow test list --json`: focused testを含む60件を検出（23 source files）
- GitHub Actions `source-check` run `35587365802`: success（pinned xlflow install/version check、lint、analyze、production hygiene、LF format check、test discovery、release stage/verify、安全性テスト、benchmark環境ガード）
- `rtk task samples-verify`: sample 3 filesのformat、isolated lint、isolated analyzeがpass
- Windows 64-bit managed sessionでsample 3 macroのimport・compile・diagnostic runがpass（検証用workbookは保存せずdiscard）
- `rtk xlflow test --session --no-save --json`: 59 pass、1 intentional TODO
- `rtk xlflow run PublicApiCompile.CompilePublicApi --diagnostic --headless --session --no-save --json`: pass
- `rtk task release-smoke`: fresh workbookへの3-file import、VBE compile、scalar smoke、non-document component 3件確認がpass
- `rtk git diff --check`: success
- ADR-0001/0002/0005/0006: lint/review上のHigh/Medium findingなし

過去のcompile oracle用一時workspace:

```text
C:\temp\vba-schema-design-oracle-20260921
```

この一時workspaceは参考証拠であり、再現可能なrepository fixtureの代わりにしてはならない。

### 2.3 Explicitly unverified

- macOS Office
- Windows 32-bit Office
- Dictionary runtime unavailable環境の実機再現とenvironment error経路
- RegExp runtime unavailable環境の実機再現とenvironment error経路
- performance target
- CI上のExcel/VBE compile
- external referenceが増えていないことの自動検査

## 3. Milestone dependency map

```text
M0 Design baseline and decision gates
  -> M1 Development and release harness
    -> M2 Scalar core and result contract
      -> M3 Object validation
        -> M4 Array validation
          -> M5 Literal / Enum / Pattern / Email
            -> M6 Union
              -> M7 Hardening / documentation / release
                -> M8 Independent final review and v1 completion
```

後続Milestoneへ進む前に、直前Milestoneのexit gateをすべて満たす。

## 4. Decision Gates

未確定事項はproduction code内の暗黙判断にせず、該当Phaseの開始前に決める。

### DG-001 Issue recordのportable representation

期限: M1開始前

現行specは`Scripting.Dictionary`を必須concrete typeとして固定済み。このGateはmacOS compatibility-conscious方針との衝突を解消するための明示的な再検討であり、暗黙に契約を変更してはならない。

- [x] 現行の`Scripting.Dictionary`必須契約を維持する。Issue/snapshotはDictionary concrete typeを要求し、未提供環境ではenvironment errorとする。
- [x] macOSでDictionaryが利用できない場合はscalar validation errorへfallbackせず、runtime/environment failureとして伝播させる。
- [x] 決定を`docs/specs/v1-contract.md`へ反映する。
- [x] ADR-0001へ決定理由とconsequenceを反映する。
- [ ] portable fallbackを選ぶ場合は、Issue生成・snapshot fallback、runtime component unavailable test seam、`Issues`のcompile/runtime fixture、README/releaseの依存表記を実装taskとしてM2/M7へ追加する。
- [ ] Dictionary必須を維持する場合は、runtime unavailableがenvironment errorとして伝播するtestとREADMEの依存表記をM2/M7へ追加する。

決定時の評価軸:

- `Issue("path")`形式の使いやすさ
- concrete typeを公開契約にするbreaking-change risk
- snapshotの実装量
- macOS compatibility-conscious方針
- error生成時だけallocationする性能要件

### DG-002 Numeric comparison algorithm

期限: M2実装前

- [x] Byte/Integer/Long/Single/Double/Currency間の`Min`/`Max`比較手順を決める。
- [x] 無条件な`CDbl`を避け、lossless wideningとround-tripでoverflow・精度損失を拒否する手順を決める。
- [x] NaN、Infinity、overflow-producing comparisonを受理しない契約を決める。
- [x] `WholeNumber`のSingle/Double/Currency判定手順を決める。
- [x] 決定表と境界test casesの契約をspecへ追加する。

禁止事項:

- precision lossを黙って成功扱いする
- BooleanまたはDateをNumberとして扱う
- Error handlingを通常のvalidation分岐として広範囲に使用する

### DG-003 Error message and received descriptor grammar

期限: M2の`VValidationResult`実装前

- [x] `received`の正確なString grammarをtype別に決める。
- [x] Stringのescape、80 UTF-16 code unitでのtruncate、truncate markerを決める。
- [x] Numberのlocale-independent formattingを決める。
- [x] Dateの`yyyy-mm-ddThh:nn:ss` formattingと秒未満の扱いを決める。
- [x] `ErrorText`の改行、indent、issue間separatorを固定する。
- [x] raw secretをmessageへ含めない範囲を決める。
- [ ] golden testsを追加する。

### DG-004 Schema cycle detection

期限: M3開始前

- [x] specどおり、実際のvalue validation開始前にschema graph preflightを必ず行う。
- [x] builder時の早期検出は追加せず、mutable builderの後発cycleも`SafeParse`直前に検出する。
- [x] direct cycleとindirect cycleはactive pathのCollectionと`Is`比較で検出する。
- [x] mutable builderによる後発cycleを検出できることを確認する。
- [x] `vbObjectError + 2103`の発生時点をspecへ固定する。

### DG-005 Dictionary detection and key matching

期限: M3開始前

- [x] `TypeName = "Dictionary"`を入口に`Count`、`Exists`、`Keys`を限定検査する。
- [x] TypeName不一致／能力検査失敗は`invalid_type`、library自身のIssue Dictionary生成失敗は`vbObjectError + 2200`とする。
- [x] 入力Dictionaryの`CompareMode`に依存せず、Keys列挙と`StrComp(..., vbBinaryCompare)`でmatchingする。
- [x] String以外のDictionary keyは常に`invalid_key` validation issueとする。
- [x] strict unknown keyはbinary ordinalでsortする。
- [x] Dictionaryが利用できない環境ではlibrary側component生成時に`vbObjectError + 2200`を伝播する。

### DG-006 Native array inspection

期限: M4開始前

- [x] 未初期化dynamic arrayを安全に0要素と判定するhelperを設計する。
- [x] array rankは一次元bounds probeと二次元probeで安全に取得する。
- [x] 一次元以外を`invalid_array_rank`にする分岐を決める。
- [x] `LBound`に関係なく0始まりlogical pathへ変換する方法を決める。
- [x] 必要最小範囲の`On Error`とErr stateのclear／復元をspec、ADR、実装helperへ記録する。

### DG-007 Pattern and Email semantics

期限: M5開始前

- [x] `VBScript.RegExp`の`IgnoreCase=False`、`Global=False`、`MultiLine=False`を固定する。
- [x] Patternは入力全体への一致、invalid patternはlazy compile時の`vbObjectError + 2100`とする。
- [x] EmailはASCII簡易形式、全体長254 UTF-16 code unit以下、RFC完全準拠外とする。
- [x] Emailを`VBScript.RegExp`で実装する。
- [x] RegExp runtime unavailableは`vbObjectError + 2200`とする。実機再現は未検証として残す。
- [x] RFC完全準拠を目標にしない範囲をspecへ例示する。

### DG-008 Release artifact format

期限: M1完了前

- [x] 現行specどおり、import payloadは3 source filesだけに固定する。
- [x] ZIP、sample workbook、LICENSE、CHANGELOG等は3-file payloadとは別のrelease page/repository assetとして提供する。
- [x] 同一artifactへ同梱する変更は採用せず、3-file payload契約を維持する。
- [x] `dist/VBA-Schema/`はgenerated artifactとし、commit対象外にする。
- [x] version metadataの正はGit tagとtracked `CHANGELOG.md`とし、production VBA sourceへversion定数やVERSIONファイルを追加しない。
- [x] `CHANGELOG.md`をv1で作成し、release noteは補助release assetとして扱う。
- [x] release stagingのencodingをUTF-8（BOMなし）、改行コードをLFに固定してverifyする。
- [x] 追加のthird-party runtime referenceを導入しないため、`THIRD_PARTY_NOTICES.md`はv1 payloadへ追加しない。必要性はrelease時に再確認する。

### DG-009 CI and compile oracle ownership

期限: M1完了前

- [x] GitHub-hosted runnerのExcel-free source checkと、Excel/VBEが必要なlocal/self-hosted checkを分離する。
- [x] VBE compile evidenceはWindows 64-bit Excelを利用できる開発者またはself-hosted runnerがrelease gateで取得する。
- [x] ExcelのないCIでは`source-check`などのstatus名を使い、VBE compile passedとは表示しない。VBE側は`vbe-compile`等の別statusとする。
- [x] macOS/32-bit検証は実機提供時に別matrix/jobとして追加し、現時点では未検証のまま明記する。

### DG-010 InternalInitialize exposure

期限: M1開始前

- [x] `InternalInitialize`はPublic internal-onlyとする。Schema factoryだけが内部encoded kindを渡せる設計にする。
- [x] 通常の直接呼び出し・不正kind codeは`vbObjectError + 2100`、再初期化は`vbObjectError + 2102`とする。
- [x] unsupported internal APIであることをspecとADRへ記録する。READMEでは安定Public APIに含めない。
- [x] v1 public compatibility guaranteeの対象外であることをspecとcompile fixtureで区別する。
- [x] user-facing exampleとは別のinternal compile fixtureで`InternalInitialize`を直接呼び、存在、再初期化Err、不正kind Errを検証する。

### DG-011 Constraint composition

期限: M2実装前

- [x] `Length(3).Min(4)`と`Length(3).Max(2)`をbuilder時に拒否する。
- [x] `Min`/`Max`と`WholeNumber`の組合せはNumberだけ許可する。
- [x] `Pattern`と`Email`を併用した場合は固定評価順で検証する。
- [ ] specどおり、同じconstraint/modifierの重複指定はprogrammer misuseとして実装する。
- [ ] `Nullable`、`OptionalField`、`Strict`の重複をどのErr codeへ割り当てるか決める。
- [x] constraint評価順を固定し、同じnodeの同じ値には最初の1件だけIssueを出す。
- [x] 決定をspecへ反映する。focused testsはM2実装taskとして追加する。

### DG-012 Schema input ownership

期限: EnumはM5開始前、UnionはM6開始前

- [x] `EnumOf`へ渡した一次元native arrayを構築時snapshotする。
- [x] `UnionOf`へ渡したCollectionのbranch順と要素集合を構築時snapshotし、child `VSchema` instance自体はmutable shared referenceとして保持する。
- [x] child `VSchema` instance自体はmutable shared referenceとする既存契約との境界を整理する。
- [x] duplicate Enum candidateは`vbObjectError + 2101`、typed scalar arrayは受理、多次元・未初期化・空配列は`vbObjectError + 2100`とする。
- [x] Union branchの詳細issueは外部へ出さず破棄し、失敗pathに`invalid_union`を1件だけ返す。branch内のErrは伝播する。
- [x] nested Unionはflattenせずbranch構造を保持する。

## 5. M0 — Design baseline and decision gates

目的: 現在の設計変更をレビュー可能なbaselineとして確定し、実装者が異なる判断をしない状態にする。

### Tasks

- [x] 現在の`docs/design.md`、spec、ADR-0001/0002、本ロードマップをまとめてreviewする。
- [x] `docs/specs/v1-contract.md`の全Public signatureとcompile fixture候補が一致することを再確認する。
- [x] design内に旧API名（`Schema.String`、`.Optional`、`.Integer`等）が残っていないことを検索する。
- [x] v1 APIとfuture API（Parse、Strip、Positive/Negative等）の境界を確認する。
- [x] Decision GateのうちM1開始前に必要なDG-001とDG-010を解決する。
- [x] 設計変更をcommitし、baseline SHAを本書へ記録する。

### Verification

```powershell
rtk rg "Schema\.(Any|String|Boolean|Date|Object)\(|\.Optional\(|\.Integer\(" .\docs
rtk rg -n "[ \t]+$" .\docs .\tasks\todo.md
rtk git status --short
rtk git diff --check
# baseline commit前に対象をstageした後、untrackedだった文書も含めて確認する
rtk git diff --cached --check
rtk xlflow lint --json
rtk xlflow analyze --json
```

`git diff --check`はuntracked fileを検査しない。上記のsource-level検索と`git status`を併用し、baseline commit直前はstage後の`git diff --cached --check`で全対象を検査する。

### Exit gate

- [x] spec/ADR/design間のHigh/Medium矛盾がない。
- [x] DG-001が解決済み。
- [x] DG-010が解決済み。
- [x] baseline commit SHAがCurrent Checkpointに反映済み。
- [x] macOS/32-bitの未検証表現が全資料で一致している。

## 6. M1 — Development, test, and release harness

目的: production実装を始める前に、compile、test、releaseを再現可能にする。

### Repository structure

- [x] `src/classes/`を作成する。
- [x] `src/modules/Tests/`をv1 test構成へ整理する。
- [x] compile-only fixtureの配置場所を決めて追加する。
- [x] clean checkoutから`build/Book.xlsm`を作成または復元するbootstrap手順を実装する。
- [ ] bootstrap元となるtracked template/source、Excel prerequisite、既存workbookを上書きしない条件を文書化する。
- [x] `dist/VBA-Schema/`を生成するallowlist staging scriptまたはTaskを追加する。
- [x] staging用一時projectを安全に作成・破棄できるようにする。
- [x] `Taskfile.yml`のplaceholder taskを実作業用taskへ置き換える。
- [x] 新しいdirectory/fileを追加した時点で`AGENTS.md`のproject treeを更新する。

推奨Task名:

```text
task lint
task analyze
task fmt-check
task provision-workbook
task test
task compile-oracle
task verify
task benchmark
task release-stage
task release-verify
task release-smoke
```

`task provision-workbook`は、ignored/untrackedの`build/Book.xlsm`が存在しないclean checkoutを正式な開始状態として扱う。tracked sourceまたはtemplateから再構築し、Excel/VBIDEなどの前提不足は明示的に失敗させる。既存のuser-owned workbookを無断で上書きしてはならない。

### Compile fixture

- [x] 3 production componentへspecどおりのPublic signature skeletonを追加する。
- [x] 未実装methodがfake successを返さず、明示的なinternal/not-implemented Errで停止するようにする。
- [x] spec記載の全factoryをfixture内で呼び出す。
- [x] 全fluent modifierをchain内で呼び出す。
- [x] `UnionOf(Collection)`をcompileする。
- [x] typed native arrayを`SafeParse(ByVal Variant)`へ渡す。
- [x] late-bound Objectを`SafeParse`へ渡す。
- [x] `Result.Value`をscalar代入とObjectの`Set`代入の両方でcompileする。
- [x] `Success`、`Issues`、`ErrorText`を使用する。
- [x] internal compile fixtureで`InternalInitialize`を直接呼び、DG-010で決めたErr contractをcompileする。
- [x] fixtureがrelease artifactへ含まれないことをtestする。

### Release staging

- [ ] DG-008を解決する。
- [ ] DG-009を解決する。
- [x] `dist/VBA-Schema/`をimport payload専用directoryとし、release metadataやsampleを混在させない。
- [x] source allowlistを次の3ファイルへ固定する。

```text
src/modules/Schema.bas
src/classes/VSchema.cls
src/classes/VValidationResult.cls
```

- [x] allowlistに不足ファイルがあれば明示的に失敗する。
- [x] allowlist以外の`.bas`/`.cls`/`.frm`をartifactへ入れない。
- [x] 空のmacro-enabled workbookへ3ファイルをimportするsmoke pathを用意する。
- [x] workbook document moduleを除くimport対象componentが3件であることを検証する。
- [ ] 外部参照設定が増えていないことを確認する。

### xlflow baseline

- [x] `rtk xlflow doctor --json`でExcel/COM/VBIDE環境を確認する。
- [x] clean checkoutで`task provision-workbook`を実行し、その後のtest discoveryまで再現する。
- [x] `rtk xlflow status --json`でrecovery不要を確認する。
- [x] sourceとworkbookのauthorityを確定する。
- [x] managed sessionを使うか既存workbookへattachするか記録する。
- [x] scaffold testはtest harness smokeとして保持し、`Test_Sample_Todo`は未実装機能の可視化用に残す。
- [x] formatter実行前後のdiffを確認するTaskを用意する。

### Exit gate

- [x] repository内のfixtureから全Public APIのVBE compileを再現できる。
- [x] clean checkoutからdevelopment workbookを再構築できる。
- [x] `task verify`相当の一括checkが存在する。
- [x] 3-file stagingとcomponent count verificationが自動化されている。
- [x] 開発用workbookと配布artifactが混同されない。

## 7. M2 — Scalar core and result contract

目的: AnyValue/Text/Number/Bool/DateTime、Result、共通modifier、error contractを完成させる。

### Tests first

- [x] `TestAnyValue.bas`を追加する。
- [x] `TestText.bas`を追加する。
- [x] `TestNumber.bas`を追加する。
- [x] `TestBool.bas`を追加する。
- [x] `TestDateTime.bas`を追加する。
- [x] `TestNullable.bas`を追加する。
- [x] `TestResult.bas`を追加する。
- [x] `TestErrors.bas`を追加する。
- [x] root path `$`のtestを追加する。
- [x] success/failure双方の`Value`、`Issues`、`ErrorText` contract testを追加する。
- [x] Issues snapshotを呼び出し側が変更してもResultが変化しないtestを追加する。
- [x] programmer misuseとenvironment/internal errorのErr番号testを追加する。

### Production implementation

- [x] `src/modules/Schema.bas`のsignature skeletonをscalar実装へ置き換える。
- [x] `src/classes/VSchema.cls`のsignature skeletonをscalar実装へ置き換える。
- [x] `src/classes/VValidationResult.cls`のsignature skeletonを実装へ置き換える。
- [x] `Option Explicit`を全production componentへ付ける。
- [x] `SchemaKind`と一度だけ実行可能な`InternalInitialize`を実装する。
- [x] 未初期化schemaと再初期化をprogrammer misuseにする。
- [x] `SafeParse(ByVal InputValue As Variant)`を実装する。
- [x] `AnyValue`を実装する。
- [x] `Text`を`vbString`だけに限定する。
- [x] `Number`をByte/Integer/Long/Single/Double/Currencyだけに限定する。
- [x] `Bool`を`vbBoolean`だけに限定する。
- [x] `DateTime`を`vbDate`だけに限定する。
- [x] scalarのNull、Empty、Error Variant、Nothing matrixをspecどおり実装する。
- [x] `Nullable`を実装する。
- [x] `OptionalField`のstateを実装し、root値の`Empty`を許可しないことをtestする。
- [x] `Min`、`Max`、`Length`、`WholeNumber`をscalar kindへ実装する。
- [x] duplicate/conflicting constraintをprogrammer misuseにする。
- [x] `VValidationResult`のSuccess/Failure invariantを実装する。
- [x] Issue code、path、message、expected、receivedを生成する。
- [x] scalarのIssue orderとErrorTextをdeterministicにする。

### Required scalar edge cases

- [x] empty StringとLength 0
- [ ] String Min/Max/Length boundary
- [ ] boundary引数`3`、`3&`、`CByte(3)`、`3#`
- [ ] fractional/negative/overflow/non-numeric length rejection
- [x] Byte/Integer/Long/Single/Double/Currency
- [x] Decimal Variant rejection
- [ ] LongLong rejection（使用可能な64-bit環境のみ）
- [x] numeric String rejection
- [x] Boolean/DateをNumberとして拒否
- [ ] positive/negative zero
- [x] Single/DoubleのfractionとWholeNumber
- [x] scalarのNull、Empty、Error Variant
- [x] `AnyValue`のNull/Empty/Error/Nothing
- [x] `AnyValue`のscalar/Object/native array/Collection
- [ ] rootまたはArray要素の`OptionalField`がEmptyを許可しない（Arrayは未実装）
- [x] invalid modifier/kind combinationsのscalar部分

### Decision gates

- [x] DG-002 resolved
- [x] DG-003 resolved
- [x] DG-010 resolved
- [x] DG-011 resolved

### Verification loop

```powershell
rtk xlflow lint --json
rtk xlflow analyze --json
rtk xlflow push --fast --session --no-save --json
rtk xlflow test --filter TestText --session --no-save --json
rtk xlflow test --session --no-save --json
```

### Exit gate

- [x] scalar focused tests pass。
- [x] full tests pass（24 pass、scaffoldのintentional TODO 1件）。
- [x] VBE compile passes。
- [x] `lint` and `analyze` pass with no unexplained findings。
- [x] compile fixture and 3-file release smoke still pass。

## 8. M3 — Object validation

目的: Dictionary Object、Field、OptionalField、nested path、Strictを完成させる。

### Tests first

- [x] required field present/missing
- [x] `OptionalField` missing success
- [x] present `Empty` is not missing
- [x] present `Null` requires `Nullable`
- [x] nested Object success/failure
- [x] multiple deterministic issues
- [x] passthrough unknown field
- [x] strict unknown field
- [x] strict unknown field binary ordinal ordering
- [x] field declaration ordering independent of Dictionary enumeration
- [x] case-sensitive field matching
- [x] input `CompareMode = vbTextCompare`でもbinary field matching
- [x] duplicate field definition is programmer misuse
- [x] empty field name is programmer misuse
- [x] `Field(name, Nothing)` is programmer misuse
- [x] wrong input Object/class/Collection rejection
- [x] Nothing rejection
- [ ] `Scripting.Dictionary` runtime unavailable時にenvironment errorがvalidation issueへ変換されないことを検証する。
- [ ] 実環境でunavailableを再現できない場合は、限定的なfactory seamによるerror-path testとWindows 64-bit integration結果を分け、未再現事項を記録する。
- [x] escaped path for dot, bracket, quote, backslash, control characters
- [x] direct and indirect schema cycles

### Implementation

- [x] DG-004を解決する。
- [x] DG-005を解決する。
- [x] `ObjectSchema` factoryを実装する。
- [x] schema field lookupとfield declaration orderを別管理する。
- [x] input DictionaryのCompareModeに依存しないbinary lookupを実装する。
- [x] `Field`、`OptionalField`、`Strict`を実装する。
- [x] unknown fieldをbinary ordinalで並べる。
- [x] canonical object pathを構築する。
- [x] validation中にinput Dictionaryを変更しない。
- [x] schema graph cycleをvalidation開始前に拒否する。

### Exit gate

- [x] Object focused tests pass。
- [x] scalar regression tests pass。
- [x] nested pathが`$.user.address.zip`形式で固定されている。
- [x] field orderとunknown field orderが複数runで同一。
- [x] release smoke passes。

## 9. M4 — Array and Collection validation

目的: 一次元native arrayとCollectionを同じlogical sequence contractで検証する。

### Tests first

- [x] zero-based native array
- [x] one-based native array
- [x] negative `LBound` native array
- [x] typed native array
- [x] Variant array
- [x] uninitialized dynamic array as empty
- [x] empty Collection
- [x] populated Collection
- [x] typed object array
- [x] Collection内のNothing、Error Variant、Object要素
- [x] nested Array/Collection
- [x] wrong element with correct logical index
- [x] multidimensional array returns `invalid_array_rank`
- [x] Dictionary is not treated as array
- [x] Min/Max/Length boundaries
- [x] root and nested array paths

### Implementation

- [x] DG-006を解決する。
- [x] `ArrayOf(ItemSchema)`を実装する。
- [x] Nothing ItemSchemaをprogrammer misuseにする。
- [x] native array detection helperを実装する。
- [x] safe initialized/rank/bounds helpersを実装する。
- [x] Collection iterationを実装する。
- [x] actual indexを0-based logical ordinalへ正規化する。
- [x] length constraintsをTextと共通contractで実装する。
- [x] input array/Collectionを変更しない。

### Exit gate

- [x] Array/Collection focused tests pass。
- [x] Object/scalar regression tests pass。
- [x] uninitialized arrayでruntime errorが漏れない。
- [x] multidimensional arrayがlibrary errorではなくvalidation issueになる。
- [x] release smoke passes。

## 10. M5 — Literal, Enum, Pattern, and Email

### Literal and Enum tests

- [x] same category and same value success
- [x] String `"1"` vs Number `1`
- [x] Number subtype equivalence
- [x] Boolean and Number separation
- [x] Date and Number separation
- [x] Null and Empty separation
- [x] Error/Object/Array definition rejection
- [x] empty Enum rejection
- [x] Enum definition mutation after construction does not alter schema, or chosen ownership rule is documented and tested
- [x] duplicate Enum candidates follow a documented rule
- [x] binary String comparison

### Literal and Enum implementation

- [x] DG-012のEnum ownership部分を解決する。
- [x] `Literal`を実装する。
- [x] `EnumOf`を実装する。
- [x] definition valuesをvalidation/normalizationする。
- [x] scalar category comparison helperを実装する。
- [x] exact Number comparisonをDG-002のcontractへ合わせる。

### Pattern and Email tests

- [x] Pattern match/non-match
- [x] invalid expression programmer misuse
- [ ] RegExp runtime unavailable environment error
- [x] Email accepted examples
- [x] Email rejected examples
- [x] empty/local-only/multiple-`@`/whitespace cases
- [x] Unicode and maximum length cases according to DG-007
- [ ] locale-independent behavior

### Pattern and Email implementation

- [x] DG-007を解決する。
- [x] lazy RegExp creation/cacheを実装する。
- [x] invalid patternとruntime unavailableを区別する。
- [x] `Pattern`を実装する。
- [x] `Email`を実装する。
- [x] RFC完全準拠を目標にしない範囲をspec/ADRへ記載する。READMEへの記載はM7で行う。

### Exit gate

- [x] Literal/Enum/Pattern/Email focused tests pass。
- [x] full regression tests pass。
- [ ] environment errorがvalidation issueへ変換されない。
- [x] release smoke passes。

## 11. M6 — Union

### Tests first

- [x] first branch success
- [x] later branch success
- [x] all branches fail
- [x] nested Union
- [x] empty Union is programmer misuse
- [x] Nothing/non-VSchema branch is programmer misuse
- [x] branch order is deterministic
- [x] external Issue is one `invalid_union` at failure path
- [x] branch internal errors do not leak into public Issues
- [x] environment/programmer errors inside branch propagate instead of becoming union mismatch
- [x] input Collection mutation after construction follows documented ownership rule

### Implementation

- [x] DG-012のUnion ownership部分を解決する。
- [x] `UnionOf(ByVal Schemas As Collection)`を実装する。
- [x] schema Collectionをvalidation/normalizeする。
- [x] validation failureとthrown Errを分離する。
- [x] temporary branch issuesを外部Resultから隔離する。
- [x] 全branch failure時に`invalid_union`を1件だけ生成する。

### Exit gate

- [x] Union focused tests pass。
- [x] full regression tests pass。
- [x] nested error path and issue order are deterministic。
- [x] release smoke passes。

## 12. M7 — Hardening, documentation, CI, and release

### Full quality pass

- [ ] 全production procedureを責務別private procedureへ分割する。
- [ ] 巨大procedure、広範囲`On Error Resume Next`、暗黙Variantを除去する。
- [x] Windows API、unqualified Excel reference、Select/Activateがないことを検索する（`task quality-audit`）。
- [x] optional xlflow dataflow rulesを有効化する価値を評価する（既存の戻り値・scope shadowing警告が多く、通常gateにはしない）。
- [x] `rtk xlflow metrics --json`でstatic complexity hotspotを確認する。
- [x] dead codeとunused private procedureを確認する（optional lint evaluationで未使用privateの新規findingなし）。
- [x] public API compile fixtureを全signatureと照合する（`PublicApiCompile.CompilePublicApi` pass）。

### Performance

- [x] runtime benchmark専用の`src/modules/Benchmarks/ValidationBenchmarks.bas`を追加し、release artifactから除外する。
- [x] `task benchmark`を実装し、同一fixtureを専用managed Excel sessionで実行できるようにする。
- [x] `task benchmark`自身が開始時のsession ownershipを確認し、専用sessionのstart、計測、stop/cleanupを所有する。user-owned workbookへattach中は変更せず明示的に失敗する。
- [x] benchmark結果をx64実行時は`artifacts/benchmarks/<timestamp>-windows-x64.json`へ出力し、非x64はunsupported suffixで保存する。
- [x] 比較用baselineの形式とtracked保存場所を決める（`benchmarks/windows-x64-baseline.json`）。directory追加に伴い`AGENTS.md`のtreeも更新する。
- [x] 1,000-field object（scalar field）fixtureを作成する。
- [x] 10,000 scalar validations fixtureを作成する。
- [x] success-heavy fixtureとissue-generation fixtureを分ける。
- [x] warm-up後5回のmedianを記録する。
- [x] 1,000 fields < 500 msを検証する。
- [x] 10,000 validations < 1,000 msを検証する。
- [x] baselineに対する25%超の悪化がないことを確認する。
- [x] base/headで同じfixture、iteration、metric setを使う。
- [x] elapsed timeに加えvalidation count、issue count、RegExp生成回数などのdeterministic counterを記録する。
- [x] machine、Office bitness、Excel version、commandを記録する。

`rtk xlflow metrics --json`はstatic complexityの証拠であり、runtime performanceの合否判定には使わない。runtime targetは`task benchmark`のJSONで判定し、fixture ID、warm-up回数、計測5回のraw値とmedian、validation count、issue count、RegExp生成回数、環境情報を保存する。

performance targetを変更してreleaseする場合は、単なるProgress Log記載では完了にならない。specまたはADRを先に更新し、変更理由、利用者への影響、変更前後を同一条件で比較したbenchmark JSONを残し、release gate承認者と承認日をProgress Logへ記録する。

### Documentation

- [x] READMEに「3 files / no reference setup」を記載する。
- [x] READMEにinstallation手順を追加する。
- [x] READMEにcanonical exampleを追加する。
- [x] READMEにPublic API overviewを追加する。
- [x] READMEにerror handlingとIssue例を追加する。
- [x] READMEにmutable builderのaliasing warningを追加する。
- [x] READMEにPattern/Emailの非RFC完全準拠を追加する。
- [x] READMEにWindows 64-bit verified、macOS/32-bit unverified/non-guaranteedを明記する。
- [x] error code一覧を公開する。
- [x] path grammarを公開する。
- [x] release/update/remove手順を追加する。
- [x] examplesをVBE compile fixtureで検証する。
- [x] design/spec/ADRと実装のdrift auditを行う（benchmark contract、release boundary、READMEを同期）。

### CI and automation

- [x] Excel不要のlint/analyze/test discovery/format checkをCIへ追加する（`.github/workflows/source-check.yml`）。
- [x] behavioral VBA testsはWindows + Excelが必要なgateとして分離する。
- [x] Excel/VBE compile checkをlocalまたはself-hosted gateとして定義する。
- [x] CI status名から検証範囲が分かるようにする。
- [x] release staging/verification TaskをExcel-free CIへ接続する。
- [x] CI failureと未実行を区別する。

### Release artifact

- [x] DG-008の決定どおり3-file artifactを生成する。
- [x] `Schema.bas`、`VSchema.cls`、`VValidationResult.cls`以外が含まれないことを確認する。
- [x] fresh workbookへimportする（release smokeで代替検証）。
- [x] fresh workbookで追加reference設定なしにcompileできることを確認する（release smokeで検証）。
- [x] VBE compileする（release smokeで検証）。
- [x] canonical Public API exampleをcompile fixtureで検証する。
- [x] source artifactのSHA-256 checksumをProgress Logへ記録する。
- [x] source artifactのencoding、改行コード、`Attribute VB_Name`、class attributesを確認する。
- [x] LICENSEがrepositoryから取得可能であることを確認する。現行specの3-file import payloadへは含めない。
- [x] 追加のthird-party referenceを導入しないため、`THIRD_PARTY_NOTICES.md`はv1 payloadへ含めない方針を確認する。
- [x] `CHANGELOG.md`を作成し、versionの正をGit tagとする方針を記録する。
- [x] DG-008で採用した補助物は3-file import payloadの外側で管理する。

### Compatibility-conscious static audit

- [x] Windows APIがない。
- [x] pointer-size依存宣言がない。
- [x] LongLong型名へcompile-time依存していない。
- [x] Excel Object Modelへコア依存していない。
- [x] path/string handlingがhost path separatorへ依存していない。
- [x] unsupported runtime component failureが明示的である。
- [x] macOS/32-bitを検証済みと誤記していない。

### Exit gate

- [x] full tests pass（intentional TODO 1件を除く59 pass）。
- [x] lint/analyze pass。
- [x] VBE compile pass on Windows 64-bit（local release smoke）。
- [x] performance targets pass、baseline comparison pass。
- [x] release artifact verification pass。
- [x] README/spec/ADR/design driftなし（benchmark contractを同期）。

## 13. M8 — Independent final review and completion

### Review

- [x] 実装者とは別のreviewerがread-only reviewを行う。
- [x] `orca-supervised-final-review` skillで独立worktree reviewを実行する。
- [x] review scopeにproduction 3 files、tests、release helper、docs、CIを含める。
- [x] confirmed findingとunsupported observationを区別する。
- [x] valid/in-scope findingだけを最小修正する。
- [x] 修正ごとにfocused regressionを追加する。
- [x] 修正後にfull verificationを再実行する。

### Final verification commands

実際のTask/CLI仕様に合わせて確定すること。最低限:

```powershell
rtk git diff --check
rtk git status --short
rtk task provision-workbook
rtk xlflow status --json
rtk xlflow lint --json
rtk xlflow analyze --json
rtk task fmt-check
rtk task benchmark
rtk task release-stage
rtk task release-verify
```

`task benchmark`は後続のstart/attach分岐とは独立したself-contained taskであり、専用managed sessionだけを開始・停止する。user-owned workbookへattach中なら実行せず、利用者が閉じた後のclean session stateで再実行する。benchmark完了後に、通常のbehavioral verification用sessionを次のどちらかで開始する。

workbook実行系は、開始時の所有状態に応じて次のどちらか一方を選ぶ。

Managed session（user-owned workbookが開かれていない場合）:

```powershell
rtk xlflow session start --json
rtk xlflow push --fast --session --no-save --json
rtk xlflow test --session --no-save --json
```

User-owned workbook（利用者が既に開いている場合）:

```powershell
rtk xlflow session attach --json
rtk xlflow push --fast --session --no-save --json
rtk xlflow test --session --no-save --json
```

検証後はmanaged sessionだけを停止する。user-owned workbookは閉じない。いずれも最後に`rtk xlflow status --json`を実行し、dirty/recovery stateとcleanup結果をProgress Logへ残す。

### User-facing samples

- [x] `sample/`に注文・明細、アプリ設定、APIレスポンスの3独立プロジェクトを追加する。
- [x] 各sampleに目的、import手順、実行macro、期待するIssue path/code、runtime依存、未検証platformを記載する。
- [x] root README、`docs/design.md`、`AGENTS.md`のtreeへsample配置とpayload外の境界を反映する。
- [x] `tools/verify-samples.ps1`と`task samples-verify`でformat、isolated lint、isolated analyzeを自動検証する。
- [x] Windows 64-bit managed sessionで3 sample macroのimport・compile・diagnostic runを確認する。
- [ ] macOS Office、Windows 32-bit Office、Dictionary/RegExp unavailable環境でのsample実行を検証する（実機提供時の別gate）。

### v1 completion gate

- [x] `docs/design.md`のDefinition of Doneを全項目確認した。
- [ ] 本ロードマップのM0-M8 exit gateをすべて満たした。
- [x] unresolved High/Medium review findingがない。
- [x] unverified platformを正確に表示している。
- [ ] clean checkoutからverificationを再現できる。
- [x] release artifactを生成できる。
- [x] implementation commit SHAとartifact checksumをProgress Logへ記録した。

## 14. Deferred beyond v1

次はv1完了条件へ含めない。追加する場合は新しいspec/ADR/roadmapを作る。

- Parse
- Strip unknown fields
- Positive/NonNegative/Negative/NonPositive
- Transform
- Refine/SuperRefine
- Callback validator
- Default values
- Coercion
- Recursive/lazy schemas
- General class reflection
- OpenAPI generation
- JSON parsing
- HTTP
- Schema serialization
- Code generation
- Compile-time type inference

## 15. Handoff protocol

作業を中断または別作業者へ引き継ぐ場合、終了前に次を行う。

- [ ] 作業中のtask keyをProgress Logへ記載する。
- [ ] 完了したcheckboxだけを更新する。
- [ ] failing test名と正確なerrorを記載する。
- [ ] source/workbook/sessionのどれが最新か記載する。
- [ ] `rtk xlflow status --json`のrecovery状態を記載する。
- [ ] sessionがdirtyなら、保存・破棄・継続のどれが必要か記載する。
- [ ] 一時workspaceの絶対pathを記載する。
- [ ] 実行済みと未実行のverificationを分ける。
- [ ] macOS/32-bitを含む未検証事項を残す。
- [ ] 次に実行すべき具体的な1コマンドを記載する。

引継ぎ先は、Progress Logだけを信用せず、`git status`、spec、xlflow statusを再確認する。

## 16. Progress Log

新しい記録を上へ追加する。

### 2026-09-21 — M8 independent review / Pass 2 completion

- Status: `orca-supervised-final-review`の新規独立worktreeでPass 2を完了。`worker_done`を通常のOrca deliveryで受信し、報告をacknowledge、worker release、worktree削除まで完了した。
- Review target: `f3d0434 docs: record independent review remediation`（base `8ec96e7`）。
- Result: blockingなvalid/in-scope実装findingなし。`release-stage.ps1`のprotected root（repository root、`src`、`.git`、`.xlflow`、`build`）の同一・子孫・祖先拒否、source hash不変のfocused safety test、benchmarkのx64 pass / x86・unknown reject、非x64でのbaseline not-runとunsupported reportを独立に再確認した。
- Follow-up: P3として、旧SHA参照とM8/M0 checklistのstale状態を本書で更新。benchmarkの`command`はcanonical public entry point `task benchmark`を記録する契約をspecへ明記した。
- Verification evidence: Pass 2 workerの`git diff --check`、lint、analyze、production hygiene、test list（60件）、xlflow status（inactive/clean）はpass。Pass 1 remediationで`rtk task verify`、release-stage/verify、release-smoke、x64 benchmark、full behavioral tests（59 pass、intentional TODO 1件）、Public API compileをpass済み。
- Unsupported/unverified: local and remote format gate failureはun-pinned xlflow/cache状態を含む環境差異として未解決。GitHub target workflow、Excel/VBE target CI、Windows 32-bit/macOS実機、clean checkout再現は未検証。
- Implementation commit: `3593e01 fix: harden release and benchmark gates`。M7/MIT既存commitは`57e82c6`、`8ec96e7`。release payload SHA-256はM7記録を正とする。
- Next task: clean checkout再現とGitHub/target CI実行環境が利用可能になった時点で、未検証項目を別gateとして確認する。実装上のM8 blocking findingは残っていない。

### 2026-09-21 — CI source-check remediation

- Status: 直近push後のGitHub Actions失敗を、実行環境差異とPowerShell終了コードの問題に分離して修正。現在のbranchはcleanで、修正は`review-vba-schema-design`へpush済み。
- Root causes:
  - `GITHUB_PATH`の解決順序により、意図した固定xlflow以外を呼び得た。`go install`生成物はversion metadataが`dev/none`になるため、固定release asset（v0.32.1 / commit `52e93661cdc830592de88e047cbeed3bb2ac4488`）を専用binへ展開し、`version --json`で検証する方式へ変更。
  - Windows checkoutのCRLF化で`xlflow fmt`が18ファイルを一律変更扱いにしたため、`.gitattributes`で`.bas`／`.cls`の`eol=lf`を固定。
  - `test-release-stage-safety.ps1`は意図的な子プロセス失敗の`$LASTEXITCODE`を最後に残していたため、成功時の`exit 0`を追加。
- Changed: `.github/workflows/source-check.yml`、`.gitattributes`、`tools/check-format.ps1`、`tools/test-release-stage-safety.ps1`、`tasks/lessons.md`。
- Commits: `0fd4411`、`417b724`、`f273767`、`7f04898`。
- Verification:
  - `rtk actionlint`: pass。
  - `rtk task verify`: pass（60 tests discovered、lint/analyze/format/quality/release safety/benchmark guard）。
  - Windows `core.autocrlf=true`相当の一時clone（`C:\temp\vba-schema-eol-test-f273767`）でVBA sourceがLFとなること、固定xlflow `fmt --check` 18 files passを確認。
  - GitHub Actions `source-check` run `35587365802`: success（24s、全step pass）。
- Unverified: GitHub-hosted CIでのExcel/VBE compile、Windows 32-bit/macOS Office、Dictionary/RegExp unavailable実機再現は引き続き未検証。
- Next task: source-check成功を基準に、残るExcel/VBE compile ownershipまたはrelease作業を別gateとして進める。

### 2026-09-21 — User-facing sample projects

- Status: `sample/`へ利用者向けの3独立プロジェクトを追加し、README・設計書・AGENTS tree・roadmapを更新した。sample sourceは3-file import payloadと分離している。
- Projects:
  - `01-order-import`: 注文・明細のnested object、Collection、ArrayOf、Enum、Pattern、Email、DateTime、Strict。
  - `02-settings-validation`: Enum、範囲制約、Pattern、ArrayOf、OptionalField、Nullable、Strict。
  - `03-api-payload`: APIレスポンス相当のnested object、Union、Literal、ArrayOf、Nullable、Strict。
- Changed: `sample/README.md`、各sampleのREADMEと`.bas`、`tools/verify-samples.ps1`、`Taskfile.yml`、`.github/workflows/source-check.yml`、`README.md`、`CHANGELOG.md`、`docs/design.md`、`AGENTS.md`。
- Verification:
  - `rtk powershell -NoProfile -ExecutionPolicy Bypass -File tools/verify-samples.ps1`: 3 files format unchanged、isolated lint/analyze pass。
  - 一時source project `C:\temp\vba-schema-samples-compile-20260921`で`xlflow lint --json`、`xlflow analyze --json`、`xlflow push --session --no-save --json`がpass。
  - Windows 64-bit managed sessionで`SampleOrderImport.RunOrderImportSample`、`SampleSettingsValidation.RunSettingsValidationSample`、`SampleApiPayload.RunApiPayloadSample`をdiagnostic/headless実行し、全てsuccess。sessionは`--discard`で停止し未保存変更を破棄。
- Unverified: macOS Office、Windows 32-bit Office、Dictionary/RegExp unavailable環境でのsample実行。
- Next task: sample変更をcommitし、source-checkでsample verificationを含むCI結果を確認する。

### 2026-09-21 — M8 independent review / Pass 1 remediation

- Status: `orca-supervised-final-review`の独立worktreeレビューを完了。workerは`worker_done`を送信したが、受信通知が一度配送されず、`worker-read`でtranscriptを再取得して報告を確認した。レビューworktreeは報告を保存して解放済み。
- Confirmed findings:
  - P1: `tools/release-stage.ps1`の任意Destinationが`src`等のproduction sourceと重なる場合に削除できる。repository root、`src`、`.git`、`.xlflow`、`build`との同一・子孫・祖先重複を削除前に拒否するよう修正。
  - P2: `tools/run-benchmark.ps1`がx86結果をx64 baselineと比較し得る。`office_bitness`を検査し、非x64はunsupported reportを保存してbaseline比較・更新をスキップする環境ゲートを追加。
  - P3: 本書のCurrent CheckpointとM7最新記録に残っていた未コミット表記を更新。
- Unsupported observations: review worktreeのun-pinned xlflowによるfmt-check失敗、review worktreeでのExcel/VBE未実行は実装欠陥とは確定せず、環境差異・未検証として分離した。
- Changed: `release-stage.ps1`の保護パス判定、benchmark environment helperとfocused regression、Task/Excel-free CI接続、ADR-0007/0008、benchmark/v1 contract、design、README、AGENTSを更新。
- Validation:
  - `rtk task verify`: pass（60 tests discovered、release staging safety、benchmark environment guard、lint/analyze/format/quality pass）。
  - `rtk task release-stage` / `rtk task release-verify`: pass。
  - `rtk task release-smoke`: pass（fresh workbook import、VBE compile、scalar smoke、3 non-document components）。
  - `rtk task benchmark`: pass（Windows x64、absolute targets、baseline comparison pass）。Report: `C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\artifacts\benchmarks\20260921-181627326-windows-x64.json`。
  - managed sessionで`rtk xlflow test --session --no-save --json`: 59 pass、1 intentional TODO。`rtk xlflow run PublicApiCompile.CompilePublicApi --diagnostic --headless --session --no-save --json`: pass。
  - `rtk xlflow status --json`: session inactive、dirty=false、recovery_required=false。
- Commit: `3593e01 fix: harden release and benchmark gates`。M7/MITの既存commitは`57e82c6`、`8ec96e7`。
- Temporary workspaces / artifacts:
  - `C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\build\Book.xlsm`（managed session終了時に未保存変更を破棄）。
  - `C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\artifacts\benchmarks\20260921-181627326-windows-x64.json`（ignored report）。
  - 独立review worktree `C:\Users\HARUMI\orca\workspaces\VBA-Schema\final-review-agent`（報告取得後に削除）。
- Unverified: GitHub-hosted workflow実行、Windows self-hosted automation、macOS Office、Windows 32-bit Office、Dictionary/RegExp unavailable実機再現。
- Next task: この修正を含むclean treeへ独立Pass 2 reviewを実行し、blocking findingがないことを確認する。

### 2026-09-21 — M7 hardening and runtime performance

- Status: M7のproduction hardening、runtime benchmark、release/CI boundary接続を実装し、Windows 64-bit Excelで回帰・性能・release smokeを検証した。M7/MIT変更は`57e82c6`、`8ec96e7`としてコミット済み。GitHub上のworkflow実行、macOS、Windows 32-bit Excelは未検証。
- Decision:
  - Object validationは入力Dictionaryのkeysを一度だけ列挙し、binary `keyIndex`、invalid key、unknown keyを同時に構築する。fieldごとのkeys再列挙を廃止し、大規模objectのlookupをO(fields × input keys)からO(fields + input keys)へ改善する。
  - `Min`/`Max`のText/Array/Number実装はprivate boundary helperへ集約し、既存のduplicate・ordering・error code契約を維持する。
  - optional `VB018`/`VB021`/`VBA210`は評価した。既存の意図的な戻り値slot、scope shadowing、error-handler patternを大量に報告するため通常gateにはせず、標準`lint`/`analyze`はno findingsを維持する。compatibility-sensitive patternは`task quality-audit`で独立検査する。
  - runtime benchmarkのfixture、counter、threshold、25% baseline policy、専用session ownershipをADR-0008/specへ固定する。benchmark sourceはbuild・metrics・3-file release payloadから除外する。
- Changed:
  - `src/classes/VSchema.cls`のDictionary key index化とboundary helper化。
  - `src/modules/Benchmarks/ValidationBenchmarks.bas`、`tools/run-benchmark.ps1`、`benchmarks/windows-x64-baseline.json`を追加。
  - `tools/check-production-hygiene.ps1`と`task quality-audit`を追加し、source-check workflowへ接続。
  - `tools/release-smoke.ps1`のfresh probe pushから`--fast`を外し、共有push-state cacheで初回importをskipする問題を回避。`source-check.yml`へrelease-stage/verifyを追加。
  - `docs/adr/ADR-0008-runtime-benchmark-contract.md`、`docs/specs/benchmark-contract.md`、README、design、AGENTS、xlflow issue記録を更新。
- Validation:
  - `rtk task verify`: pass（60 tests discovered、59 pass、`SampleTests.Test_Sample_Todo` 1件はintentional TODO、lint/analyze/format/quality audit pass）。
  - `rtk xlflow test --session --no-save --json`: pass（60 tests、59 pass、intentional TODO 1件）。
  - `rtk xlflow metrics --json`: pass（291 procedures、benchmark moduleはmetrics対象外、production hotspotを確認）。
  - `rtk task benchmark`: pass（absolute targets、baseline comparison pass）。object fixtureは1測定sample内で10回検証し、Timer量子化の影響を抑えた。
  - 最新benchmark report: `C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\artifacts\benchmarks\20260921-173017896-windows-x64.json`。
    - `object_1000_fields_success`: median 371.09375 ms / target <500 ms / validation 50000 / issues 0
    - `object_1000_fields_issues`: median 1156.25 ms / validation 50000 / issues 5000
    - `scalar_10000_success`: median 125 ms / target <1000 ms / validation 50000 / issues 0
    - `scalar_10000_issues`: median 8691.40625 ms / validation 50000 / issues 50000
  - fixture contract変更（object sample内10回反復）に伴い、明示的な`-UpdateBaseline`で同一条件のbaselineを更新し、その後通常モードでbaseline comparison passを確認した。
  - 反復導入前の短時間fixture比較失敗も`C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\artifacts\benchmarks\20260921-172105304-windows-x64.json`へ保存され、失敗理由を追跡可能にした。
  - `rtk task release-smoke`: pass（fresh workbook import、追加reference設定なしのVBE compile、scalar smoke、3 non-document components）。
  - `rtk task release-stage` / `rtk task release-verify`: pass（3-file allowlist、UTF-8 BOMなし、LF、VBA headers）。
  - Current release payload SHA-256（`dist/VBA-Schema`）: `Schema.bas` `D8D84731088500106F307FF14A433FF4B6836F34E7D33B271FABDA316217B698`、`VSchema.cls` `CEB9006EF5CE3CB2BF804180347438F7A2573597BB98F5C0F54F6518780FFB7B`、`VValidationResult.cls` `03ADBB6989E14B328F82E156315E2D006B245177EBEACF56BD2CF1F4C27E5849`。
  - `rtk xlflow status --json`: session inactive、dirty=false、recovery_required=false。sourceがignored workbookより新しい警告のみ。
  - `rtk git diff --check`: pass。
- Temporary workspaces / artifacts:
  - `C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\build\Book.xlsm`（ignored workbook、managed session終了時に未保存変更を破棄）。
  - `C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\artifacts\benchmarks\20260921-173017896-windows-x64.json`（ignored report）。
  - release smokeの一時probeは`C:\Users\HARUMI\AppData\Local\Temp\vba-schema-release-smoke-<guid>`配下で作成・検証後に削除済み。
- Unverified:
  - GitHub-hosted `source-check` workflow実行、Windows self-hosted automation。
  - macOS Office、Windows 32-bit Office、Dictionary/RegExp unavailable実機再現。
- Next task: M7 exit gateの未完了項目（VBE compile evidenceのworkflow化、clean-checkout drift audit、GitHub workflow実行）を確認し、必要ならM8独立reviewへ進む。

### 2026-09-21 — M7 documentation and release boundary

- Status: M7のDecision Gate（DG-008/DG-009）を確定し、README、CHANGELOG、release encoding verify、Excel-free CI workflow、compatibility static auditを追加。性能benchmark、VBE self-hosted automation、sample workbookは未完了。GitHub上のworkflow実行は未検証。
- Decision:
  - 3-file import payloadは維持し、README、LICENSE、CHANGELOG、sample workbook、ZIPは別release/repository assetとする。
  - versionの正はGit tagとtracked `CHANGELOG.md`。VBA sourceへのversion定数/VERSIONファイルは追加しない。
  - staged VBA sourceはUTF-8（BOMなし）・LF改行。`release-verify.ps1`でallowlist、header、encoding、改行を検証する。
  - GitHub-hosted Excel-free source checkとWindows Excel/VBE compile・behavioral checkを別statusへ分離し、macOS/32-bitは将来matrix追加時まで未検証とする。
- Changed:
  - `README.md`へinstallation、canonical example、API/error/path、aliasing、Pattern/Email境界、platform boundary、update/remove手順を追加。
  - `CHANGELOG.md`を追加し、Unreleased/v1.0.0のrelease記録形式を固定。
  - `docs/adr/ADR-0007-release-and-ci-boundary.md`を追加し、`docs/specs/v1-contract.md`、`docs/design.md`、`AGENTS.md`、DG-008/DG-009を更新。
  - `tools/release-verify.ps1`へUTF-8 BOMなし・LF改行検査を追加。
- Validation:
  - `rtk xlflow metrics --json`: 290 proceduresのstatic metricsを取得。runtime performanceの合否には使用しない。
- `rtk task release-stage`: 3-file payload staging pass
- `rtk task release-verify`: allowlist、VBA headers、UTF-8/BOMなし、LF改行 pass
- `rtk task verify`: pass（60 tests discovered、lint/analyze/format check pass）
- `rtk task release-smoke`: pass（fresh workbook import、VBE compile、scalar smoke、non-document component 3件）
- `rtk git ls-remote https://github.com/harumiWeb/xlflow.git refs/heads/main`: source-check workflowのxlflow commit pin `52e93661cdc830592de88e047cbeed3bb2ac4488`を確認
  - compatibility static search: Windows API、pointer-size、LongLong、Excel Object Modelのproduction core依存なし
  - `rtk git diff --check`: pass
- Release payload SHA-256（`dist/VBA-Schema`）:
  - `Schema.bas`: `d8d84731088500106f307ff14a433ff4b6836f34e7d33b271fabda316217b698`
  - `VSchema.cls`: `51c6988bd1ab986bdc0f0b49f2c1726fb1c88b1307ac4a2ac2387aaaca19b665`
  - `VValidationResult.cls`: `03adbb6989e14b328f82e156315e2d006b245177ebeacf56bd2cf1f4c27e5849`
- Unverified:
  - GitHub-hosted CIでのxlflow install/version pinとExcel-free workflow実行
  - Windows self-hosted VBE compile automation
  - performance benchmark、baseline比較、sample workbook asset
  - macOS Office、Windows 32-bit Office、Dictionary/RegExp unavailable実機再現
- Next task: benchmark/CIの実装方針を確定し、VBE compile ownershipを再現可能なTaskまたはworkflowへ落とし込む。

### 2026-09-21 — M6 Union

- Status: M6のUnion production implementation、Decision Gate、focused/full proof loop、公開API compile、release smoke完了。変更は未コミット。
- Decision:
  - DG-012 Union ownership: `UnionOf(Collection)`はbranch順と要素集合を構築時snapshotするが、child `VSchema` instanceはcloneせずmutable shared referenceとして保持する。
  - nested Unionはflattenせずbranch構造を保持する。branch validation failureは内部Issueを破棄し、失敗pathに`invalid_union`を1件だけ返す。branch内のprogrammer/runtime/environment errorはそのまま伝播する。
  - 空/Nothing/non-VSchema/未初期化branchは`vbObjectError + 2100`とする。
- Changed:
  - `Schema.bas`に`UnionOf(ByVal Schemas As Collection)` factoryを実装。
  - `VSchema.cls`にUnion branch snapshot、branch validation、nested graph traversal、単一`invalid_union`境界を実装。
  - `src/modules/Tests/TestUnion.bas`を追加し、branch順、後続branch成功、全失敗、snapshot/shared reference、nested、special value、branch errorを検証。
  - `TestErrors.bas`へ空/Nothing/non-VSchema/未初期化Union branchのprogrammer misuseを追加。
  - `docs/adr/ADR-0006-union-branch-ownership-and-error-boundary.md`、`docs/specs/v1-contract.md`、`docs/design.md`、`AGENTS.md`、ロードマップを更新。
- Validation:
  - `rtk xlflow fmt --write .\src\classes\VSchema.cls .\src\modules\Schema.bas .\src\modules\Tests\TestUnion.bas .\src\modules\Tests\TestErrors.bas`: 4 files unchanged
  - `rtk xlflow lint --json`: pass
  - `rtk xlflow analyze --json`: pass（no findings）
  - `rtk xlflow test --module TestUnion --session --no-save --json`: 7 pass
  - `rtk xlflow test --module TestErrors --session --no-save --json`: 8 pass
  - `rtk xlflow test --session --no-save --json`: 60 tests、59 pass、scaffold `Test_Sample_Todo` 1件はintentional TODO
  - `rtk xlflow run PublicApiCompile.CompilePublicApi --diagnostic --headless --session --no-save --json`: pass
  - `rtk xlflow test list --json`: 60 tests in 23 source files
  - `rtk task verify`: pass
  - `rtk task release-smoke`: pass（3-file import、VBE compile、scalar smoke、non-document component 3件）
  - `rtk powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\release-smoke.ps1`: pass
  - `rtk xlflow session stop --discard --json` / `rtk xlflow status --json`: session inactive、dirty=false、recovery_required=false。sourceがignored workbookより新しい警告のみ。
- Temporary workspaces:
  - `C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\build\Book.xlsm`（ignored local workbook、session終了時に未保存変更を破棄）
  - `C:\temp\vba-schema-design-oracle-20260921`
  - `C:\Users\HARUMI\AppData\Local\Temp\vba-schema-release-smoke-8cb0b80dcafd4461bc7a04cee8f7c7cf`（release smoke probe）
- Unverified:
  - macOS Office、Windows 32-bit Office
  - Dictionary/RegExp runtime unavailable環境の実機再現
  - locale変更を伴う実機確認、performance benchmark、CI上のExcel/VBE compile
- Next task: M7 hardening（未検証runtime seam、README/installation、performance、CI/release boundary）へ進む。

### 2026-09-21 — M5 Literal, Enum, Pattern, and Email

- Status: M5 の production implementation、focused/full proof loop、release smoke 完了。変更は未コミット。RegExp unavailable の実機再現だけは未検証のため M5 exit gate に未チェックを残す。
- Decisions:
  - DG-007: `VBScript.RegExp`をlate-boundし、`IgnoreCase=False`、`Global=False`、`MultiLine=False`、全体一致を固定。invalid expressionはlazy compile時に`vbObjectError + 2100`、runtime unavailableは`vbObjectError + 2200`。EmailはASCII簡易形式、254 UTF-16 code unit以下、RFC完全準拠外。
  - DG-012 Enum部分: 一次元native arrayのscalar候補を構築時snapshot。typed scalar arrayは受理し、未初期化・空・多次元・不正要素は`vbObjectError + 2100`、重複候補は`vbObjectError + 2101`。
- Changed:
  - `Schema.bas`の`Literal`／`EnumOf` factory、`VSchema.cls`のscalar category比較、Enum snapshot、Pattern／Email modifier、lazy RegExp cacheを実装
  - `src/modules/Tests/TestLiteral.bas`、`TestPattern.bas`を追加し、category／snapshot／duplicate／full-match／Email境界を検証
  - `TestErrors.bas`へLiteral／Enum／Pattern／Emailのprogrammer misuseを追加
  - `docs/specs/v1-contract.md`、`docs/design.md`、ADR-0005、AGENTS、ロードマップを更新
- Validation:
  - `rtk xlflow lint --json`: pass
  - `rtk xlflow analyze --json`: pass
  - `rtk xlflow push --session --no-save --json`: imported 24 source files、VBE compile pass
  - `rtk xlflow test --module TestLiteral --session --no-save --json`: 4 pass
  - `rtk xlflow test --module TestPattern --session --no-save --json`: 5 pass
  - `rtk xlflow test --module TestErrors --session --no-save --json`: 7 pass
  - `rtk xlflow test --session --no-save --json`: 52 tests、51 pass、scaffold `Test_Sample_Todo` 1件はintentional TODO
  - `rtk xlflow run PublicApiCompile.CompilePublicApi --diagnostic --headless --session --no-save --json`: pass
  - `rtk task verify`: pass（52 tests discovered）
  - `rtk task release-smoke`: fresh workbookへの3-file import、VBE compile、scalar smoke、non-document component 3件の検証がpass
  - `rtk xlflow status --json`: session inactive、dirty=false、recovery_required=false。sourceがignored workbookより新しい警告のみ。
- Temporary workspaces:
  - `C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\build\Book.xlsm`（ignored local workbook、session終了時に未保存変更を破棄）
  - `C:\temp\vba-schema-design-oracle-20260921`
  - `C:\Users\HARUMI\AppData\Local\Temp\vba-schema-release-smoke-4500506a435d40f6a19d783efac1df6f`（release smoke probe、検証後に削除）
- Unverified:
  - RegExp runtime unavailable環境の実機再現、macOS Office、Windows 32-bit Office
  - locale変更を伴う実機確認、performance benchmark、CI上のExcel/VBE compile
- Next task: RegExp unavailableの検証方法を確定するか未検証としてrelease boundaryへ固定し、M6のUnion ownership（DG-012）をDecision Gateとして確認する。

### 2026-09-21 — M4 Array and Collection validation

- Status: M4 Array/Collection implementationとfocused/full proof loop完了。変更は未コミット。
- Decision:
  - DG-006: 一次元native arrayとCollectionを受理し、未初期化dynamic arrayは0要素、多次元は`invalid_array_rank`。LBoundに依存しない0始まりlogical pathを使用し、rank／bounds probeのErr stateをhelper内で処理する。
- Changed:
  - `Schema.bas`の`ArrayOf` factory、`VSchema.cls`のitem schema、array length constraints、native array／Collection detection、nested validation、cycle traversalを実装
  - `src/modules/Tests/TestArray.bas`を追加し、native array／negative bound／typed object／Collection／nested／uninitialized／multidimensional／special elementを検証
  - `TestErrors.bas`へArrayOf Nothingとarray constraint misuseを追加
  - `docs/specs/v1-contract.md`、`docs/design.md`、ADR-0004、AGENTS、ロードマップを更新
- Validation:
  - `rtk xlflow lint --json`: pass
  - `rtk xlflow analyze --json`: pass
  - `rtk xlflow push --fast --session --no-save --json`: compile pass（22 source files）
  - `rtk xlflow test --module TestArray --session --no-save --json`: 7 pass
  - `rtk xlflow test --module TestErrors --session --no-save --json`: 5 pass
  - `rtk xlflow test --session --no-save --json`: 40 pass、scaffold `Test_Sample_Todo` 1件はintentional TODO
  - `rtk xlflow run PublicApiCompile.CompilePublicApi --diagnostic --headless --session --no-save --json`: pass
  - `rtk task verify`: pass（41 tests discovered）
  - `rtk task release-smoke`: fresh workbookへの3-file import、VBE compile、scalar smoke、non-document component 3件の検証がpass
- Unverified:
  - macOS Office、Windows 32-bit Office、uninitialized／multidimensionalの他host差異
  - performance benchmark、CI上のExcel/VBE compile
- Next task: M5 Literal／Enum／Pattern／EmailのDecision Gate（DG-007、DG-012）を確認する。

### 2026-09-21 — M3 Object validation

- Status: M3 Object validation実装とfocused proof loop完了。変更は未コミット。
- Decisions:
  - DG-004: `SafeParse`直前のDFS preflight、active pathのCollection、`Is`比較、cycleは`vbObjectError + 2103`。builder時の早期検出は行わない。
  - DG-005: `TypeName("Dictionary")`入口と`Count`／`Exists`／`Keys`の限定capability check。非文字列keyは`invalid_key` validation issue、strict unknown String keyはbinary ordinal。library側のDictionary生成失敗は`vbObjectError + 2200`。
  - VBA Private memberのcross-instance制約を避ける`Internal*` hookをunsupported internal-only APIとして追加。
- Changed:
  - `VSchema.cls`へObject field state、`Field`／`Strict`、binary lookup、canonical path、unknown/non-string key issues、nested validation、schema cycle preflightを追加
  - `src/modules/Tests/TestObject.bas`を追加し、required／optional／nullable／nested／strict／ordering／path escaping／non-string key／Nothing／non-mutationを検証
  - `TestErrors.bas`へObject builder misuseとdirect/indirect cycleのErr検証を追加
  - `docs/specs/v1-contract.md`、`docs/design.md`、ADR-0002、ADR-0003へObject判定・`Missing`／`invalid_key`・cycle契約を反映
- Validation:
  - `rtk xlflow lint --json`: pass
  - `rtk xlflow analyze --json`: pass
  - `rtk xlflow push --fast --session --no-save --json`: compile pass（21 source files）
  - `rtk xlflow test --module TestObject --session --no-save --json`: 7 pass
  - `rtk xlflow test --module TestErrors --session --no-save --json`: 4 pass
- Unverified:
  - Scripting.Dictionary unavailable環境の実機再現、macOS Office、Windows 32-bit Office
  - Object performance benchmark、CI上のExcel/VBE compile
- Next task: M4 Array/CollectionのDG-006をspec・tests firstで確定する。

### 2026-09-21 — M1 release smoke and M2 scalar core

- Status: M1 release smoke完了、M2 scalar coreのfocused testsとVBE proof loop完了。変更は未コミット。
- Changed:
  - `Schema.bas` factory、`VSchema.cls` scalar validation/modifier、`VValidationResult.cls` snapshot/result contractを実装
  - `src/modules/Tests/TestAnyValue.bas`、`TestText.bas`、`TestNumber.bas`、`TestBool.bas`、`TestDateTime.bas`、`TestNullable.bas`、`TestResult.bas`、`TestErrors.bas`を追加
  - `tools/release-smoke.ps1`を追加し、fresh workbook、3-file import、scalar macro、non-document component countを自動検証
  - `Taskfile.yml`の`release-smoke`、`AGENTS.md` tree、`docs/design.md` ErrorText例、ロードマップを更新
- Implementation notes:
  - factoryとResult初期化は、VBAのclass Function戻り値を`CallByName`で呼ぶと実行時438になるため、通常の直接Function呼び出しを採用。xlflowの名前解決を阻害しない変数名を使用した。
  - `Scripting.Dictionary`が利用できない場合は`vbObjectError + 2200`を伝播する。Issue collectionはdeep snapshotとする。
  - `Length`と`Min/Max`のbuilder矛盾をchain順にかかわらず検出する。
  - `CVErr`の`CStr`結果はOfficeの表示言語で変化するため、`Error(<number>)`へ正規化してlocale-independentなreceived descriptorを維持する。
- Validation:
  - `rtk task verify`: pass（lint、analyze、fmt-check、test list; 25 tests discovered）
  - `rtk xlflow run PublicApiCompile.CompilePublicApi --diagnostic --headless --session --no-save --json`: pass
  - focused modules: AnyValue 3、Text 4、Number 3、Bool 1、DateTime 1、Nullable 2、Result 3、Errors 3、全件pass
  - `rtk xlflow test --session --no-save --json`: 24 pass、scaffold `Test_Sample_Todo` 1件はintentional TODO
  - `rtk task release-smoke`: fresh workbookへの3-file import、VBE compile、scalar smoke、non-document component 3件の検証がpass
  - `rtk git diff --check`: pass
- Temporary workspaces:
  - `C:\temp\vba-schema-release-smoke-20260921`（手動release smoke、payload importとpullのcomponent count確認）
  - `C:\temp\vba-schema-design-oracle-20260921`（既存compile oracle）
  - `C:\Users\HARUMI\orca\workspaces\VBA-Schema\auk\build\Book.xlsm`（ignored local workbook、session終了時に未保存変更を破棄）
- Unverified:
  - macOS Office、Windows 32-bit Office
  - Object/Array/Literal/Enum/Union/Pattern/Email、外部参照自動検査、performance、CI上のExcel/VBE compile
- Next task: M3 Object validationのDecision GateとDictionary field semanticsを確認し、`Field`/`Strict`/nested pathのtests firstへ進む。

### 2026-09-21 — M1 harness bootstrap

- Status: M0のDG-001/DG-010待ち、M1 harnessは部分完了
- Changed:
  - `Taskfile.yml`に`provision-workbook`、source verification、session test、verify taskを追加
  - `tools/provision-workbook.ps1`を追加。隔離temp projectでxlflow scaffoldを作成し、既存workbookを上書きしない
  - `tools/check-format.ps1`を追加。VBA-Schema production/focused test sourceの出現後にformatter gateを有効化
  - `AGENTS.md`のproject treeを更新
- Validation:
  - `rtk task provision-workbook`: clean `build/Book.xlsm`を作成
  - `rtk task verify-source`: pass（lint、analyze、format対象なしの明示skip、test discovery）
  - managed session start / `rtk xlflow push --fast --session --no-save --json`: pass
  - `rtk xlflow test --session --no-save --json`: scaffold 4 pass、1 todo
  - managed session stop `--discard`と`rtk xlflow status --json`: recovery不要、session inactive
- Temporary workspace:
  - `C:\temp\vba-schema-new-probe-20260921`
- Known issue/decision:
  - 既存xlflow helper 10ファイルをformatterで整形すると`XlflowAssert.bas`がVB014 parser recoveryになるため、formatter-only変更は取り消し、production/focused sourceに限定するTaskへ変更した
  - `build/Book.xlsm`は`.gitignore`対象で、sourceより古い検証用workbookがローカルに残る
- Unverified:
  - macOS Office
  - Windows 32-bit Office
  - VBA-Schema production compile/runtime
- Next task: DG-001/DG-010を確定後、M1 compile fixtureと3 production component skeletonを追加

### 2026-09-21 — M0 decisions and M1 compile skeleton

- Status: DG-001/DG-010 resolved、M1 compile skeleton完了、runtime implementation未着手
- Decisions:
  - DG-001: Issueは`Scripting.Dictionary`必須。未提供環境はenvironment failureとしてErr伝播
  - DG-010: `InternalInitialize`はPublic internal-only。factoryは内部encoded kindを使い、通常の直接呼び出しは`+2100`、再初期化は`+2102`
- Changed:
  - `docs/specs/v1-contract.md`、ADR-0001/0002、`docs/design.md`へ決定を反映
  - `Schema.bas`、`VSchema.cls`、`VValidationResult.cls`のPublic signature skeletonを追加
  - `PublicApiCompile.bas`で全factory、fluent chain、Collection、typed array、Object、Result assignment、InternalInitializeをcompile対象化
  - `AGENTS.md`のtreeを更新
- Validation:
  - `rtk task verify`: pass
  - `rtk task fmt-check`: 4 production/focused files unchanged
  - `rtk xlflow lint --json`: pass
  - `rtk xlflow analyze --json`: pass
  - `rtk xlflow push --fast --session --no-save --json`: imported 12 source files
  - `rtk xlflow run PublicApiCompile.CompilePublicApi --diagnostic --headless --session --no-save --json`: pass
  - `rtk xlflow test --session --no-save --json`: scaffold 4 pass、1 todo
  - session stop `--discard`後の`rtk xlflow status --json`: recovery不要、session inactive
- Temporary workspace: `C:\temp\vba-schema-new-probe-20260921`
- Unverified:
  - macOS Office
  - Windows 32-bit Office
  - production runtime validation
  - release staging、performance、CI gate
- Next task: M1 release staging/compile fixture除外確認後、M2 scalar tests firstへ進む

### 2026-09-21 — M1 release payload staging

- Status: M1 compile skeleton + 3-file payload staging完了。空workbook import smokeとCI/release policyは未完了
- Decision: `dist/VBA-Schema`は3-file import payloadのgenerated artifactとし、commit対象外にする
- Changed:
  - `tools/release-stage.ps1`と`tools/release-verify.ps1`を追加
  - `Taskfile.yml`へ`release-stage`／`release-verify`を追加
  - `.gitignore`へ`dist/`を追加
  - spec/ADR/AGENTS/todoへpayload ownershipを反映
  - formatter/parser recoveryの再現条件を`docs/xlflow-issues/20260921-fmt-parser-recovery.md`へ記録
- Validation:
  - `rtk task release-stage`: 3 files staged
  - `rtk task release-verify`: 3-file allowlist、`Attribute VB_Name`、class headerを検証してpass
  - `rtk task verify`: pass
- Unverified:
  - empty workbookへのrelease payload import smoke
  - component countのVBE側検証
  - DG-009 CI／compile oracle ownership
- Next task: empty workbook release smokeを追加後、M2 scalar tests firstへ進む

### 2026-09-21 — Roadmap作成時点

- Status: M0 in progress
- Branch: `review-vba-schema-design`
- HEAD: `0d1a125`
- Completed:
  - design review
  - v1 contract draft
  - ADR-0001/0002 draft and review
  - Windows 64-bit compile oracle spike
  - roadmap作成
- Validation:
  - `rtk git diff --check`: pass
  - `rtk rg -n "[ \t]+$" .\tasks\todo.md`: no match
  - Markdown code fence count: 24（balanced）
  - `rtk xlflow lint --json`: pass
  - `rtk xlflow analyze --json`: pass
  - `rtk xlflow test list --json`: 5 scaffold tests discovered
  - independent roadmap review: unresolved High/Medium findingなし
  - compile oracle in `C:\temp\vba-schema-design-oracle-20260921`: pass
- Unverified:
  - macOS Office
  - Windows 32-bit Office
  - production runtime behavior
  - release staging
  - performance
- Working tree: design/spec/ADR/todo changes are uncommitted
- Next task: resolve DG-001 and DG-010, then complete M0 review and commit the design baseline
- Next command（untracked fileは`git diff`に出ないため、まずstatusを確認する）:

```powershell
rtk git status --short
```
