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
- Baseline HEAD: `0d1a125 add skills adr-manager`
- production implementation: 未着手
- production files: まだ存在しない
- 現在のtest: xlflow scaffoldの`SampleTests.bas`のみ
- xlflow configured workbook: `build/Book.xlsm`
- xlflow session: inactive
- 現在の設計変更は未コミット

現在の未コミット対象:

```text
docs/design.md
docs/specs/v1-contract.md
docs/adr/ADR-0001-small-distribution-and-portable-core.md
docs/adr/ADR-0002-vba-safe-api-and-error-boundary.md
tasks/todo.md
```

### 2.2 Confirmed evidence

- Windows 64-bit OfficeのVBEで、v1 Public API候補の宣言とfluent chainがcompile成功している。
- typed native array、late-bound Object、Result propertyのObject/scalar assignmentを含むcompile fixtureが成功している。
- `rtk xlflow lint --json`: success
- `rtk xlflow analyze --json`: success
- `rtk xlflow test list --json`: scaffold testを5件検出
- `rtk git diff --check`: tracked変更はsuccess（untracked文書は対象外）
- ADR-0001/0002: lint/review上のHigh/Medium findingなし

過去のcompile oracle用一時workspace:

```text
C:\temp\vba-schema-design-oracle-20260921
```

この一時workspaceは参考証拠であり、再現可能なrepository fixtureの代わりにしてはならない。

### 2.3 Explicitly unverified

- macOS Office
- Windows 32-bit Office
- production implementationのruntime behavior
- performance target
- 3-file release stagingとartifact verification
- READMEのinstallation手順
- CI上のExcel/VBE compile

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

- [ ] 現行の`Scripting.Dictionary`必須契約を維持するか、key accessだけを公開契約としてportable fallbackを許可するか決める。
- [ ] macOSでDictionaryが利用できない場合でもscalar validation errorを表現可能にする必要があるか決める。
- [ ] 決定を`docs/specs/v1-contract.md`へ反映する。
- [ ] 3-file constraintへの影響があればADR-0001を更新またはsupersedeする。
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

- [ ] Byte/Integer/Long/Single/Double/Currency間の`Min`/`Max`比較手順を決める。
- [ ] 無条件な`CDbl`を避けながらoverflowと精度損失をどう検出するか決める。
- [ ] NaN、Infinity、overflow-producing comparisonの検出方法を決める。
- [ ] `WholeNumber`のSingle/Double/Currency判定手順を決める。
- [ ] 決定表と境界test casesをspecへ追加する。

禁止事項:

- precision lossを黙って成功扱いする
- BooleanまたはDateをNumberとして扱う
- Error handlingを通常のvalidation分岐として広範囲に使用する

### DG-003 Error message and received descriptor grammar

期限: M2の`VValidationResult`実装前

- [ ] `received`の正確なString grammarをtype別に決める。
- [ ] Stringのescape、80 UTF-16 code unitでのtruncate、truncate markerを決める。
- [ ] Numberのlocale-independent formattingを決める。
- [ ] Dateの`yyyy-mm-ddThh:nn:ss` formattingと秒未満の扱いを決める。
- [ ] `ErrorText`の改行、indent、issue間separatorを固定する。
- [ ] raw secretをmessageへ含める範囲を決める。
- [ ] golden testsを追加する。

### DG-004 Schema cycle detection

期限: M3開始前

- [ ] specどおり、実際のvalue validation開始前にschema graph preflightを必ず行う。
- [ ] builder時の早期検出も追加する場合は、Err発生時点の変更として先にspecを更新する。
- [ ] direct cycleとindirect cycleの検出方法を決める。
- [ ] mutable builderによる後発cycleを検出できることを確認する。
- [ ] `vbObjectError + 2103`の発生時点をspecへ固定する。

### DG-005 Dictionary detection and key matching

期限: M3開始前

- [ ] 正式な`Scripting.Dictionary`判定方法を決める。
- [ ] `TypeName = "Dictionary"`だけに依存するか、限定的なcapability checkを併用するか決める。
- [ ] 入力Dictionaryの`CompareMode`に依存せずbinary matchingするalgorithmを決める。
- [ ] String以外のDictionary keyをvalidation failure、programmer misuse、environment errorのどれにするか決める。
- [ ] strict unknown keyのbinary ordinal sort方法を決める。
- [ ] Dictionaryが利用できない環境のErr発生点を決める。

### DG-006 Native array inspection

期限: M4開始前

- [ ] 未初期化dynamic arrayを安全に0要素と判定するhelperを設計する。
- [ ] array rankの安全な取得方法を決める。
- [ ] 一次元以外を`invalid_array_rank`にする分岐を決める。
- [ ] `LBound`に関係なく0始まりlogical pathへ変換する方法を決める。
- [ ] 必要最小範囲の`On Error`と必ず復元するErr stateをspecまたは実装commentへ記録する。

### DG-007 Pattern and Email semantics

期限: M5開始前

- [ ] `VBScript.RegExp`の`IgnoreCase`、`Global`、`MultiLine`を固定する。
- [ ] invalid patternをlazy compile時のprogrammer misuseとして投げる詳細を決める。
- [ ] Emailの許容例・拒否例と長さ上限を決める。
- [ ] EmailをRegExpで実装するかpure VBAで実装するか決める。
- [ ] RegExp runtime unavailable時の`vbObjectError + 2200` test方法を決める。
- [ ] RFC完全準拠を目標にしない範囲をspecへ例示する。

### DG-008 Release artifact format

期限: M1完了前

- [ ] 現行specどおり、import payloadは3 source filesだけに固定する。
- [ ] ZIP、sample workbook、LICENSE、CHANGELOG等を3-file payloadとは別のrelease page/repository assetとして提供するか決める。
- [ ] 追加物を同一artifactへ同梱する案を採る場合は、実装前にspecの「3ファイルだけ」を更新し、ADR-0001を更新またはsupersedeする。
- [ ] `dist/VBA-Schema/`をcommit対象にするかgenerated artifactにするか決める。
- [ ] version metadataの保持場所を決める。
- [ ] CHANGELOGとrelease noteをv1で作成するか決める。
- [ ] release stagingで改行コードとencodingを固定する。
- [ ] `THIRD_PARTY_NOTICES.md`の要否を確認する。

### DG-009 CI and compile oracle ownership

期限: M1完了前

- [ ] GitHub-hosted runnerで実行するcheckと、Excel/VBEが必要なlocal/self-hosted checkを分離する。
- [ ] merge/release時に誰がVBE compile evidenceを取得するか決める。
- [ ] ExcelのないCIで「compile passed」と誤表示しないstatus名を決める。
- [ ] macOS/32-bit検証が将来提供された場合のmatrix追加位置を決める。

### DG-010 InternalInitialize exposure

期限: M1開始前

- [ ] `Schema.bas`からclassを初期化するための`InternalInitialize`のvisibilityと命名を確定する。
- [ ] 利用者が直接呼び出した場合、再初期化した場合、不正kind codeを渡した場合のErrを固定する。
- [ ] unsupported internal APIであることをREADMEへ露出するか、source commentだけにするか決める。
- [ ] v1 public compatibility guaranteeの対象外であることをspecとcompile fixtureで区別する。
- [ ] user-facing exampleとは別のinternal compile fixtureで`InternalInitialize`を直接呼び、存在、再初期化Err、不正kind Errを検証する。

### DG-011 Constraint composition

期限: M2実装前

- [ ] `Length(3).Min(4)`と`Length(3).Max(2)`をbuilder時に拒否するか決める。
- [ ] `Min`/`Max`と`WholeNumber`の組合せを許可する条件を決める。
- [ ] `Pattern`と`Email`を併用した場合の評価順を決める。
- [ ] specどおり、同じconstraint/modifierの重複指定はprogrammer misuseとして実装する。
- [ ] `Nullable`、`OptionalField`、`Strict`の重複をどのErr codeへ割り当てるか決める。
- [ ] constraint評価順と、同じ値に複数issueを出すか最初の1件だけにするか決める。
- [ ] 決定をspecとfocused testsへ反映する。

### DG-012 Schema input ownership

期限: EnumはM5開始前、UnionはM6開始前

- [ ] `EnumOf`へ渡したarrayをsnapshotするか共有するか決める。
- [ ] `UnionOf`へ渡したCollectionをsnapshotするか共有するか決める。
- [ ] child `VSchema` instance自体はmutable shared referenceとする既存契約との境界を整理する。
- [ ] duplicate Enum candidate、typed candidate array、多次元candidate array、未初期化candidate arrayの扱いを決める。
- [ ] Union branchの詳細issueを内部保持するか破棄するか決める。
- [ ] nested Unionをflattenするかbranch構造を保持するか決める。

## 5. M0 — Design baseline and decision gates

目的: 現在の設計変更をレビュー可能なbaselineとして確定し、実装者が異なる判断をしない状態にする。

### Tasks

- [ ] 現在の`docs/design.md`、spec、ADR-0001/0002、本ロードマップをまとめてreviewする。
- [ ] `docs/specs/v1-contract.md`の全Public signatureとcompile fixture候補が一致することを再確認する。
- [ ] design内に旧API名（`Schema.String`、`.Optional`、`.Integer`等）が残っていないことを検索する。
- [ ] v1 APIとfuture API（Parse、Strip、Positive/Negative等）の境界を確認する。
- [ ] Decision GateのうちM1開始前に必要なDG-001とDG-010を解決する。
- [ ] 設計変更をcommitし、baseline SHAを本書へ記録する。

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

- [ ] spec/ADR/design間のHigh/Medium矛盾がない。
- [ ] DG-001が解決済み。
- [ ] DG-010が解決済み。
- [ ] baseline commit SHAがCurrent Checkpointに反映済み。
- [ ] macOS/32-bitの未検証表現が全資料で一致している。

## 6. M1 — Development, test, and release harness

目的: production実装を始める前に、compile、test、releaseを再現可能にする。

### Repository structure

- [ ] `src/classes/`を作成する。
- [ ] `src/modules/Tests/`をv1 test構成へ整理する。
- [ ] compile-only fixtureの配置場所を決めて追加する。
- [ ] clean checkoutから`build/Book.xlsm`を作成または復元するbootstrap手順を実装する。
- [ ] bootstrap元となるtracked template/source、Excel prerequisite、既存workbookを上書きしない条件を文書化する。
- [ ] `dist/VBA-Schema/`を生成するallowlist staging scriptまたはTaskを追加する。
- [ ] staging用一時projectを安全に作成・破棄できるようにする。
- [ ] `Taskfile.yml`のplaceholder taskを実作業用taskへ置き換える。
- [ ] 新しいdirectory/fileを追加した時点で`AGENTS.md`のproject treeを更新する。

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
```

`task provision-workbook`は、ignored/untrackedの`build/Book.xlsm`が存在しないclean checkoutを正式な開始状態として扱う。tracked sourceまたはtemplateから再構築し、Excel/VBIDEなどの前提不足は明示的に失敗させる。既存のuser-owned workbookを無断で上書きしてはならない。

### Compile fixture

- [ ] 3 production componentへspecどおりのPublic signature skeletonを追加する。
- [ ] 未実装methodがfake successを返さず、明示的なinternal/not-implemented Errで停止するようにする。
- [ ] spec記載の全factoryをfixture内で呼び出す。
- [ ] 全fluent modifierをchain内で呼び出す。
- [ ] `UnionOf(Collection)`をcompileする。
- [ ] typed native arrayを`SafeParse(ByVal Variant)`へ渡す。
- [ ] late-bound Objectを`SafeParse`へ渡す。
- [ ] `Result.Value`をscalar代入とObjectの`Set`代入の両方でcompileする。
- [ ] `Success`、`Issues`、`ErrorText`を使用する。
- [ ] internal compile fixtureで`InternalInitialize`を直接呼び、DG-010で決めたErr contractも検証する。
- [ ] fixtureがrelease artifactへ含まれないことをtestする。

### Release staging

- [ ] DG-008を解決する。
- [ ] DG-009を解決する。
- [ ] `dist/VBA-Schema/`をimport payload専用directoryとし、release metadataやsampleを混在させない。
- [ ] source allowlistを次の3ファイルへ固定する。

```text
src/modules/Schema.bas
src/classes/VSchema.cls
src/classes/VValidationResult.cls
```

- [ ] allowlistに不足ファイルがあれば明示的に失敗する。
- [ ] allowlist以外の`.bas`/`.cls`/`.frm`をartifactへ入れない。
- [ ] 空のmacro-enabled workbookへ3ファイルをimportするsmoke pathを用意する。
- [ ] workbook document moduleを除くimport対象componentが3件であることを検証する。
- [ ] 外部参照設定が増えていないことを確認する。

### xlflow baseline

- [ ] `rtk xlflow doctor --json`でExcel/COM/VBIDE環境を確認する。
- [ ] clean checkoutで`task provision-workbook`を実行し、その後のtest discoveryまで再現する。
- [ ] `rtk xlflow status --json`でrecovery不要を確認する。
- [ ] sourceとworkbookのauthorityを確定する。
- [ ] managed sessionを使うか既存workbookへattachするか記録する。
- [ ] scaffold testの扱い（削除またはtest harness smokeとして保持）を決める。
- [ ] formatter実行前後のdiffを確認するTaskを用意する。

### Exit gate

- [ ] repository内のfixtureから全Public APIのVBE compileを再現できる。
- [ ] clean checkoutからdevelopment workbookを再構築できる。
- [ ] `task verify`相当の一括checkが存在する。
- [ ] 3-file stagingとcomponent count verificationが自動化されている。
- [ ] 開発用workbookと配布artifactが混同されない。

## 7. M2 — Scalar core and result contract

目的: AnyValue/Text/Number/Bool/DateTime、Result、共通modifier、error contractを完成させる。

### Tests first

- [ ] `TestAnyValue.bas`を追加する。
- [ ] `TestText.bas`を追加する。
- [ ] `TestNumber.bas`を追加する。
- [ ] `TestBool.bas`を追加する。
- [ ] `TestDateTime.bas`を追加する。
- [ ] `TestNullable.bas`を追加する。
- [ ] `TestResult.bas`を追加する。
- [ ] `TestErrors.bas`を追加する。
- [ ] root path `$`のtestを追加する。
- [ ] success/failure双方の`Value`、`Issues`、`ErrorText` contract testを追加する。
- [ ] Issues snapshotを呼び出し側が変更してもResultが変化しないtestを追加する。
- [ ] programmer misuseとenvironment/internal errorのErr番号testを追加する。

### Production implementation

- [ ] `src/modules/Schema.bas`のsignature skeletonを実装へ置き換える。
- [ ] `src/classes/VSchema.cls`のsignature skeletonを実装へ置き換える。
- [ ] `src/classes/VValidationResult.cls`のsignature skeletonを実装へ置き換える。
- [ ] `Option Explicit`を全production componentへ付ける。
- [ ] `SchemaKind`と一度だけ実行可能な`InternalInitialize`を実装する。
- [ ] 未初期化schemaと再初期化をprogrammer misuseにする。
- [ ] `SafeParse(ByVal InputValue As Variant)`を実装する。
- [ ] `AnyValue`を実装する。
- [ ] `Text`を`vbString`だけに限定する。
- [ ] `Number`をByte/Integer/Long/Single/Double/Currencyだけに限定する。
- [ ] `Bool`を`vbBoolean`だけに限定する。
- [ ] `DateTime`を`vbDate`だけに限定する。
- [ ] Null、Empty、Error Variant、Nothingのmatrixをspecどおり実装する。
- [ ] `Nullable`を実装する。
- [ ] `OptionalField`のstateを実装し、root値やArray要素の`Empty`を許可しないことをtestする。
- [ ] `Min`、`Max`、`Length`、`WholeNumber`を適用可能kindだけに実装する。
- [ ] duplicate/conflicting constraintをprogrammer misuseにする。
- [ ] `VValidationResult`のSuccess/Failure invariantを実装する。
- [ ] Issue code、path、message、expected、receivedを生成する。
- [ ] Issue orderとErrorTextをdeterministicにする。

### Required scalar edge cases

- [ ] empty StringとLength 0
- [ ] String Min/Max/Length boundary
- [ ] boundary引数`3`、`3&`、`CByte(3)`、`3#`
- [ ] fractional/negative/overflow/non-numeric length rejection
- [ ] Byte/Integer/Long/Single/Double/Currency
- [ ] Decimal Variant rejection
- [ ] LongLong rejection（使用可能な64-bit環境のみ）
- [ ] numeric String rejection
- [ ] Boolean/DateをNumberとして拒否
- [ ] positive/negative zero
- [ ] Single/DoubleのfractionとWholeNumber
- [ ] Null、Empty、Error Variant
- [ ] `AnyValue`のNull/Empty/Error/Nothing
- [ ] `AnyValue`のscalar/Object/native array/Collection
- [ ] rootまたはArray要素の`OptionalField`がEmptyを許可しない
- [ ] invalid modifier/kind combinations

### Decision gates

- [ ] DG-002 resolved
- [ ] DG-003 resolved
- [ ] DG-010 resolved
- [ ] DG-011 resolved

### Verification loop

```powershell
rtk xlflow lint --json
rtk xlflow analyze --json
rtk xlflow push --fast --session --no-save --json
rtk xlflow test --filter TestText --session --no-save --json
rtk xlflow test --session --no-save --json
```

### Exit gate

- [ ] scalar focused tests pass。
- [ ] full tests pass。
- [ ] VBE compile passes。
- [ ] `lint` and `analyze` pass with no unexplained findings。
- [ ] compile fixture and 3-file release smoke still pass。

## 8. M3 — Object validation

目的: Dictionary Object、Field、OptionalField、nested path、Strictを完成させる。

### Tests first

- [ ] required field present/missing
- [ ] `OptionalField` missing success
- [ ] present `Empty` is not missing
- [ ] present `Null` requires `Nullable`
- [ ] nested Object success/failure
- [ ] multiple deterministic issues
- [ ] passthrough unknown field
- [ ] strict unknown field
- [ ] strict unknown field binary ordinal ordering
- [ ] field declaration ordering independent of Dictionary enumeration
- [ ] case-sensitive field matching
- [ ] input `CompareMode = vbTextCompare`でもbinary field matching
- [ ] duplicate field definition is programmer misuse
- [ ] empty field name is programmer misuse
- [ ] `Field(name, Nothing)` is programmer misuse
- [ ] wrong input Object/class/Collection rejection
- [ ] Nothing rejection
- [ ] `Scripting.Dictionary` runtime unavailable時にenvironment errorがvalidation issueへ変換されないことを検証する。
- [ ] 実環境でunavailableを再現できない場合は、限定的なfactory seamによるerror-path testとWindows 64-bit integration結果を分け、未再現事項を記録する。
- [ ] escaped path for dot, bracket, quote, backslash, control characters
- [ ] direct and indirect schema cycles

### Implementation

- [ ] DG-004を解決する。
- [ ] DG-005を解決する。
- [ ] `ObjectSchema` factoryを実装する。
- [ ] schema field lookupとfield declaration orderを別管理する。
- [ ] input DictionaryのCompareModeに依存しないbinary lookupを実装する。
- [ ] `Field`、`OptionalField`、`Strict`を実装する。
- [ ] unknown fieldをbinary ordinalで並べる。
- [ ] canonical object pathを構築する。
- [ ] validation中にinput Dictionaryを変更しない。
- [ ] schema graph cycleをvalidation開始前に拒否する。

### Exit gate

- [ ] Object focused tests pass。
- [ ] scalar regression tests pass。
- [ ] nested pathが`$.user.address.zip`形式で固定されている。
- [ ] field orderとunknown field orderが複数runで同一。
- [ ] release smoke passes。

## 9. M4 — Array and Collection validation

目的: 一次元native arrayとCollectionを同じlogical sequence contractで検証する。

### Tests first

- [ ] zero-based native array
- [ ] one-based native array
- [ ] negative `LBound` native array
- [ ] typed native array
- [ ] Variant array
- [ ] uninitialized dynamic array as empty
- [ ] empty Collection
- [ ] populated Collection
- [ ] typed object array
- [ ] Collection内のNothing、Error Variant、Object要素
- [ ] nested Array/Collection
- [ ] wrong element with correct logical index
- [ ] multidimensional array returns `invalid_array_rank`
- [ ] Dictionary is not treated as array
- [ ] Min/Max/Length boundaries
- [ ] root and nested array paths

### Implementation

- [ ] DG-006を解決する。
- [ ] `ArrayOf(ItemSchema)`を実装する。
- [ ] Nothing ItemSchemaをprogrammer misuseにする。
- [ ] native array detection helperを実装する。
- [ ] safe initialized/rank/bounds helpersを実装する。
- [ ] Collection iterationを実装する。
- [ ] actual indexを0-based logical ordinalへ正規化する。
- [ ] length constraintsをTextと共通contractで実装する。
- [ ] input array/Collectionを変更しない。

### Exit gate

- [ ] Array/Collection focused tests pass。
- [ ] Object/scalar regression tests pass。
- [ ] uninitialized arrayでruntime errorが漏れない。
- [ ] multidimensional arrayがlibrary errorではなくvalidation issueになる。
- [ ] release smoke passes。

## 10. M5 — Literal, Enum, Pattern, and Email

### Literal and Enum tests

- [ ] same category and same value success
- [ ] String `"1"` vs Number `1`
- [ ] Number subtype equivalence
- [ ] Boolean and Number separation
- [ ] Date and Number separation
- [ ] Null and Empty separation
- [ ] Error/Object/Array definition rejection
- [ ] empty Enum rejection
- [ ] Enum definition mutation after construction does not alter schema, or chosen ownership rule is documented and tested
- [ ] duplicate Enum candidates follow a documented rule
- [ ] binary String comparison

### Literal and Enum implementation

- [ ] DG-012のEnum ownership部分を解決する。
- [ ] `Literal`を実装する。
- [ ] `EnumOf`を実装する。
- [ ] definition valuesをvalidation/normalizationする。
- [ ] scalar category comparison helperを実装する。
- [ ] exact Number comparisonをDG-002のcontractへ合わせる。

### Pattern and Email tests

- [ ] Pattern match/non-match
- [ ] invalid expression programmer misuse
- [ ] RegExp runtime unavailable environment error
- [ ] Email accepted examples
- [ ] Email rejected examples
- [ ] empty/local-only/multiple-`@`/whitespace cases
- [ ] Unicode and maximum length cases according to DG-007
- [ ] locale-independent behavior

### Pattern and Email implementation

- [ ] DG-007を解決する。
- [ ] lazy RegExp creation/cacheを実装する。
- [ ] invalid patternとruntime unavailableを区別する。
- [ ] `Pattern`を実装する。
- [ ] `Email`を実装する。
- [ ] RFC完全準拠を目標にしない範囲をREADMEへ記載する。

### Exit gate

- [ ] Literal/Enum/Pattern/Email focused tests pass。
- [ ] full regression tests pass。
- [ ] environment errorがvalidation issueへ変換されない。
- [ ] release smoke passes。

## 11. M6 — Union

### Tests first

- [ ] first branch success
- [ ] later branch success
- [ ] all branches fail
- [ ] nested Union
- [ ] empty Union is programmer misuse
- [ ] Nothing/non-VSchema branch is programmer misuse
- [ ] branch order is deterministic
- [ ] external Issue is one `invalid_union` at failure path
- [ ] branch internal errors do not leak into public Issues
- [ ] environment/programmer errors inside branch propagate instead of becoming union mismatch
- [ ] input Collection mutation after construction follows documented ownership rule

### Implementation

- [ ] DG-012のUnion ownership部分を解決する。
- [ ] `UnionOf(ByVal Schemas As Collection)`を実装する。
- [ ] schema Collectionをvalidation/normalizeする。
- [ ] validation failureとthrown Errを分離する。
- [ ] temporary branch issuesを外部Resultから隔離する。
- [ ] 全branch failure時に`invalid_union`を1件だけ生成する。

### Exit gate

- [ ] Union focused tests pass。
- [ ] full regression tests pass。
- [ ] nested error path and issue order are deterministic。
- [ ] release smoke passes。

## 12. M7 — Hardening, documentation, CI, and release

### Full quality pass

- [ ] 全production procedureを責務別private procedureへ分割する。
- [ ] 巨大procedure、広範囲`On Error Resume Next`、暗黙Variantを除去する。
- [ ] Windows API、unqualified Excel reference、Select/Activateがないことを検索する。
- [ ] optional xlflow dataflow rulesを有効化する価値を評価する。
- [ ] `rtk xlflow metrics --json`でstatic complexity hotspotを確認する。
- [ ] dead codeとunused private procedureを確認する。
- [ ] public API compile fixtureを全signatureと照合する。

### Performance

- [ ] runtime benchmark専用の`src/modules/Benchmarks/ValidationBenchmarks.bas`を追加し、release artifactから除外する。
- [ ] `task benchmark`を実装し、同一fixtureを専用managed Excel sessionで実行できるようにする。
- [ ] `task benchmark`自身が開始時のsession ownershipを確認し、専用sessionのstart、計測、stop/cleanupを所有する。user-owned workbookへattach中は変更せず明示的に失敗する。
- [ ] benchmark結果を`artifacts/benchmarks/<timestamp>-windows-x64.json`へ出力する。
- [ ] 比較用baselineの形式とtracked保存場所を決める。directoryを追加した場合は`AGENTS.md`のtreeも更新する。
- [ ] 1,000 scalar fields fixtureを作成する。
- [ ] 10,000 scalar validations fixtureを作成する。
- [ ] success-heavy fixtureとissue-generation fixtureを分ける。
- [ ] warm-up後5回のmedianを記録する。
- [ ] 1,000 fields < 500 msを検証する。
- [ ] 10,000 validations < 1,000 msを検証する。
- [ ] baselineに対する25%超の悪化がないことを確認する。
- [ ] base/headで同じfixture、iteration、metric setを使う。
- [ ] elapsed timeに加えvalidation count、issue count、RegExp生成回数などのdeterministic counterを記録する。
- [ ] machine、Office bitness、Excel version、commandを記録する。

`rtk xlflow metrics --json`はstatic complexityの証拠であり、runtime performanceの合否判定には使わない。runtime targetは`task benchmark`のJSONで判定し、fixture ID、warm-up回数、計測5回のraw値とmedian、validation count、issue count、RegExp生成回数、環境情報を保存する。

performance targetを変更してreleaseする場合は、単なるProgress Log記載では完了にならない。specまたはADRを先に更新し、変更理由、利用者への影響、変更前後を同一条件で比較したbenchmark JSONを残し、release gate承認者と承認日をProgress Logへ記録する。

### Documentation

- [ ] READMEに「3 files / no reference setup」を記載する。
- [ ] READMEにinstallation手順を追加する。
- [ ] READMEにcanonical exampleを追加する。
- [ ] READMEにPublic API overviewを追加する。
- [ ] READMEにerror handlingとIssue例を追加する。
- [ ] READMEにmutable builderのaliasing warningを追加する。
- [ ] READMEにPattern/Emailの非RFC完全準拠を追加する。
- [ ] READMEにWindows 64-bit verified、macOS/32-bit unverified/non-guaranteedを明記する。
- [ ] error code一覧を公開する。
- [ ] path grammarを公開する。
- [ ] release/update/remove手順を追加する。
- [ ] examplesをVBE compile fixtureで検証する。
- [ ] design/spec/ADRと実装のdrift auditを行う。

### CI and automation

- [ ] Excel不要のlint/analyze/test discovery/format checkをCIへ追加する。
- [ ] behavioral VBA testsはWindows + Excelが必要なgateとして分離する。
- [ ] Excel/VBE compile checkをlocalまたはself-hosted gateとして定義する。
- [ ] CI status名から検証範囲が分かるようにする。
- [ ] release staging/verification TaskをCIまたはrelease手順へ接続する。
- [ ] CI failureと未実行を区別する。

### Release artifact

- [ ] DG-008の決定どおりartifactを生成する。
- [ ] `Schema.bas`、`VSchema.cls`、`VValidationResult.cls`以外が含まれないことを確認する。
- [ ] fresh workbookへ手動importする。
- [ ] external referenceが追加されていないことを確認する。
- [ ] VBE compileする。
- [ ] canonical exampleを実行する。
- [ ] source artifactと検証workbookのchecksumを記録する。
- [ ] source artifactのencoding、改行コード、`Attribute VB_Name`、class attributesを確認する。
- [ ] LICENSEがrelease pageまたはrepositoryから取得可能であることを確認する。現行specの3-file import payloadへは含めない。
- [ ] THIRD_PARTY_NOTICESの要否を最終確認する。
- [ ] DG-008で採用した場合は`CHANGELOG.md`、release note、version metadata、ZIP、sample workbookを作成・更新する。
- [ ] DG-008で採用した追加物は3-file import payloadの外側で生成・検証する。同一artifactへ含める場合は、先にspec/ADR更新が完了していることをgateにする。

### Compatibility-conscious static audit

- [ ] Windows APIがない。
- [ ] pointer-size依存宣言がない。
- [ ] LongLong型名へcompile-time依存していない。
- [ ] Excel Object Modelへコア依存していない。
- [ ] path/string handlingがhost path separatorへ依存していない。
- [ ] unsupported runtime component failureが明示的である。
- [ ] macOS/32-bitを検証済みと誤記していない。

### Exit gate

- [ ] full tests pass。
- [ ] lint/analyze pass。
- [ ] VBE compile pass on Windows 64-bit。
- [ ] performance targets pass、またはspec/ADR・比較benchmark・release gate承認を伴うtarget改定が完了している。
- [ ] release artifact verification pass。
- [ ] README/spec/ADR/design driftなし。

## 13. M8 — Independent final review and completion

### Review

- [ ] 実装者とは別のreviewerがread-only reviewを行う。
- [ ] 利用可能なら`orca-supervised-final-review` skillで独立worktree reviewを実行する。
- [ ] review scopeにproduction 3 files、tests、release helper、docs、CIを含める。
- [ ] confirmed findingとunsupported observationを区別する。
- [ ] valid/in-scope findingだけを最小修正する。
- [ ] 修正ごとにfocused regressionを追加する。
- [ ] 修正後にfull verificationを再実行する。

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

### v1 completion gate

- [ ] `docs/design.md`のDefinition of Doneを全項目確認した。
- [ ] 本ロードマップのM0-M8 exit gateをすべて満たした。
- [ ] unresolved High/Medium review findingがない。
- [ ] unverified platformを正確に表示している。
- [ ] clean checkoutからverificationを再現できる。
- [ ] release artifactを生成できる。
- [ ] final commit SHAとartifact checksumをProgress Logへ記録した。

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
