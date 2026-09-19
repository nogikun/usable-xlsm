# usable-xlsm VBA開発ツール ロードマップ仕様

## 目的

`usable-xlsm`を、特定の電卓サンプルを編集する道具ではなく、信頼済みの
Microsoft 365デスクトップ版Excel（Windows/macOS）でVBAとワークブック
オブジェクトを短い反復で開発・検証・リリースするための汎用スキルにする。

「Excelで動かせるすべて」は、未検証の機能まで成功扱いする意味にはしない。
対象を次の4状態で明示し、各操作で状態を返す。実装していない将来機能を
`supported`と呼ばない。現状未対応の機能は`planned`として仕様書にだけ置き、
inventory・保存後比較・失敗時の原本不変が実装されるまで昇格させない。

- `supported`: このスキルが変更と保存後検証まで保証する
- `preserve-only`: 読み取り・棚卸し・保存時の保持は保証するが、編集はしない
- `unsupported`: 入力検出時に理由を返し、黙って変更しない
- `planned`: 将来候補。現行コマンドでは受け付けない

## 調査から得た設計制約

### xlsxスキルから継承する短縮原則

`xlsx`スキルは、操作をファイル種別ではなく作業内容に分け、既存ファイルの
規約を優先し、スクリプトで反復編集し、保存後に再計算・エラー検査を行う。
`usable-xlsm`も同じく「1回の宣言的スクリプト」「入力検証をExcel起動前に
完了」「同一Excelジョブで関連変更をまとめる」「保存後の値・構造検証」を
基本にする。ただし、VBAは`openpyxl`のようなパッケージ編集だけでは新規
ソースを注入できず、ExcelのVBE/COM操作と実行時検証が必要である。

### Excelの公式オブジェクト境界

MicrosoftのVisual Basic Add-In Modelでは、`VBProject`が`VBComponents`と
`References`を持ち、`CodeModule`がコンポーネントのコードを行単位で追加・
削除・編集する。このため、標準モジュール、クラス、既存のドキュメント
モジュールは共通のソースワークフローに載せられる。

一方、シート上のForm Controlは`Shapes.AddFormControl`、ActiveX/OLEは
`OLEObjects.Add`という別経路である。UserFormはデザイナーとコントロール
コレクションを持ち、`.frm`テキストだけでなく`.frx`バイナリとの組を壊さず
扱う必要がある。したがって「ボタン」は単一機能ではなく、Form Control、
ActiveX、UserForm内コントロールを別 capability として扱う。

Trust Centerの「Trust access to the VBA project object model」が無効なら、
VBAソース編集はできない。署名済みプロジェクトはコード変更で署名が無効に
なるため、編集後は未署名候補として扱い、署名は別のリリース工程に分離する。

### Windows/macOSの実装境界

VBA自体はExcel for WindowsとExcel for Macの両方で利用できるが、COM Add-inは
Macでは利用できず、ActiveX controlsもMacではサポートされない。Mac固有の
VBAからAppleScriptを呼ぶ`AppleScriptTask`はあるが、これはVBAの実行機能で
あって、WindowsのCOM/VBE自動化の代替ではない。

したがって、同じmanifestを両OSで検証できるようにするが、Excel workerの
backend capabilityは分ける。

| backend | 自動化する範囲 | 失敗時の扱い |
|---|---|---|
| `windows-excel-com` | VBE import/export、worksheet/cell、Form Control、run/test、close/reopen検証 | workerが所有するExcelだけを終了し、失敗はpromotion拒否 |
| `macos-excel-manual` | manifest/sourceの検証、差分計画、Mac用の手順出力 | `pending_manual`を返す。VBA注入・実行を成功扱いしない。SHA-256付き証跡がない限り未完了 |
| `macos-excel-automation` | 将来候補。AppleScript/GUI自動化を想定しない | 実Macと実M365で検証できるまで`planned` |

Macの完全自動VBA注入は、Windows COM経路の名前だけを置き換えて実装しない。
実Mac上のExcel、VBE、Trust Center、署名、ファイル権限を確認できるworkerを
用意し、同じfixtureの実測を通過した場合だけ自動backendを昇格する。

Macの手動証跡は次のJSONを成果物と同じ作業記録へ保存し、対象成果物のSHA-256と
manifest/sourceのSHA-256が一致する場合だけ`macos_manual_passed`とする。

```json
{
  "workbook_sha256": "<candidate sha256>",
  "manifest_sha256": "<manifest sha256>",
  "source_sha256": "<source tree manifest sha256>",
  "platform": "macos",
  "excel_version": "Microsoft 365",
  "operator": "<operator id>",
  "checks": ["open", "VBA import", "smoke", "save", "reopen"],
  "result": "passed"
}
```

### GitHub上の既存設計から学ぶこと

Rubberduckは、VBAのソース同期だけでなく、パーサー、Code Explorer、静的
解析、リファクタリング、テスト実行までを別機能として持つ。これは開発者体験
の目標にはなるが、同時に大きなCOMアドインであり、巨大プロジェクトでは
パースや解析の負荷もある。`usable-xlsm`はRubberduckを再実装せず、まず
Excelのネイティブな適用・抽出・実行・保存検証を最短経路にする。

## Capability matrix

| 対象 | 現状 | 方針 | 完了条件 |
|---|---|---|---|
| 標準`.bas` | supported | 追加・置換・削除・抽出 | Excel再抽出と名前/ソース一致 |
| クラス`.cls` | supported | 追加・置換・削除。ただしドキュメントモジュールは更新のみ | 型とソースの保存後検証 |
| `Sheet*.cls`/`ThisWorkbook.cls` | supported | 既存コードの更新のみ | イベントコードを保持して再抽出 |
| UserForm`.frm`/`.frx` | planned | Windowsでの対保持を検証できるまで編集しない。Macはデザイナー編集を前提にしない | 実fixtureの`.frm`/`.frx`対と部品ハッシュを確認 |
| Form Control | supported | 作成・置換・caption/geometry/macro設定。単独削除はplanned | 閉じたOOXML/VMLとExcel読戻しを確認 |
| ActiveX/OLE | planned | Windows専用候補。Macではunsupported。追加/編集は明示的finalizerへ | `xl/activeX`等の部品差分を検査 |
| ワークシート/セル/式/形式 | supported | 宣言的に同一Excelジョブで適用 | 保存後の値・式・表示形式・可視性を確認 |
| 名前定義/テーブル/チャート/図形 | planned | 棚卸し・差分・保持を実装するまで編集対象ブックを拒否 | 未許可差分がないことを確認 |
| VBA References | planned | broken reference検出と一覧化を先に実装する | `IsBroken`相当の診断を返す |
| Custom UI/Ribbon | planned | `customUI`部品を保持・差分表示するまで編集しない | XML部品の無断消失を検出 |
| XLMマクロ | unsupported | 実行・編集しない。既存TrustPolicyの拒否を緩和しない | 入力検出またはscan不能でExcel起動前に拒否 |
| 署名 | preserve-only | 入力署名を検出したら編集を既定拒否。明示承認時もcandidate止まりとし、外部再署名と一致SHA検証後だけrelease可能 | 署名検証に失敗した成果物をpromotionしない |
| `.xlsb` | planned | 現行の受入契約を整理するまで新機能を追加しない | `.xlsm`経路に誤送信しない |
| マクロ実行/テスト | supported | 単一smoke、選択Test、release全件 | 実Excelの終了コード/結果セルを記録 |

## 最短経路の設計

新しい機能ごとに専用CLIを増やさず、既存の`apply`を一つの変更トランザク
ションとして拡張する。OSに依存しないmanifest/source/checkと、OS依存の
Excel workerを分離する。実装順は次の通りとする。

1. **検出**: まず既存`preflight`/`extract`に、現在検出できるVBAコンポーネント、
   シート、セル、Form Control、OS/backendを返す。未実装の部品を「保持確認済み」
   と表示しない。汎用`inspect`は出力schemaを固定してから追加する。
2. **ソース**: 現行の`modules`、`remove_modules`、`sync`を汎用VBA契約として
   維持し、標準モジュールと編集可能クラスを最初のsupported範囲にする。
3. **オブジェクト**: 現行の`worksheets`、`cells`、`buttons`を一つのExcel
   ジョブで処理する。Form Controlは`buttons`として明示し、ActiveXを同じ
   JSONに偽装しない。
4. **実行**: `run --macro`を最小の実行確認とし、必要なときだけ
   `test --filter`、release時だけisolated全件テストを使う。
5. **保持境界**: 実際に比較できる部品だけを`preserve-only`と呼び、編集対象外の
   部品が検出されたが比較不能なら、backendを問わずExcel起動前に拒否する。
6. **高度機能**: UserForm、ActiveX、References、Names/Charts、Custom UI、
   署名は、各々の保存後検証を先に定義できたものだけ個別に`planned`から昇格する。

### capability出力の最小schema

操作結果には次の形で対象ごとの状態を返す。複数状態が混在するブックでは
`overall`を最も厳しい状態（`unsupported` > `planned` > `preserve-only` >
`supported`）にし、各itemの`reason_code`を必ず残す。

```json
{
  "platform": "windows",
  "backend": "windows-excel-com",
  "overall": "planned",
  "items": [
    {"kind": "vba.standard_module", "id": "Module1", "status": "supported", "reason_code": null},
    {"kind": "activex", "id": "Sheet1.OLEObject1", "status": "planned", "reason_code": "ACTIVEX_EDIT_NOT_IMPLEMENTED"}
  ]
}
```

上の例は`planned`項目を含むため、実際の`overall`は`planned`でなければならない。
このschemaを返せない機能は、成功ではなく`planned`または`unsupported`として
拒否する。未回答だった設計事項は次の既定値で固定する。

- 対象: Windows/macOSのMicrosoft 365デスクトップ版
- 外部依存: 追加しない。Windowsは既存Python＋Excel COM、Macは手順出力のみ
- UserForm/ActiveX: Phase 1では編集しない。ActiveXはMacでunsupported
- 開発検証: `apply --no-test`＋期待値付きsmoke
- release: isolated testと署名検証を必須にする

### 原本・candidate・promotionの契約

- 入力ブックはExcelで開かず、最初にSHA-256を記録する。
- Excelは作業コピーだけを開く。元ファイルをExcelで開いた場合は失敗とする。
- save/close/reopen、VBA再抽出、対象オブジェクト、許可済みパッケージ差分、
  runtime smoke、監査の全てがcandidate上で成功するまで元ファイルを置換しない。
- 失敗時はcandidateとscratchを削除または隔離し、元ファイルのSHA-256が開始時と
  一致することを確認して失敗にする。rollback不能ならpromotionしない。
- promotion直前に`ready_to_promote`を記録し、promotion後の監査失敗を成功にしない。

### パッケージ保存の初期境界

Phase 1で扱えるのは、preflightが検出し、保存前後を比較できる部品だけとする。
検出できないOOXML relationship、VBA、extension、UserForm、ActiveX、Custom UI、
external link、Names/Charts/Shapesがある場合は、対象機能を編集せずExcel起動前に
`UNVERIFIED_PACKAGE_PART`で拒否する。Phase 2のinventoryと正規化比較が実装される
まで、これらを`preserve-only`とは呼ばない。

## 宣言スクリプト契約

既存のトップレベルキーを壊さず、最初は次の5キーだけを安定契約とする。

```json
{
  "modules": ["vba/Module1.bas"],
  "remove_modules": [],
  "worksheets": [{"name": "Main", "create": true, "visible": "visible"}],
  "cells": [{"sheet": "Main", "address": "A1", "value": "ready"}],
  "buttons": [{
    "sheet": "Main", "name": "btnRun", "caption": "Run",
    "macro": "Module1.Run", "left": 12, "top": 24,
    "width": 120, "height": 24, "replace": true
  }]
}
```

未知キーはExcel起動前に拒否する。新しい対象を追加するときは、キー追加の
前に「競合規則」「dry-run/inventory表示」「保存後検証」「失敗時の原本不変」
「p50時間」を仕様化する。高度機能を一つの`objects`汎用辞書に押し込まない。

必須キーは配列として存在する場合だけ処理し、各entryの必須型・範囲をExcel起動
前に検証する。同一casefold名のmodule/sheet/button、同一sheet/addressのcell、
`modules`と`remove_modules`の交差、`replace: false`での既存button、無効なA1、
負または0のgeometryは拒否する。処理順はvalidate → stage → modules → worksheets/
cells → buttons → save/close/reopen verificationで固定する。

Form ControlのgeometryはExcel points、左上は対象worksheetのA1を基準とし、
`replace: true`は同じsheet/nameだけを削除して同一設定で作り直す。単独削除は
別manifestを追加するまで受け付けない。

## 開発とリリースの検証予算

開発中も実際のExcel適用は毎回行うが、毎回すべてのテストを起動しない。

| 場面 | 1コマンド経路 | 証明すること |
|---|---|---|
| 反復編集 | `apply --no-test`を使い捨てコピーへ | Excel import/save/close、再抽出、対象オブジェクトとパッケージ保持 |
| 挙動変更 | 上記後に`run --macro Module.Procedure`を1回 | 変更した実行経路の実Excel挙動 |
| テスト追加 | `test --filter Name --shared-copy` | 選んだTestだけを共有コピーで実行 |
| リリース | 既定のisolatedテスト | 各Testを新しいコピーで実行し、監査・原本ハッシュ・atomic replace |

ベンチマークは同じブック・スクリプト・Excel状態で3回以上測り、p50の壁時計、
Excelジョブ数、コマンド数、成功率を別々に記録する。目標は、機能追加後も
通常反復がExcelジョブ1回、CLI 1回であり、releaseの安全な検証を削らずに
開発だけを短縮することとする。

`isolated`はworkbook copyの分離を意味し、任意VBAのファイル、ネットワーク、
COM、メール、レジストリ等のOS副作用をsandboxするものではない。したがって、
外部副作用を含むテストはtrusted dedicated workerでのみ実行し、untrusted workbook
には実行経路を提供しない。テスト結果に`side_effect_policy: trusted_worker`を記録する。

promotion順序は、(1)入力hash、(2)candidate作成、(3)Excel適用、(4)close/reopen・
再抽出・package比較・runtime検証、(5)監査`ready_to_promote`、(6)原本hash再確認、
(7)atomic replaceの順に固定する。署名検証が必要な成果物は外部再署名後の別工程で
この順序を再実行し、検証前のcandidateをrelease名へ置換しない。

## 段階計画

### Phase 0: 契約を固定する

- `SKILL.md`のcapability matrixと開発/リリース経路を正とする。
- `preflight`/`extract`の出力にplatform/backend/capabilityを含める。
- `inspect`はschemaと失敗条件を実装できる段階で追加する。
- `.xlsb`は現行CLIの受入を直ちに破壊せず、`legacy-unverified`として結果へ明示し、
  新規のsupported保証対象には含めない。新しい安全ゲートを実装した後にplannedへ
  移行する。
- 既存の`apply` JSONと電卓fixtureを汎用サンプルとして維持する。

### Phase 1: 保存安全境界を先に実装する

- 原本hash、作業copy、candidate、cleanup、rollback不能時のpromotion拒否をテストする。
- preflightで未比較のpackage partを拒否し、XLM検出またはscan failureはTrustPolicyの
  `allow_xlm`/`allow_scan_failures`に関係なく拒否する。
- 標準モジュール、クラス、ドキュメントモジュール、Form Control、
  worksheet/cellを対象に、抽出・適用・閉包後検証を揃える。
- `verify_saved_workbook_objects`でnumber_formatも再読込後に比較する。
- signed inputは既定拒否し、明示的なsignature-removal承認があってもcandidate止まり
  とする。外部再署名後、candidate SHAと署名検証結果を照合できた場合だけreleaseする。
- `run --macro`はruntime error、modal dialog、timeout、Excel crashを失敗にする。
  期待値なしの結果は`executed`、期待値一致だけを`passed`、不一致を`failed`とする。
- `run --macro`は期待値を指定しない限り`executed`であり、`passed`とは呼ばない。
  期待値を指定した場合だけ不一致を失敗にする。
- 期待値は最初は`--expect-cell Sheet!A1=15`のscalar比較だけに限定し、型と文字列表現を
  JSON結果へ記録する。
- WindowsではWorkbookイベント、XLM、外部リンク更新を実行経路から除外できない
  場合は、Excel起動前に拒否する。
- Macでは自動適用を成功扱いせず、source/manifest/check結果と手動適用手順を返す。
- 代表的なfixtureを「追加」「置換」「削除」「イベント更新」「ボタン割当」に
  分け、各々の最短コマンドとp50を記録する。

### Phase 2: preserve-onlyの可視化

- UserForm/FRX、ActiveX/OLE、Custom UI、XLM、署名、外部リンク、References、
  Names/Charts/Shapesをinventoryと差分報告に載せる。
- 部品ごとのID、比較対象、XML/binary正規化、Excelが変更するvolatile fieldの許容値、
  ハッシュ方式、検査不能時の拒否条件を仕様化する。
- 変更対象でない部品の消失・変更を検出し、promotionを止める。
- これにより、未対応機能を「たぶん保持される」と誤認させない。

### Phase 3: 個別機能の昇格

利用例が実際にある順に、1機能ずつ実装する。候補はReferences、Names/Charts、
Form Controlの拡張、UserForm、ActiveX、Custom UIの順とする。各機能は専用の
manifest section、live verification、closed-package verification、runtime smoke
を持ち、合格しない限り`preserve-only`のままにする。

### Phase 4: IDE補助は必要性が出たときだけ

Code Explorer、静的解析、リファクタリング、VBE add-in、LSP、UserForm designer
は別プロダクト級の範囲である。Rubberduck相当を再実装せず、外部ツールとの
source syncや既存VBEを呼び出す統合で足りるかを先に検証する。

## 完了条件

- 任意の対象を「supported/preserve-only/unsupported/planned」のいずれかでinventory
  できる。未実装対象をsupportedとして報告しない。
- supported対象の変更は1回の宣言`apply`で、実Excelの1ジョブとして完了する。
- source、ワークブックオブジェクト、Form Control、保存済みパッケージの各検証が
  失敗時に原本を変更しない。
- runtime変更は最小smoke、リリースはisolated suiteで検証される。
- fixtureだけでなく、標準モジュール、イベント、クラス、署名、XLM、壊れた参照、
  ダイアログ、タイムアウト、原本hash不変、candidate cleanupの代表ケースを測定する。
- `planned`対象が残っていても、それを明示した上でPhase 1のsupported範囲だけを完了
  と判定できる。未実装対象をsupportedと報告した場合は失敗とする。
- 既存の開発ベンチマークと同じ条件で、機能別にp50とExcelジョブ数を比較できる。

## 確定した意思決定

- 対象はWindows/macOSのMicrosoft 365デスクトップ版。
- UserForm/ActiveXは「できる範囲を調査する」が、Phase 1の必須機能にはしない。
  実Macでの自動化検証ができない間はMac側を手動手順または`planned`とする。
- 外部依存は追加せず、Windowsは既存Python＋Excel COMを使う。
- 開発時の完了判定は`apply --no-test`＋期待値付きsmoke、全件テストはreleaseのみ。
- 実装順は、capability schema、preserve差分、runtime期待値、Names/Charts、
  UserForm/ActiveXの順とする。

## 参照資料

- Microsoft, [Objects (Visual Basic Add-In Model)](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/objects-visual-basic-add-in-model)
- Microsoft, [Shapes.AddFormControl method](https://learn.microsoft.com/en-us/office/vba/api/Excel.shapes.addformcontrol)
- Microsoft, [OLEObjects.Add method](https://learn.microsoft.com/en-us/office/vba/api/excel.oleobjects.add)
- Microsoft Support, [Overview of forms, Form controls, and ActiveX controls](https://support.microsoft.com/en-US/Excel/overview-of-forms-form-controls-and-activex-controls-on-a-worksheet)
- Microsoft Support, [Digitally sign your VBA macro project](https://support.microsoft.com/en-us/office/vba-digitally-sign-your-vba-macro-project)
- Microsoft Learn, [VBA access to create/open a VSTO system project](https://learn.microsoft.com/en-us/visualstudio/vsto/enable-access-to-vba-to-create-or-open-a-visual-studio-tools-for-office-system-project)
- Microsoft Learn, [Macros from the internet are blocked by default](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked)
- Microsoft Learn, [Office for Mac VBA](https://learn.microsoft.com/en-us/office/vba/api/overview/office-mac)
- Microsoft Learn, [Differences between Office Scripts and VBA macros](https://learn.microsoft.com/en-us/office/dev/scripts/resources/vba-differences)
- Microsoft Learn, [Run an AppleScript with VB](https://learn.microsoft.com/en-gb/office/vba/office-mac/applescripttask)
- Microsoft Support, [Assign a macro to a Form or a Control button](https://support.microsoft.com/en-us/excel/assign-a-macro-to-a-form-or-a-control-button)
- GitHub, [rubberduck-vba/Rubberduck](https://github.com/rubberduck-vba/Rubberduck)
- GitHub Wiki, [Rubberduck Unit Testing](https://github.com/rubberduck-vba/Rubberduck/wiki/Unit-Testing)
- GitHub Wiki, [Rubberduck Features](https://github.com/rubberduck-vba/Rubberduck/wiki/Features)
