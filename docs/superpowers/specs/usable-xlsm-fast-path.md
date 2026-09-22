# usable-xlsm 最短適用経路仕様

## 目的

既存の `.xlsm` に対するVBAモジュール追加・編集とForm Controlボタン配置を、
手作業のExcel操作や個別のインストーラーマクロなしで、1回の宣言的な
`apply`にまとめる。

## 適用契約

`apply` JSONのトップレベルで許可するキーは次の5つとする。

- `modules`: JSONファイルから見た `.bas`/`.cls` の相対パス
- `remove_modules`: 削除する標準 `.bas`/`.cls` のファイル名
- `worksheets`: シート作成・可視性の宣言
- `cells`: 単一A1セルの値・数式・表示形式
- `buttons`: 既存シートに置くForm Controlの宣言

未知キーは無視せず、Excel起動前にエラーにする。

`worksheets`は既存シートを再利用し、存在しない場合は`create: true`の
ときだけ作成する。`visible`は`visible`、`hidden`、`veryHidden`に限定する。
`cells`は`value`または`formula`のちょうど一方を持ち、`value: null`で
クリアできる。既存値は標準で上書きし、`overwrite: false`なら非空セルで
失敗する。同一シートの同一セルや同名シートの重複は入力エラーとする。

## 処理契約

1. 静的preflight、信頼認可、入力検証をExcel起動前に行う。
2. 既存VBAを抽出し、指定モジュールだけをオーバーレイする。
3. ステージングコピーを1回だけExcelワーカーで開く。
4. VBA、worksheet、cell、buttonを同じジョブ内で適用して保存する。
5. ワーカー終了後、閉じた保存済みパッケージからシート、セル、ボタンを再検証する。
6. `--no-test`は使い捨て開発コピー専用とし、リリースはテストを有効にする。
7. 失敗時は元ファイルを変更せず、既存のバックアップ・監査契約を維持する。

## 性能の測定契約

- 同じ入力ブック、同じスクリプト、同じWindows/Excel状態で比較する。
- cold/warmの区別を記録し、各3回以上のp50を採用する。
- 壁時計時間はCLI開始からJSON結果受領までとする。
- Excelジョブ数は適用ジョブとテストジョブを分けてJSONに出す。
- `test_mode`と`excel_jobs`を結果JSONに含める。
- コマンド数削減と壁時計時間短縮を別指標として報告する。
- `skills/usable-xlsm/scripts/benchmark_apply.py`で3回以上のp50を測る。

## 受け入れ条件

- 既存モジュール編集、新規モジュール追加、新規シート作成、セル初期化、ボタン配置を1回の`apply`で完了できる。
- Form Controlボタン、シート、セルが保存後のパッケージ再検証で確認できる。
- 開発経路はExcel適用ジョブ1回である。
- `--post-macro`は高度な信頼済みfinalizerに限定し、通常のボタン配置では使わない。

## 非目標

- UserForm、ActiveX、署名済みVBA、XLMマクロの自動編集
- 任意VBAの自動生成
- `openpyxl`によるVBAソース注入
