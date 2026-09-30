# usable-xlsm

AIエージェントでExcel VBAを開発・修正するためのスキルです。
WindowsのExcelワーカーによる安全な更新・実行・検証を維持しながら、
macOS/Linuxでも抽出、ソース修正、差分確認、静的検証、新規候補の作成ができます。

## インストール

```bash
npx skills add nogikun/usable-xlsm
```

スキル利用時は、最初にブックを利用するOSを Windows / macOS / Linux から選択します。
利用先のOSと、AIエージェントが実際に動いているOSは別々に扱います。

## 開発環境

```text
uv sync --project skills/usable-xlsm --frozen
uv run --project skills/usable-xlsm usable-xlsm environment --target-os windows
uv run --project skills/usable-xlsm usable-xlsm doctor
```

| 環境 | 開発・確認 | マクロ実行・更新 |
| --- | --- | --- |
| Windows | 抽出・差分・構文・新規作成 | 専用のExcelワーカーで自動検証・段階更新 |
| macOS | 同上 | Mac Excelで手動確認、またはWindowsワーカーへ引継ぎ |
| Linux / Excelなし | 同上 | LibreOfficeは任意の互換性確認、Excelの検証はWindowsへ引継ぎ |

WindowsでExcelを使わない作業には `doctor --mode static` を指定できます。
macOS/Linuxでは `doctor` が静的作業の準備を確認し、Excelがないだけで作業を止めません。

## 小さな修正を小さく扱う

```text
uv run --project skills/usable-xlsm usable-xlsm extract --workbook book.xlsm --output work/vba
uv run --project skills/usable-xlsm usable-xlsm plan --source work/vba --workbook book.xlsm
uv run --project skills/usable-xlsm usable-xlsm check --source work/vba/Module1.bas
```

変更した既存モジュールだけのフォルダには `plan --partial` と
Windowsの `update --partial` を使えます。省略したモジュールは保持します。
全文出力は避け、必要なときだけ `plan --diff` を指定します。
最終更新では従来どおりバックアップ、再抽出、全テスト、監査、原子的な置換を行います。
実際の時間・トークン削減率は利用タスクごとに測定してください。

新規作成の `create` はXlsxWriterでシートとボタンを作り、既存のVBAバイナリを埋め込みます。
ソースからの独自バイナリ生成は `--experimental` で明示的に選択できます。
静的検証や構造の再抽出は、Excelのコンパイル・動作保証とは別の結果として表示します。

利用手順は [SKILL.md](skills/usable-xlsm/SKILL.md)、
新規作成は [creation.md](skills/usable-xlsm/references/creation.md)、
OSごとの注意点は [portable-workflow.md](skills/usable-xlsm/references/portable-workflow.md) を参照してください。

## 検証

```text
uv run --project skills/usable-xlsm python -m unittest discover -s skills/usable-xlsm/tests -v
```

GitHub Actionsで Windows / macOS / Linux の静的・回帰テストを実行します。
このCIにはデスクトップExcelがなく、実ExcelのCOM操作やMac/LibreOfficeのマクロ実行は含みません。
