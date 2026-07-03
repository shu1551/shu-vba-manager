# 変更履歴

## 2026-07-03

Claude Fable 5 の提供再開（7/1）を受けて、一日で実施した大規模改修。

### 修正（総点検の成果）

- **重大**: プロシージャ置換の Attribute 経路で `End Sub 'コメント` を境界として認識できず、条件が揃うと**隣のプロシージャを巻き込んで消す**バグを修正（再現テストで確認済み）
- **重大**: 未知のコマンドラインオプションを黙って無視していたのを**エラー停止**に（`--contets` のようなタイプミスが「全消し」既定動作に化ける事故の防止）
- **重大**: `save-as` の無確認上書き・未対応拡張子の無言フォールバック・xlsm→xlsx のマクロ喪失無警告を修正
- **重大**: `run-macro` 実行後に Excel の DisplayAlerts が無効のまま残る問題を修正
- **重大**: 「シート名だけ」の範囲指定が使用範囲全域の破壊対象になる問題をガード（`--whole-sheet` で明示時のみ許可）
- replace-module に VB_Name 照合（別モジュール取り違え防止）と Import 失敗時の自動復旧を追加
- .frm/.frx ペアの取り扱い（フォームのレイアウト消失の穴）・一時ファイルの掃除漏れを修正
- 改行二重化（CP932 の \r\r\n）ガードを入口層にも追加（多層防御の復元）ほか多数

### 追加（マクロ管理）

- `grep` — 全マクロ横断のコード検索（モジュール／プロシージャ／行番号つき）
- `code-replace` — 一括置換（diff プレビュー → 確認 → バックアップ → 変更行のみ書換。ショートカット定義を壊さない）
- `add-procedure` / `delete-procedure` — プロシージャの追加・削除の軽量経路
- `list-backups` / `restore` — 自動バックアップの一覧と復元（undo）
- `batch` — コマンド列を COM 接続1回で一括実行（実測: Excel 起動込み6コマンド約1秒）
- `export-all` — 全モジュール一括エクスポート（`--check` で取り込み前検査つき）
- `setup-check` — 導入セルフ診断（Python / pywin32 / Excel / VBOM 信頼設定を○×表示）
- `export-pdf` — PDF 出力（ブック／シート／範囲）
- `read-range --tsv` — 読んだ範囲を TSV に書き出し、編集して `write-range` で書き戻す往復
- `list --detail`・`sheet-info --preview`・`table read`・`get` の複数取得・タイポ時の近似候補提示 ほか

### 追加（UserForm）

- **`form_layout.py`（新規）** — 宣言的レイアウトエンジン。行構造を書くだけで整列済みフォームを構築。
  Excel 不要の配置プレビュー（`preview`）、イベント雛形の自動生成（`vba_stub`）、必須項目チェック、
  Frame / MultiPage（タブ）/ 範囲選択欄 / スピン付き数値 / 画像に対応
- **`form_tool.py`（新規）** — 幾何操作 CLI（scale / set / move / align / size-match / distribute /
  tab-order / rename-control / delete-control / copy-form）。破壊的操作は事前に自動退避
- **`form_inspect.py`（強化）** — 実表示 PNG 撮影（`--png` / 全タブ `--png-all` / 名前描き込み `--names`）、
  機械検査 `--lint`（重なり・はみ出し・不揃い・タブ順・孤児イベントハンドラ）、
  既存フォームの宣言コード逆変換 `--to-layout`（往復で同一レイアウトを確認済み）

### テスト

- `test_tools.py`（新規） — pytest による自動テスト22件。レイアウト計算・エンコーディングガード・
  lint・逆変換など COM 不要部分を検証（今回の修正の回帰テストを含む）

## 2026-06-13

- Excel MCP（Stefan Broenner さん）を参考に「目」（read-range / sheet-info / screenshot）と
  「手」（write-range / format-range / 行列操作 / 検索置換 / 印刷設定）を追加
- グラフ／ピボット／スライサー／PowerQuery／データモデル(DAX) 操作を追加（9 → 約50コマンド）

## 2026-05

- 初回公開（vba_manager.py / form_builder.py ほか）
