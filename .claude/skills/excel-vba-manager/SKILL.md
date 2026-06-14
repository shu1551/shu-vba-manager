---
name: excel-vba-manager
description: |
  Excel の VBA マクロ・シート・テーブル・ピボット・チャート・パワークエリ等を
  Python(vba_manager.py / form_builder.py)から操作する汎用スキル。
  対象は常に「今アクティブに開いている Excel ブック（ActiveWorkbook）」で、特定のブック・パス・
  モジュール名に依存しない。秀.xlsm はこのツールの利用例の一つにすぎない。

  以下のときに使う：
  - 「マクロを追加/修正/作って」「VBAマクロ」など Excel VBA の作成・修正
  - vba_manager.py / form_builder.py を使う作業
  - 開いている Excel ブックのプロシージャ/モジュールの取得・置換
  - シート操作・セル編集・テーブル・ピボット・チャート・パワークエリ・データモデル
  - UserForm の作成・修正、ショートカットキー(Attribute行)の操作

  ※「秀」「秀.xlsm」「アドイン」「shu001/003/005」と名指しされた秀.xlsm固有の話（パス・
    モジュール地図・フォーム一覧）は shu-addin-manager スキルを併用する。
---

# Excel VBA マネージャー（汎用）

## 対象の原則（最重要・まずこれを当てる）

vba_manager.py / form_builder.py は **アクティブな開いているブック（ActiveWorkbook）** に対して動く。

- 引数なし → 今アクティブな Excel ブックを自動使用。
- 第1引数にパスを渡せばそのファイルを対象にする
  （例: `py vba_manager.py get "C:\…\文書件名簿.xlsm" 簿冊フォーム TextBox1_Change`）。
- だから**任意の Excel ファイルの VBA も、Excel で開いてアクティブにすれば同じ手順で修正できる**。
- 既存コードに `XLSM_PATH`・`shu001`/`shu003` 等の固有名詞があっても「対象」と鵜呑みにしない
  （中身は `get_workbook()` を読めば一目瞭然）。**固有名詞より「対象はアクティブブック」を先に当てる。**
  パス・ブック名を決め打ちしない（決め打ちは亡霊の燃料）。

## ツールの場所

```
SCRIPTS:            C:\Users\shu\Desktop\アプリ\秀 20260113\作業ファイル\project\python_scripts\
vba_manager.py:     SCRIPTS\vba_manager.py（メイン管理ツール）
form_builder.py:    SCRIPTS\form_builder.py（UserForm作成モジュール）
_last_proc.vba:     SCRIPTS\_last_proc.vba（get の出力先、UTF-8）
backups:            作業ファイル\backups\（自動バックアップ格納場所）
```

> SCRIPTS は「ツールが置いてある場所」であって「操作対象のブック」ではない。
> 対象は常にアクティブブック（上の原則）。
> すべてのコマンドは SCRIPTS ディレクトリから実行する。**`py` を使うこと（`python` ではない）。**

> ⚠ Excel が未起動の状態でパス指定すると、ツールが Excel を自動化起動する。
> **COM起動の Excel はアドイン(秀.xlam等)・PERSONAL.XLSB を読み込まない**ため、そのインスタンスを
> 普段使いに流用すると「アドインが効かない」状態になる（2026-06-13 のゾンビExcel事故と同症状）。
> この経路に入ると vba_manager が警告を表示する。作業後はそのExcelを閉じ、普段使いは手動起動した
> Excel で行うこと。

## vba_manager.py コマンド一覧

```bash
cd "C:\Users\shu\Desktop\アプリ\秀 20260113\作業ファイル\project\python_scripts"

# マクロ一覧表示
py vba_manager.py list

# モジュール一覧表示
py vba_manager.py list-modules

# プロシージャのコード取得 → _last_proc.vba に保存
py vba_manager.py get <Sub名>

# モジュール指定してプロシージャ取得（同名プロシージャが複数ある場合に使う）
py vba_manager.py get <モジュール名> <Sub名>
py vba_manager.py get <モジュール名>.<Sub名>   # ドット区切りも可

# _last_proc.vba の内容でプロシージャを置換（バックアップ自動作成）
py vba_manager.py replace-procedure -y
#  ※ -y/--yes で確認プロンプト(y/N)をスキップ。Claude経由の非対話実行では必ず -y を付ける。
#    （-y なしだと input() 待ちで止まる。PowerShellから "y" をパイプで流す方式は失敗することがある）
#    差分(Diff)は -y を付けても表示される。--module <名> で対象モジュールの明示も可。

# モジュール全体を Remove+Import で置換
py vba_manager.py replace-module <モジュール名> <basファイル>

# モジュールを .bas にエクスポート
py vba_manager.py export-module <モジュール名>
```

### 目コマンド（シート状態の読み取り・読み取り専用）

```bash
py vba_manager.py read-range     [excel_file] [range]   # セル値をテキスト格子で読む
py vba_manager.py read-range     A1:D10 --formula        # 計算結果でなく数式(.Formula)を表示
py vba_manager.py read-selection [excel_file] [--formula] # 今選択している範囲を読む（--formulaで数式）
py vba_manager.py sheet-info     [excel_file]           # シート構成・使用範囲の一覧
py vba_manager.py screenshot     [excel_file] [range]   # 範囲を画像(PNG)で書き出す
```

> **数式の修正ループ**: `read-range --formula` で今の式を読む → 直した式を `write-range`
> で書き戻す。読み(.Formula)も書き(.Value)も US規約（英語関数名・カンマ区切り）でそろえてあるので往復が一致する。

### 手コマンド（シートの編集・整形・構造操作）

Excel MCP 相当の編集機能を「**開いたままのアクティブブック**」に対して COM で直接行う。
Excel MCP はファイルを閉じる必要があるが、これらは開いたまま使える代わりに
**プログラム経由の変更は Excel の Undo 履歴を消す**。
**既定では保存しない**（Excelで確認後に手動保存、または保存せず閉じれば変更を破棄できる＝Undo代わり）。

```bash
# 値・数式を書き込む（'='始まりは数式）。単一値はインライン、グリッドは TSV。
py vba_manager.py write-range A1 "こんにちは"
py vba_manager.py write-range C1 "=SUM(A1:A10)"
py vba_manager.py write-range A3 --tsv data.tsv     # 省略時は _last_values.tsv（タブ区切り）

# 範囲をクリア（既定すべて／--contents 値のみ／--formats 書式のみ）
py vba_manager.py clear-range A1:D10 --contents

# 書式・整形（複数オプション同時指定可）
py vba_manager.py format-range A1:D1 --bold --bg "#FFFF00" --color "#FF0000" --align center --border thin
#  指定可: --font 名 --size N --bold --unbold --italic --color/--bg '#RRGGBB'
#          --number-format 書式 --align left|center|right --valign top|center|bottom
#          --wrap --border thin|medium|thick|hairline|none --col-width N --row-height N
#          --merge --unmerge --autofit

# シート操作
py vba_manager.py sheet add 新シート [--before 既存 | --after 既存]
py vba_manager.py sheet rename 旧 新
py vba_manager.py sheet copy 元 [新名]
py vba_manager.py sheet delete 名 / activate 名 / show 名 / hide 名 / very-hide 名
py vba_manager.py sheet visibility 名                         # 表示状態を表示
py vba_manager.py sheet tab-color 名 "#FF8800"               # タブ色（#RRGGBB / R G B / --clear / 引数なしで取得）

# テーブル(ListObject)
py vba_manager.py table create Sheet1!A1:D10 売上表 [--no-headers]
py vba_manager.py table list
py vba_manager.py table delete 売上表
# 列・フィルタ・ソート（excel-mcp の table_column 相当）
py vba_manager.py table column add 売上表 単価 [--at 3]       # add/remove/rename/format
py vba_manager.py table column rename 売上表 単価 税込単価
py vba_manager.py table column format 売上表 数量 "#,##0"     # 引数なしで取得
py vba_manager.py table filter 売上表 数量 ">8"              # 条件フィルタ
py vba_manager.py table filter-values 売上表 商品 りんご ぶどう # 値フィルタ
py vba_manager.py table filters 売上表 / filter-clear 売上表
py vba_manager.py table sort 売上表 数量 [--desc]
py vba_manager.py table sort-multi 売上表 商品:asc 数量:desc

# 名前付き範囲
py vba_manager.py name add 基準値 Sheet1!A2
py vba_manager.py name list
py vba_manager.py name delete 基準値
```

#### 手コマンド 第2弾（編集の足回り／検索置換／保存印刷／仕上げ）

```bash
# --- a. 編集の足回り ---
py vba_manager.py row insert 5 2          # 5行目に2行挿入 / row delete 5 2
py vba_manager.py col insert C 1          # C列に1列挿入 / col delete C 1
py vba_manager.py copy-range A1:C1 E1     # 範囲コピー（--values で値のみ）
py vba_manager.py fill D2:D5              # 先頭セルを下にフィル（--right で右）
py vba_manager.py sort A1:C20 --key B --desc --header   # B列キーで降順、見出しあり
py vba_manager.py autofilter A1:C20       # オートフィルタ設定（--off で解除）

# --- b. 検索・置換 ---
py vba_manager.py find 田中 --book        # 全シート横断で探して番地を返す（--whole/--formula）
py vba_manager.py find-replace 旧 新 A1:Z99   # 範囲一括置換（範囲省略で使用範囲全体）

# --- c. 保存・印刷まわり ---
py vba_manager.py save                    # 上書き保存（手コマンドはOK出たらこれで確定）
py vba_manager.py save-as "C:\path\out.xlsx"   # 別名保存（拡張子で形式判定）
py vba_manager.py print-setup --area A1:H50 --title-rows 1:3 --landscape --fit-wide 1
#  指定可: --area --title-rows 1:3 --title-cols A:B --landscape/--portrait
#          --fit-wide N --fit-tall N --zoom N --center-h --center-v

# --- d. 仕上げ・見た目 ---
py vba_manager.py cond-format B2:B20 --gt 85 --bg "#FFC7CE"   # 85超を赤に（--clearで全削除）
#  比較: --gt --lt --ge --le --eq --ne 値 / --between v1 v2、色: --bg --color、--bold
py vba_manager.py hyperlink A1 "https://..." --text "リンク"   # --remove で削除
py vba_manager.py validation C2:C20 --list "A,B,C"            # ドロップダウン（--clearで削除）
py vba_manager.py freeze B2               # B2の左上で枠固定 / freeze off
py vba_manager.py comment A1 "見出し"     # セルコメント / comment A1 --remove
```

#### 重量級コマンド

```bash
# (1) チャート  ※ 重量級は一通り実装済（PowerQuery の load 読み込み配線まで完了）
py vba_manager.py chart create A1:B5 --type column --title "月別売上" --name 売上グラフ
#  --type column|bar|line|pie|scatter|area|doughnut（既定 column）
#  --at セル（左上を合わせる／省略時はデータ範囲の右隣）--width N --height N
py vba_manager.py chart list
py vba_manager.py chart delete 売上グラフ
# (1b) グラフ詳細設定 chart-config（excel-mcp の chart_config 相当）
py vba_manager.py chart-config set-title 売上グラフ "月別売上"
py vba_manager.py chart-config set-type 売上グラフ line
py vba_manager.py chart-config set-axis-title 売上グラフ value 数量
py vba_manager.py chart-config legend 売上グラフ right            # bottom/top/right/left/corner/off
py vba_manager.py chart-config style 売上グラフ 5                 # 組込スタイル 1-48
py vba_manager.py chart-config axis-scale 売上グラフ value --min 0 --max 40 --major 10
py vba_manager.py chart-config gridlines 売上グラフ value --major on --minor off
py vba_manager.py chart-config axis-format 売上グラフ value "#,##0"   # 引数なしで取得
py vba_manager.py chart-config data-labels 売上グラフ --value --percent [--position outsideend]
py vba_manager.py chart-config placement 売上グラフ 3            # 1=移動+サイズ/2=移動のみ/3=自由
py vba_manager.py chart-config add-series 売上グラフ C2:C4 --series-name 在庫 [--category-range A2:A4]
py vba_manager.py chart-config remove-series 売上グラフ 2
py vba_manager.py chart-config series-format 売上グラフ 1 --marker-style 2 --marker-size 8 --marker-bg "#FF0000" [--invert]
py vba_manager.py chart-config trendline add 売上グラフ 1 linear   # list/add/delete、種別 linear/exponential/logarithmic/movingaverage/polynomial/power

# (2) ピボットテーブル
py vba_manager.py pivot create 元データ!A1:C100 --rows 部門 --cols 月 --values 売上 --func sum --sheet 集計 --name 売上ピボット
#  --rows/--cols/--values はカンマ区切りで複数可、--func sum|count|average|max|min（既定 sum）
#  出力先: --sheet シート名（無ければ作成）/ --at セル / 省略時は新規シート「ピボット」
py vba_manager.py pivot list
py vba_manager.py pivot delete 売上ピボット
# (2b) フィールド管理 pivot-field（excel-mcp の pivottable_field 相当）
py vba_manager.py pivot-field list 売上ピボット
py vba_manager.py pivot-field add-row|add-col|add-filter 売上ピボット 部門
py vba_manager.py pivot-field add-value 売上ピボット 売上 --func sum [--name 表示名]
py vba_manager.py pivot-field remove 売上ピボット 部門
py vba_manager.py pivot-field set-func 売上ピボット 売上 average     # 値フィールドの集計
py vba_manager.py pivot-field set-name 売上ピボット 部門 部署        # 表示名(Caption)
py vba_manager.py pivot-field set-format 売上ピボット 売上 "#,##0"
py vba_manager.py pivot-field set-filter 売上ピボット 部門 営業 開発  # 表示する値を限定
py vba_manager.py pivot-field sort 売上ピボット 部門 desc
py vba_manager.py pivot-field group-date 売上ピボット 日付 months    # days/months/quarters/years
py vba_manager.py pivot-field group-numeric 売上ピボット 金額 0 1000 100
# (2c) 計算フィールド・レイアウト pivot-calc（excel-mcp の pivottable_calc 相当）
py vba_manager.py pivot-calc get-data 売上ピボット                   # 出力範囲の値を表示
py vba_manager.py pivot-calc calc-field create 売上ピボット 利益 "=売上-原価"   # その後 add-value で値に
py vba_manager.py pivot-calc calc-field list|delete 売上ピボット [名前]
py vba_manager.py pivot-calc layout 売上ピボット tabular            # compact/tabular/outline
py vba_manager.py pivot-calc subtotals 売上ピボット 部門 off
py vba_manager.py pivot-calc grand-totals 売上ピボット both off      # rows/cols/both

# (3) スライサー（ピボット名 or テーブル名に紐づけ）
py vba_manager.py slicer add 売上ピボット 部門 --name 部門スライサー --at H1
py vba_manager.py slicer list
py vba_manager.py slicer delete 部門スライサー
```

#### 計算モード（大量書き込みの高速化）

```bash
py vba_manager.py calc-mode             # 現在のモードを表示
py vba_manager.py calc-mode manual      # 手動に（大量 write-range の前に）
py vba_manager.py calc-mode recalc      # 今すぐ一括再計算
py vba_manager.py calc-mode auto        # 自動に戻す
```

#### PowerQuery（一覧・更新・作成・M式書換・削除・読み込み配線）

```bash
py vba_manager.py powerquery list           # クエリ（M式行数・説明）と接続の一覧
py vba_manager.py powerquery refresh         # 全クエリ/接続を更新（RefreshAll）
py vba_manager.py powerquery refresh 売上    # 指定クエリ/接続を更新
py vba_manager.py powerquery add 商品マスタ          # _last_query.m のM式から新規作成（接続のみ）
py vba_manager.py powerquery add 数列 --m "let S=#table(...) in S"   # インラインM式
py vba_manager.py powerquery edit 商品マスタ          # _last_query.m のM式で既存クエリを書き換え
py vba_manager.py powerquery delete 数列     # クエリ削除
py vba_manager.py powerquery load 商品マスタ --to sheet            # シートにテーブルとして読み込み（アクティブシートA1）
py vba_manager.py powerquery load 商品マスタ --to sheet --sheet 一覧 --at B2  # 出力先シート・左上セル指定
py vba_manager.py powerquery load 商品マスタ --to model            # データモデル(Power Pivot)に読み込み
```

> add/edit とも M式は `_last_query.m`(UTF-8) / `--m-file f` / `--m "..."` から取得。
> add は **接続のみ**のクエリを作る。edit は `WorkbookQuery.Formula` への代入（書込可・検証済）。
> **load（読み込み配線・実装済/検証済）**: 接続のみクエリをシートのテーブルまたはデータモデルに読み込む。
> - `--to sheet`: `ListObjects.Add(xlSrcExternal)` でシートに表として出す。作られる接続は
>   `Query - <名前>` に揃えるので `powerquery refresh <名前>` で更新できる。`--sheet`/`--at` で出力先指定。
> - `--to model`: `Connections.Add2(..., CommandText=クエリ名, lCmdtype=6=xlCmdTableCollection, CreateModelConnection=True)`。
>   この形だとモデルテーブル名がクエリ名になる（SQL/SELECT形式だと "クエリ" の汎用名になるので不可）。
>   一度モデルを使うと空でも `ThisWorkbookDataModel` 接続が残るが無害。`datamodel list` で確認。
> いずれも **保存はしない**（Excelで確認後に手動保存、破棄したいなら保存せず閉じる）。

#### コネクション / データモデル

```bash
# ブック接続（外部データ・クエリ接続）
py vba_manager.py connection list            # 種別・接続文字列つき一覧
py vba_manager.py connection refresh 顧客接続  # 更新（name省略で全件RefreshAll）
py vba_manager.py connection delete 顧客接続   # 削除

# データモデル（Power Pivot）
py vba_manager.py datamodel list             # テーブル(行数)・リレーション・メジャーを一覧
py vba_manager.py datamodel relation add    売上 商品ID 商品マスタ ID  # リレーション作成
py vba_manager.py datamodel relation delete 売上 商品ID 商品マスタ ID  # リレーション削除
py vba_manager.py datamodel measure add 売上 合計数量 --dax "SUM('売上'[数量])"  # メジャー(DAX)作成
py vba_manager.py datamodel measure add 売上 平均数量            # DAXは _last_dax.dax(UTF-8) から
py vba_manager.py datamodel measure add 売上 売上額 --dax "SUM('売上'[数量])" --format currency --symbol JPY --decimals 0
py vba_manager.py datamodel measure add 売上 構成比 --dax "DIVIDE(...)" --format percent --decimals 1
py vba_manager.py datamodel measure delete 合計数量             # メジャー削除
```

> **datamodel relation add/delete（実装済/検証済）**: `ModelRelationships.Add(FK列, PK列)` で作成。
> 引数順は **FKテーブル FK列 PKテーブル PK列**（FK=多側/参照する側、PK=一側/参照される側）。
> PK側の列は値が一意であること。テーブル自体の追加は `powerquery load --to model`。
>
> **datamodel measure add/delete（実装済/検証済）**: `ModelMeasures.Add(名前, テーブル, DAX, 書式, 説明)`。
> DAX は `--dax` / `--dax-file` / `_last_dax.dax`(UTF-8) から取得。**先頭の `=` は自動で外す**（付けても可）。
> **DAX 内の日本語テーブル名はシングルクォート必須**（例: `SUM('売上'[数量])`／`売上[...]` は構文エラー）。
> 書式は `--format general|whole|decimal|currency|percent|scientific`（既定 general）。
> `--decimals N`（小数桁）、`--thousands`（桁区切り）、`--symbol`（**通貨コード** USD/JPY/EUR。`$`等のグリフは不可、無効時は既定にフォールバック）。
> 引数付き書式は `GetModelFormat*` メソッド経由（`ModelFormat*` プロパティは既定値専用で引数を渡せない）。

> **write-range のグリッド入力**: `_last_values.tsv`（UTF-8・タブ区切り）に
> 行＝改行、列＝タブで書き出してから `write-range <左上セル>` で適用する。
> `get` → `_last_proc.vba` と同じ発想。`'='` 始まりのセルは数式として書き込まれる。

## 標準作業フロー（マクロ修正）

この手順を必ず守ること。勝手にコードを変更しない。対象は今アクティブにしているブック。

1. `py vba_manager.py list` でマクロ一覧を確認
2. `py vba_manager.py get <Sub名>` で対象コードを取得
   - 同名プロシージャが複数フォームにある場合は **モジュール指定** を使う：
     `py vba_manager.py get <モジュール名> <Sub名>`
3. `_last_proc.vba` を Read ツールで読み、修正内容を検討
4. 修正後のコードを `_last_proc.vba` に Write
5. `py vba_manager.py replace-procedure -y` で適用（非対話実行のため -y 必須）
6. ユーザーに動作確認を依頼

## 標準作業フロー（フォームの .bas を修正して適用）

フォームや複数プロシージャをまとめて修正する場合（.bas ファイルを直接編集）：

1. `py vba_manager.py export-module <モジュール名>` で最新の .bas をエクスポート
2. Python スクリプトで CP932 のまま編集する（Edit/Write ツール禁止）：
   ```python
   # 編集スクリプトの雛形
   path = r"SCRIPTS\対象フォーム.bas"
   with open(path, 'r', encoding='cp932') as f:
       lines = f.readlines()
   # 行番号ベースで修正（多行文字列マッチは使わない）
   # lines[N] = 新しい行内容
   with open(path, 'w', encoding='cp932') as f:
       f.writelines(lines)
   ```
   **行番号は inspect スクリプトで事前確認する：**
   ```python
   with open(path, 'r', encoding='cp932') as f:
       lines = f.readlines()
   for i, line in enumerate(lines, 1):
       if "対象キーワード" in line:
           print(f"L{i}: {repr(line)}")
   ```
3. `py vba_manager.py replace-module <モジュール名> <basファイル>` で適用

## 標準作業フロー（標準モジュール全体の適用）

.bas ファイルを編集してモジュール全体を置換する場合：

1. `py vba_manager.py export-module <モジュール名>` でエクスポート
2. .bas ファイルを Python で CP932 のまま編集
3. `py vba_manager.py replace-module <モジュール名> <basファイル>` で適用

## UserForm の作成・修正

form_builder.py を使う。既存の実装例: create_f_calendar.py, create_node_edit_form.py
（UserForm 専門の作業は excel-userform-builder スキルも参照）

```python
from form_builder import FormBuilder, add_btn, add_lbl, add_txt, add_lst, add_combo

with FormBuilder.connect() as fb:  # アクティブブック接続
    frm = fb.get_or_create("FormName", caption="タイトル", width=300, height=200)
    f = fb.clear_controls(frm)
    add_btn(f, "BtnOK", "OK", 80, 160, 60, 20)
    add_lbl(f, "Label1", "テキスト", 10, 10, 100, 18)
    add_txt(f, "TextBox1", 10, 30, 200, 22)
    add_lst(f, "ListBox1", 10, 60, 200, 90)
    fb.inject_vba(frm, "_form_code.vba")  # VBAコードを注入
    fb.save()
```

デフォルトフォントは全コントロール 12pt。

## 絶対に守るべきルール

### 勝手な変更の禁止
- 指示されていない機能を追加しない
- 指示されていないコードを変更しない
- 影響範囲を確認せずに変更しない

### エンコーディング（最重要 - 違反厳禁）
- .bas ファイルは **CP932**（Shift-JIS）で保存
- win32com 経由の文字列は Unicode
- _last_proc.vba は UTF-8

**⚠ .bas ファイルに Edit ツール・Write ツールを絶対に使うな ⚠**
Claude の Edit / Write ツールは UTF-8 で書き込むため、CP932 の .bas ファイルが破壊される。
VBA にインポートするとモジュール名・プロシージャ名・日本語文字列が全て文字化けする。

.bas ファイルを修正するときは必ず以下のいずれかを使うこと：
1. **Python で CP932 のまま読み書き**:
   ```python
   with open(path, 'r', encoding='cp932') as f:
       content = f.read()
   # 修正処理
   with open(path, 'w', encoding='cp932') as f:
       f.write(content)
   ```
2. **_last_proc.vba 経由で replace-procedure**（プロシージャ単位の修正）

### 取り込み前の単体検査 check-bas（COM不要・バイパス時の安全網）

`.bas` を VBA に取り込む前に、機械的な事故（文字コード・改行二重化・重複）を1コマンドで検査する。
**COM接続が落ちていても動く**ので、何らかの理由でツールの通常経路を通さず手書きで `.bas` を
作った場合でも、取り込み前にこれを必ず通すこと。

```bash
py vba_manager.py check-bas <file.bas>         # 検査だけ（問題があれば終了コード1）
py vba_manager.py check-bas <file.bas> --fix   # 改行二重化(\r\r\n)だけCP932のまま自動修正
```
検査項目: ①UTF-8化/BOM（文字コード事故）②改行二重化（行数が倍に膨れる事故）
③Sub/Function 名の重複（重複挿入）④連続する同一コード行（重複挿入の臭い・警告）。
③④は判断が要るので自動修正しない（報告のみ）。replace-module は取り込み直前に①②を
自動で行うが、**ツールを通さず Import する場合の最後の砦**がこの check-bas。

### モジュール適用方式
- AddFromString は Attribute 行を正しく処理しないことがある
- **必ず Remove + Import 方式（replace-module）を使うこと**
- ショートカットキーは Attribute VB_ProcData.VB_Invoke_Func で定義される

### replace-module の副作用
- Remove+Import でモジュールが VBComponents の末尾に移動する
- マクロの表示順（メニュー等）が変わる場合がある
- 影響を受けたモジュールも replace-module して順番を揃える

### InsertLines の改行問題
- Python の \n では正しく複数行に分割されないことがある
- .bas ファイル直接編集方式を使うこと

## 変更前の確認チェックリスト

コードを変更する前に必ず以下を確認：

1. 変更対象のフォーム/モジュールが他の機能から参照されていないか
2. AutoFilter の起点列（B1始まりかA1始まりか）
3. 既存のコントロール名と用途（別機能で使われていないか）
4. 変更がユーザーの指示の範囲内か
