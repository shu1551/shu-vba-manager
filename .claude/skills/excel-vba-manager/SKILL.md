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

## 作業の型（速度既定・2026-07-12 制定）

「時間かかりすぎ」の叱責を受けて制定。**段取りをその場で考え直さず、この型どおり動く。**

1. **道具は MCP が第一選択**（`mcp__vba-manager__vba("コマンド1行")` ほか）。常駐 COM で応答 0.01〜0.2 秒、
   PowerShell の引用符/エンコーディング事故もない。確認プロンプト系は必ず `-y`。
   PowerShell + CLI に落とすのは ①vba_manager 本体を改修中で新コードを検証するとき
   （MCP はセッション起動時のコードを保持し、途中の本体修正を反映しない）②MCP 不在時、の2つだけ。
2. **点検の型**: list → export-all → コードを読む → 所見を数行で報告（詳細は聞かれたら）。
   1,000 行超のブックはモジュール別に並列サブエージェントへ読ませて壁時計を圧縮。
   対象モジュールが分かっているときの list は `--module 名` で絞る
   （実測: 182マクロ規模のブックで `list --detail` 全件は約3秒、`--module` 絞りはその場でほぼ即応）。
3. **修正の型（差分の大きさで経路を選ぶ・2026-07-17 制定）**:
   - **1〜数行の小修正（引数追加・条件変更・文言修正など）**: `code-replace "旧" "新" --module 名 -y` が最短。
     **変更行だけ送る＝プロシージャ全文を再送しない**（diff プレビュー・バックアップ・ReplaceLine 方式で
     Attribute のショートカット定義も壊れない）。実測では所要時間の支配項はツール実行（0.01〜0.2秒）
     ではなくコード全文の打ち直しであり、小修正の全文再送は数十倍遅い。
   - **新規プロシージャ**: add-procedure。**プロシージャの大改造（過半が変わる）**: get→replace-procedure
     （自動バックアップ・自動保存つき）。
   - モジュール先頭行（Option Explicit 等）だけは replace-module か COM 直接（InsertLines）。
4. **検証の型（規模に応じて使い分け）**:
   - **1〜2プロシージャの軽微な修正**: `run-macro` 1本（コンパイル確認を兼ねる）だけで良い。
   - **複数モジュールにまたがる修正**: 影響するマクロの run-macro / test を同じ往復に相乗り。
   小修正にE2Eや過剰検証を足さない。**検証を独立の往復にしない**——最後の置換と同じ往復に
   run-macro を相乗りさせれば、検証の追加コストはツールの数秒だけになる。
5. **公開の型**: `py publish_check.py <公開用コピー> --scrub` → commit/push → GitHub 生ファイルを
   ダウンロードして publish_check で再チェック（実測まで）。
6. **報告は結果数行が既定**。往復を減らす（独立な調査・実行は1往復に並列で畳む）。

## ツールの場所

```
SCRIPTS:            vba_manager.py 等のツールが置いてあるフォルダ（絶対パス決め打ち禁止。不明なら Glob **/vba_manager.py で特定）
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
cd <SCRIPTS>   # vba_manager.py のあるフォルダへ。不明なら Glob **/vba_manager.py で特定（絶対パスを決め打ちしない）

# マクロ一覧表示
py vba_manager.py list
#  オプション: --standard（標準モジュールのみ）--personal（PERSONAL.XLSB）--addin（アドイン）
#              --all（全ブック横断）--json（機械可読出力）
#              --detail（所属モジュール・行数・先頭コメント付き）--module 名（対象モジュール限定）

# 全モジュール横断のVBAコード検索（どのマクロが何を使っているか調べる）
py vba_manager.py grep "ActiveSheet"                  # [モジュール] プロシージャ:行番号: 該当行
py vba_manager.py grep "On Error" --module shu003 -i  # --regex（正規表現）--max N --json も可

# ブックの構成ドキュメント（取説）を自動生成 → _last_docs.md
py vba_manager.py docs [--out f.md] [--preview 3]  # --preview で各シート先頭N行のMarkdown表も含める

# マクロの呼び出し関係を解析（Call/Application.Run/裸呼び）
py vba_manager.py call-graph                    # 未解決Call(存在しないマクロ呼び=一語バグ)・関係一覧・孤立
py vba_manager.py call-graph --macro 親処理      # そのマクロ起点の呼び出しツリー＋呼び元
py vba_manager.py call-graph --mermaid          # Mermaid図を _last_callgraph.md に（未解決は赤ノード）

# 対話セッション（接続を張ったままコマンドを打ち続ける。2コマンド目から再接続なし）
py vba_manager.py shell                         # exit で終了。batch のファイル版に対する対話版
#  history で履歴一覧、!番号 で再実行（!! は直前のコマンド）

# ブックの健康診断は vba_manager から退役（2026-07-17）
#  AI の裁量で診断が乱発される構造を絶つため、独立ツールへ分離し本体からコマンドを撤去した。
#  診断は使う人が自分の手で実行するもの＝AI はこのスキル経由で診断を回さない。
#  ※ シート検査は COM Find を使うため Excel の検索ダイアログ設定が変わる（内容は変えない）

# マクロ修正前の影響範囲予告（このマクロに手を入れるとどこまで波及するか）
py vba_manager.py impact <マクロ名>              # 別名: 影響範囲。--json も可
#  ■呼び元（直すと影響が及ぶ先・間接含む。フォームのボタン/イベントも出る）
#  ■呼び先（依存している部品）■入口（ショートカット/自動実行イベントから届く経路）

# モジュール一覧表示
py vba_manager.py list-modules

# 開いているブックの一覧（Excel生存確認にも。素の py -c ワンライナーでなくこれを使う）
py vba_manager.py list-open

# マクロを実行（ブック名!マクロ名 or マクロ名だけなら全プロジェクトから検索）
py vba_manager.py run-macro <マクロ名> [--json]

# モジュール丸ごと削除（要約表示→確認→バックアップ。戻すのは restore）
py vba_manager.py delete-module <モジュール名> -y

# マクロを引数つきで実行（数値に見える引数は数値化して渡す）
py vba_manager.py run-macro 加算 3 4        # Function なら戻り値も表示

# 予行演習run（2026-07-15）＝マクロをコピーに試し撃ちして差分報告（本体は無傷）
py vba_manager.py rehearse <マクロ名> [引数...] [--auto-dialog ok] [--addins] [--out path] [--discard] [--max N]
py vba_manager.py 予行演習 <マクロ名>            # 日本語別名（同じもの）
#  流れ: SaveCopyAs（未保存の変更込みの「今この瞬間」）→ 別インスタンスでコピーを開く
#  （Workbook_Openは起こさない）→ 前snapshot → 実行 → 後snapshot → snapshot-diff報告。
#  マクロが途中で落ちても、そこまでの変化を差分で報告する（それも収穫）。
#  結果コピーと前後snapshotは残す（--discard で削除）。本体に焼く判断は人がやる。
#  ⚠ 非破壊が保証できるのはブックの中だけ。ファイル出力・メール送信など外への副作用は止められない。
#  アドイン（秀.xlam等）の関数を呼ぶマクロは --addins を付ける（演習用Excelは素の環境のため）。

# VBAテストランナー（xlflowのテスト基盤から発想だけ移植・2026-07-09）
py vba_manager.py test [excel_file] [絞り込み] [--module 名] [--auto-dialog ok] [--json]
py vba_manager.py テスト                     # 日本語別名（同じもの）
#  名前が「テスト」または「test」で始まる**引数なしの公開Sub**を一括実行し成否一覧。
#  失敗の知らせ方は Err.Raise だけ（assert例: If 実際 <> 期待 Then Err.Raise 5, , "説明"）。
#  補助モジュール不要＝テストSubも単体で他ブックへ移植できる自立ユニット。
#  実行はエラー捕捉ハーネス（一時モジュール注入→直接呼びラッパー→撤去）経由なので
#  実行時エラーでVBAダイアログは出ない。ブックは保存しない。
#  全部成功で終了コード0／1本でも失敗なら1（自動化ゲートに使える）。
#  ⚠ テスト中にMsgBoxを出すコードには --auto-dialog ok を付ける（付けないと止まる）。

# ショートカットキー付きマクロの一覧 / メニュー表示順の入替
py vba_manager.py list-shortcuts
py vba_manager.py reorder-macro <マクロ名> <up|down|top|bottom|位置番号>   # top/bottom/番号は一発移動

# 導入セルフ診断（初心者が最初に打つ1コマンド。Python/pywin32/Excel/VBOM信頼設定を○×表示）
py vba_manager.py setup-check [--json]

# 環境診断 / VBA構文チェック / プリンタ
py vba_manager.py diag
py vba_manager.py check [excel_file]            # ブックの全モジュールを静的診断（.bas 単体は check-bas）
py vba_manager.py printer-list
py vba_manager.py printer-setup --printer <プリンタ名> [--duplex ...]
#  ↑ プリンタ名は必ず --printer で渡す（位置引数は無視され、既定プリンタの設定が変わる）。
#     OS のプリンター設定そのものを即時・不可逆に変更し、他アプリにも影響する。

# 全マクロ横断の一括置換（grepの対。diffプレビュー→確認→バックアップ→変更行だけReplaceLine）
py vba_manager.py code-replace "旧" "新" [-y] [--regex] [--module 名]
#  ReplaceLine方式なのでショートカット定義(Attribute)は壊れない。戻すのは restore

# プロシージャのコード取得 → _last_proc.vba に保存
py vba_manager.py get <Sub名>

# モジュール指定してプロシージャ取得（同名プロシージャが複数ある場合に使う）
py vba_manager.py get <モジュール名> <Sub名>
py vba_manager.py get <モジュール名>.<Sub名>   # ドット区切りも可
py vba_manager.py get 名1 名2 名3              # 3個以上は複数取得（1接続・連結保存。書き戻しは1本ずつ）
#  --out f で _last_proc.vba 以外に保存（参照用コピー）、--json も可。見つからないときは近似候補を提示

# _last_proc.vba の内容でプロシージャを置換（バックアップ自動作成）
py vba_manager.py replace-procedure -y
#  ※ -y/--yes で確認プロンプト(y/N)をスキップ。Claude経由の非対話実行では必ず -y を付ける。
#    （-y なしだと input() 待ちで止まる。PowerShellから "y" をパイプで流す方式は失敗することがある）
#    差分(Diff)は -y を付けても表示される。--module <名> で対象モジュールの明示も可。
#    --code-file <f> で _last_proc.vba 以外のコードファイルを指定。差分ゼロなら置換せずスキップ。
#    バックアップが取れないと停止する（--force で強行可）。

# モジュール全体を Remove+Import で置換
py vba_manager.py replace-module <モジュール名> <basファイル>
#  .bas の Attribute VB_Name と指定モジュール名が一致しないと停止（対象取り違え防止）。
#  Import 失敗時はモジュールバックアップから自動復旧を試みる。

# 新規プロシージャの追加 / 削除（get→replace と対称の軽量経路）
py vba_manager.py add-procedure <モジュール名> -y      # _last_proc.vba のコードを末尾に追加（同名重複は停止）
py vba_manager.py delete-procedure <Sub名> -y          # 削除コードを表示して確認。--module で対象明示

# モジュールを .bas にエクスポート
py vba_manager.py export-module <モジュール名>
py vba_manager.py export-all [--dir 出力先] [--check]  # 全モジュール一括（1接続・--check で検査つき）

# バックアップの一覧と復元（undo導線）
py vba_manager.py list-backups [キーワード]            # 新しい順・COM不要
py vba_manager.py restore <バックアップ.bas>           # VB_Name から対象を特定し replace-module 経路で復元

# コマンド列を1接続で連続実行（一括作業の高速化。実測: 18本export 数分→約2秒級）
py vba_manager.py batch cmds.txt                       # 1行=1コマンド。#始まりと空行は無視
printf 'get shu003 マクロA\nreplace-procedure -y\n' | py vba_manager.py batch -   # 標準入力からも可
#  途中失敗で停止（--keep-going で最後まで実行）
```

### 目コマンド（シート状態の読み取り・読み取り専用）

```bash
py vba_manager.py read-range     [excel_file] [range...]  # セル値をテキスト格子で読む（複数範囲可・1接続）
py vba_manager.py read-range     A1:D10 --formula        # 計算結果でなく数式(.Formula)を表示
py vba_manager.py read-range     "集計!A1:D50" --tsv     # _last_values.tsv に書き出し→編集→write-range で書き戻す往復
py vba_manager.py read-range     A1:D10 --width 80       # 列の表示幅を広げる（既定40、切り詰めは…付き）。--json も可
py vba_manager.py read-selection [excel_file] [--formula] # 今選択している範囲を読む（--formulaで数式）
py vba_manager.py sheet-info     [excel_file] [--preview 3] # シート構成一覧（--preview で各シート先頭N行も＝ブック俯瞰が1接続）
py vba_manager.py table read <テーブル名> [--tsv]        # テーブル名で直接読む（番地調べ→read-range の2段を1段に）
py vba_manager.py screenshot     [excel_file] [range] [--out f.png]  # 範囲を画像(PNG)で書き出す（省略時 _last_view.png）
py vba_manager.py snapshot       [excel_file] [--sheet 名] [--out f.json] [--max-rows N]
                                 [--format-rows N] [--no-format]
#  ブック(または1シート)を意味構造JSONに畳む（セル疎+結合+書式+図形/ボタン(OnAction)+テーブル）
#  → _last_snapshot.json。機械的事実だけ吐き意味付けはAIがやる＝このJSONを Read して
#  質問に答えれば「開いたままブックLM」になる
#  書式（四つ目の目）: フォント/太字/文字色/塗り/罫線/表示形式/横位置/列幅/行高。
#    シート全体→列→行の順に「揃っているか」で降りて採る（セル単位は1万セルで33秒＝実用外）。
#    値 None は「その範囲の中で不揃い」の意味（欠測ではない）。読めなければ format_error を残す。
#    行ごとにCOM往復が要るので既定2000行まで（--format-rows）。急ぐときは --no-format
py vba_manager.py snapshot-diff  <before.json> [after.json] [--max N]
#  2つの snapshot の機械的差分（COM不要）。after 省略時は _last_snapshot.json（直近のsnapshot）と比較。
#  セル変更/追加/削除・結合・書式・図形（文字/OnAction/位置/大きさ）・テーブルをシート別に列挙
#  ★クリーン化系マクロは値も結合も図形も残したまま書式だけを吹き飛ばす。書式の目が無いと
#    これを「差分なし（一致）」と報告する（ポスター.xlsm 破壊の教訓）。
#  シート全体で起きた変化は列・行では繰り返さない（局所的な変化だけが個別に出る）
py vba_manager.py wiring         [excel_file] [--json]    # 別名: 配線図
#  ボタン⇔マクロの配線図＝シート上の全 OnAction（グループ内も展開）をマクロ名簿と突き合わせ、
#  行き先のないボタン（存在しないマクロを指す壊れた配線）を検出。孤立マクロ側は call-graph で
```

> **screenshot の注意**: 内部で一時グラフの作成・削除を伴うため、目コマンドの中で唯一
> **ブックの Undo 履歴が消える**（内容は変えない）。実行後は選択を A1 に戻しクリップボードもクリアする。

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
py vba_manager.py write-range A1 "007" --raw        # 数値変換せず文字列で書く（先頭ゼロ保持）
#  ※ シート名だけの指定（使用範囲全域）は write-range では拒否される。セル/範囲を明示すること。

# 範囲をクリア（既定すべて／--contents 値のみ／--formats 書式のみ）
py vba_manager.py clear-range A1:D10 --contents
#  ※ シート名だけの指定（使用範囲全域）は --whole-sheet を付けない限り拒否される。
#    余分な位置引数もエラーになる（"Sheet1 A1:B2" と分けて渡す事故の防止）。fill / sort も同様。

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
py vba_manager.py table delete 売上表     # テーブル解除（データは残る）
py vba_manager.py table ref 売上表 [列名]  # 構造化参照の書き方と列一覧を表示
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
py vba_manager.py row insert 5 2          # 5行目に2行挿入 / row delete 5 2（--sheet 名 で対象明示可）
py vba_manager.py col insert C 1          # C列に1列挿入 / col delete C 1
py vba_manager.py copy-range A1:C1 E1     # 範囲コピー（--values で値のみ）
py vba_manager.py fill D2:D5              # 先頭セルを下にフィル（--right で右）
py vba_manager.py sort A1:C20 --key B --desc --header   # B列キーで降順、見出しあり
py vba_manager.py autofilter A1:C20       # オートフィルタ設定（--off で解除）

# --- b. 検索・置換 ---
py vba_manager.py find 田中 --book        # 全シート横断で検索（--whole/--formula/--max N）
#  結果は「シート名!$A$1: 値」形式＝そのまま次コマンドの range 引数に貼れる
py vba_manager.py find-replace 旧 新 A1:Z99   # 範囲一括置換（範囲省略で使用範囲全体）
#  ヒットセル数を事前カウントして報告、0件なら置換しない。--match-case で大小文字区別。

# --- c. ブックの開閉・保存・印刷まわり ---
py vba_manager.py open "C:\path\book.xlsm"     # ブックを開く（2026-07-15）
#  「人が開くのと同じ場所」に開く: 既に開いていれば前面化のみ／見えているExcelが居れば合流／
#  Excel未起動なら通常起動（アドイン・PERSONALが普段どおり読み込まれる）。
#  非表示の残骸Excelしか居ないときは取り込まれないよう開かずに止まる。
py vba_manager.py close <ブック名> --no-save -y   # ブックを1冊閉じる（2026-07-15）
#  鎧の三点セット: ①ブック名指し必須 ②--save/--no-save の明示必須 ③確認プロンプト（-y でスキップ）。
#  Excel本体は終了しない。PERSONAL.XLSB とアドインブックは閉じない。同名複数はフルパス名指しを要求。
py vba_manager.py save                    # 上書き保存（手コマンドはOK出たらこれで確定）
py vba_manager.py save-as "C:\path\out.xlsx"   # 別名保存（拡張子で形式判定）
#  既存ファイルへは --overwrite が無いと停止。未対応拡張子はエラー。
#  xlsm→xlsx はマクロが落ちる旨を警告。[excel_file] を第1引数に取る流儀も他コマンドと同じ。
py vba_manager.py export-pdf 出力.pdf [--sheet 名 | --range "集計!A1:H50"] [--overwrite]  # PDF出力（ブックは変更しない）
py vba_manager.py print-setup --area A1:H50 --title-rows 1:3 --landscape --fit-wide 1
#  指定可: --area --title-rows 1:3 --title-cols A:B --landscape/--portrait
#          --fit-wide N --fit-tall N --zoom N --center-h --center-v

# --- d. 仕上げ・見た目 ---
py vba_manager.py cond-format B2:B20 --gt 85 --bg "#FFC7CE"   # 85超を赤に（--clearで全削除）
#  比較: --gt --lt --ge --le --eq --ne 値 / --between v1 v2、色: --bg --color、--bold
#  数式ルール: --formula "=B2>AVERAGE($B$2:$B$20)"（相対参照は範囲左上セル基準）
py vba_manager.py hyperlink A1 "https://..." --text "リンク"   # --remove で削除
py vba_manager.py hyperlink A1                                # そのセルのリンクを表示
py vba_manager.py hyperlink --list [シート名]                  # シート内の全リンク一覧
py vba_manager.py sheet protect シート名 [--password p]       # シート保護 / unprotect で解除
py vba_manager.py format-range A1:B5 --unlock                 # 保護中も編集可のセルに（--lockで戻す）
py vba_manager.py validation C2:C20 --list "A,B,C"            # ドロップダウン（--clearで削除）
py vba_manager.py freeze B2               # B2の左上で枠固定 / freeze off
py vba_manager.py comment A1 "見出し"     # セルコメント / comment A1 --remove
```

> **--append（ログ追記）**: `write-range "ログ!A" 値 --append` で使用範囲の最終行の次に書く
> （書き込み先番地は実行結果に必ず表示される）。
> **--sheet 分離指定**: read-range / write-range / clear-range / format-range / fill / sort /
> find-replace は `--sheet シート名 A1:B2` の分離指定も可（`!` や記号を含むシート名のクォート地獄の救済）。
> row / col の `--sheet` と同趣旨。

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
py vba_manager.py chart-config set-source 売上グラフ A1:C10       # データ範囲の再設定
py vba_manager.py chart-config remove-series 売上グラフ 2          # index 必須（省略はエラー）
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
py vba_manager.py calc-mode auto        # 自動に戻す（automatic も可。recalc は now/calculate も可）
```

#### PowerQuery（一覧・更新・作成・M式書換・削除・読み込み配線）

```bash
py vba_manager.py powerquery list           # クエリ（M式行数・説明）と接続の一覧
py vba_manager.py powerquery refresh         # 全クエリ/接続を更新（RefreshAll）
py vba_manager.py powerquery refresh 売上    # 指定クエリ/接続を更新
py vba_manager.py powerquery add 商品マスタ          # _last_query.m のM式から新規作成（接続のみ、--desc 説明 も可）
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
py vba_manager.py datamodel measure add 売上 平均数量            # DAXは _last_dax.dax(UTF-8) から（--desc 説明 も可）
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

## 現地調査（シートに触るマクロの設計・破壊的操作の前に必ず）

シートを読み書きするマクロの新規作成・改修や、破壊的操作（行列削除・クリア・並べ替え等）の前に、
`snapshot` で対象の実勢（結合セル・見出しの位置・図形/ボタン・テーブル）を先に読む。

- 自動の最終行/最終列判定や「たぶんこうなっている」の推測は手作りシートで誤る
  （実害例: 最終列の誤判定で結合セルを巻き込み削除）。snapshot の事実を見てからコードを書く。
- 変更を伴う作業は**前後で snapshot を取り snapshot-diff**（変更前後を事実で照合する検分）。
  「実際に何が変わったか」をセル単位の事実で確認・報告する。「直したはず」で済ませない。
- ボタンとマクロの対応を調べるときは `wiring`（行き先のないボタンも検出）、
  マクロ側から見た地図は `call-graph` / `docs`。

## 動作検証（マクロを動かして確かめる。UI操縦は最終手段）

修正したマクロの動作確認は **COM経由を第一選択**にする。computer-use での
マウス/キー操縦は使わない（打鍵がダイアログからシートに漏れ、選択範囲に Delete を
送る事故が実際に起きた。2026-07-12・未保存だったため実害なしの結果オーライ）。

- **ダイアログを開くマクロの検証**: `py vba_manager.py run-macro <マクロ名>` 一発でよい。
  run-macro のダイアログ安全解除ガードが Excel のモーダルダイアログ(#32770)を検出し、
  **タイトル（例: タイトル「セルの書式設定」）と本文を報告してから安全に閉じる**。
  この報告テキストが「どのダイアログが開いたか」の物証になる＝スクリーンショット不要。
- **セルへの効果の検証**: 前後で `snapshot` → `snapshot-diff`（セル単位の事実で差分確認）。
  範囲の見た目は `screenshot [range]`（範囲をPNG化。モーダルダイアログは写らない点に注意）。
- **MsgBox を出すマクロ**: 応答を選びたいときだけ `--auto-dialog ok|cancel|yes|no`。
  既定でもキャンセル優先で自動解除され、内容は報告される（無言ハングしない）。
- どうしても実画面の目視が要るときだけ computer-use。その場合も
  ①打鍵前にスクリーンショットでフォーカス位置を確信 ②矢印連打・BackSpace連打の
  探り打ち禁止 ③シートが背後にある状態で Delete/Enter 系を打たない。

## 標準作業フロー（マクロ修正）

この手順を必ず守ること。勝手にコードを変更しない。対象は今アクティブにしているブック。

1. `py vba_manager.py list` でマクロ一覧を確認
2. `py vba_manager.py get <Sub名>` で対象コードを取得
   - 同名プロシージャが複数フォームにある場合は **モジュール指定** を使う：
     `py vba_manager.py get <モジュール名> <Sub名>`
3. `_last_proc.vba` を Read ツールで読み、修正内容を検討
4. 修正後のコードを `_last_proc.vba` に Write
5. `py vba_manager.py replace-procedure -y` で適用（非対話実行のため -y 必須）
6. 動作を自分で検証できるものは上の「動作検証」の COM 経由手順で確認し、
   その結果を添えてユーザーに最終確認を依頼

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

最短ルートの使い分け（UserForm 専門の作業は excel-userform-builder スキルも参照）:

1. **現状把握（目）**: `py form_inspect.py <フォーム名>` — 1接続でフォーム情報＋コントロール配置＋コード。
   `--font`（Font.Size列）`--json`（機械可読）`--png`（実表示してPNG撮影・見た目の確認）`--list`（フォーム一覧）
   `--png --names`（枠と名前を画像に描き込み＝名指しで直せる）`--png-all`（MultiPage の全タブを1枚ずつ撮影）
   `--lint`（重なり/はみ出し/不揃い/タブ順に加え、**孤児イベントハンドラ・Click未実装・右端の揃い忘れ**も機械検査）
2. **コードだけ直す**: form_builder で作り直さず `vba_manager.py get → replace-procedure`（速い・自動保存）
3. **幾何だけ直す（手）**: `py form_tool.py` — 1接続の機械操作CLI。
   ```
   py form_tool.py scale <フォーム> 1.15            # フォーム+全コントロール+フォント一括倍率
   py form_tool.py set <フォーム> btnOK,btnNG --top 160 --font-size 14
   py form_tool.py move <フォーム> btnA,btnB --dx 8 --dy -4      # 相対移動
   py form_tool.py align <フォーム> btnA,btnB --left 10          # 辺を揃える
   py form_tool.py size-match <フォーム> btnA,btnB --ref btnB    # 同サイズ化
   py form_tool.py distribute <フォーム> a,b,c --vertical --gap 8  # 等間隔配置
   py form_tool.py tab-order <フォーム>                          # TabIndex を視線順に自動整列
   py form_tool.py rename-control <フォーム> btnA btnB   # イベント宣言行も機械追随（件数報告）
   py form_tool.py delete-control <フォーム> btnOld      # イベントコードは残す（警告表示）
   py form_tool.py copy-form <フォーム> <新名>           # レイアウト+コード丸ごと複製
   ```
   既定では保存しない（--save で保存）。lint が指摘した内容はここで直せる。
   rename-control / delete-control は実行前に .frm/.frx を backups へ自動退避する。
   tab-order の自動整列はコンテナ（フォーム/Frame/Page）単位で行われる。
4. **新規・作り直し（推奨: 宣言的レイアウト）**: form_layout.py — 行構造を書くだけで
   ラベル列整列・8ptリズム余白・ボタンバー右寄せ・TabIndex・Default/Cancel まで機械計算。
   ```python
   from form_layout import (build_form, preview_layout, row, lbl, txt, combo, lst, chk,
                            opt_group, button_bar, ok, cancel, spacer, heading, frame)
   rows = [
       heading("顧客情報"),
       row(lbl("顧客名"), txt("txtName")),
       row(lbl("区分"), combo("cmbKind", items=["法人", "個人"])),
       frame("配送オプション",
             row(lbl("優先度"), opt_group(("optHigh", "急ぎ"), ("optNorm", "通常"))),
             row(chk("chkGift", "ギフト包装"))),
       spacer(),
       button_bar(ok("btnSave", "登録", accel="S"), cancel("btnClose", "閉じる")),
   ]
   preview_layout(rows)         # Excel 不要のワイヤーフレームPNG（設計の高速な試行錯誤）
   build_form("F_Order", "受注入力", rows, vba_stub=True, png=True, launcher="Module1")
   #  vba_stub=True: Initialize（items の AddItem・先頭入力へ SetFocus）と
   #  各ボタンの Click 雛形を機械生成して注入。png=True: 構築後に実表示PNG。
   #  launcher="モジュール名": Sub <フォーム名>を開く() を標準モジュールに自動追加
   #  （メニュー方式のブックならそのままメニューに載る）。
   #  combo/lst は rowsource="シート名!A1:A10" でシート範囲に直結も可。
   #  既存フォームは backups へ .frm/.frx 自動退避してから作り直す。
   ```
   **CLI からも使える**: `py form_layout.py preview 宣言.py`（Excel不要の配置図）／
   `py form_layout.py build 宣言.py`（実構築）。同じ宣言ファイルが両方で動く。
   幅未指定の入力は右端まで自動ストレッチ。frame は入れ子1段まで。スタイル定数は form_layout.STYLE。
   **追加部品**: `refedit("refX")`＝範囲選択欄（TextBox+選択ボタン+InputBox(Type:=8)。本物の RefEdit は
   COM挿入が信頼設定でブロックされるため複合部品で実装）、`spin_txt("txtN", min_=1, max_=99)`＝▲▼付き数値、
   `img("imgX", w, h, picture=パス)`＝画像（実行時 LoadPicture）、`txt(..., required=True)`＝ラベルに＊＋
   実行ボタンに空チェック雛形。定番の組み合わせは excel-userform-builder スキルの「定番レシピ」参照。
   **タブ付きフォーム**は `multipage("mpMain", page("基本", row(...)), page("詳細", row(...)))`。
   ページ内に frame も置ける。--to-layout の逆変換・preview_layout（1ページ目＋タブ帯）にも対応済み。
   ※ --png の実表示撮影は「表示中のタブ」しか写らない（他タブは build 後にExcelで切り替えて確認）。
   構築後は `form_inspect --lint` と `--png --names` で検証（デザイン規範は excel-userform-builder スキル参照）。
5. **既存フォームの大改修**: `py form_inspect.py <フォーム> --to-layout` で
   既存フォームを form_layout の宣言コード（たたき台）に逆変換 → 宣言を編集 → build_form。
   往復（逆変換→再構築）で見た目が保たれることは検証済み。items 等の実行時情報だけ手動補完。
6. **自由配置（カレンダー格子等）**: form_builder.py（下記）

```python
from form_builder import FormBuilder, add_btn, add_lbl, add_txt, add_lst, add_combo
from form_builder import Grid, vstack, hstack   # 機械的な座標計算ヘルパー

with FormBuilder.connect() as fb:  # アクティブブック接続
    frm = fb.get_or_create("FormName", caption="タイトル", width=300, height=200)
    f = fb.clear_controls(frm)     # ※レイアウト全消し。コードだけ直すなら使わない（上の2へ）
    add_btn(f, "BtnOK", "OK", 80, 160, 60, 20)
    add_lbl(f, "Label1", "テキスト", 10, 10, 100, 18)
    add_txt(f, "TextBox1", 10, 30, 200, 22)
    add_lst(f, "ListBox1", 10, 60, 200, 90)
    fb.inject_vba(frm)             # VBAコード注入（省略時 _last_form_code.vba / 明示指定も可）
    fb.save()
```

デフォルトフォントは全コントロール 12pt。グリッド配置は `g = Grid(10, 40, 24, 20); g.pos(row, col)`、
縦積み/横並びは `vstack(直前コントロール)` / `hstack(直前コントロール)`。

> フォームの .frm を replace-module に渡すときは **同名 .frx を同じフォルダに置く**こと
> （無いと停止する。.frx にレイアウトが入っているため）。restore も .frm/.frx ペアで復元できる。

## 絶対に守るべきルール

### プロシージャ名・モジュール名は VBA の識別子規則で（先頭 `_` 禁止）
- VBA の識別子は**英字か日本語で始める**。`_`・数字・記号では始められない。
- `_tmp検証` のような先頭 `_` の Sub を注入する事故が過去に繰り返された。
  AddFromString / InsertLines は構文検査をしないため**注入自体は成功報告になり**、
  モジュールがコンパイルエラーで死ぬ（成り済まし成功の典型）。
- 一時検証用の Sub にも普通の名前を使う（例: `tmp検証`・`テスト検証`。`_` を頭に付けない）。
- ツール側でも機械拒否する: validate_vba_code（replace-procedure / add-procedure）・
  replace-module・add-module・check-bas・form_builder.inject_vba・
  form_layout（起動マクロ名）・form_tool copy-form（新フォーム名）が
  識別子規則違反を検出して注入前に停止する。エラーが出たら名前を直す（--force で潰さない）。
- **注入経路の台帳**: test_tools.py の `test_injection_route_ledger` が全注入プリミティブ
  （AddFromString / InsertLines / VBComponents.Import）の所在とガード状態を機械照合する。
  新しい注入経路を実装したら、①ガードを配線 ②台帳に理由つきで登録（未登録はテストが落ちる）。
  「どの経路が塞がっているか」を調べたいときも、この台帳を読めば再調査ゼロで済む。

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
③Sub/Function 名の重複（重複挿入）④VBA識別子規則違反（先頭 `_` 等＝コンパイルエラーになる名前）
⑤連続する同一コード行（重複挿入の臭い・警告）。
③⑤は判断が要るので自動修正しない（報告のみ）。④も自動修正しない（名前を直してから取り込む）。
replace-module は取り込み直前に①②を
自動で行うが、**ツールを通さず Import する場合の最後の砦**がこの check-bas。

### モジュール適用方式
- AddFromString は Attribute 行を正しく処理しないことがある
- **必ず Remove + Import 方式（replace-module）を使うこと**
- ショートカットキーは Attribute VB_ProcData.VB_Invoke_Func で定義される

### replace-module の副作用
- Remove+Import でモジュールが VBComponents の末尾に移動する
- マクロの表示順（メニュー等）が変わる場合がある
- 影響を受けたモジュールも replace-module して順番を揃える
- ショートカットキー（Attribute VB_Invoke_Func）の**セッション登録は Remove で剥がれる**が、
  Import 直後にツールが MacroOptions で自動再登録する（2026-07-12〜）。
  「修正直後にショートカットが効かない・開き直すと直る」症状はこれで解消済み。
  再登録された鍵は実行結果に「ショートカット再登録: Ctrl+Shift+X → マクロ名」と表示される。

### InsertLines の改行問題
- Python の \n では正しく複数行に分割されないことがある
- .bas ファイル直接編集方式を使うこと

## 変更前の確認チェックリスト

コードを変更する前に必ず以下を確認：

1. 変更対象のフォーム/モジュールが他の機能から参照されていないか
2. AutoFilter の起点列（B1始まりかA1始まりか）
3. 既存のコントロール名と用途（別機能で使われていないか）
4. 変更がユーザーの指示の範囲内か
5. シートに触るマクロ・破壊的操作は、snapshot で実勢（結合・見出し・図形）を確認したか（上の「現地調査」）
