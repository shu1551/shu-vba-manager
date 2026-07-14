"""
VBAマネージャー (アクティブブック対応版)

【特徴】
- target_file を省略するとアクティブなExcelブックを自動使用
- get で取得したコードは _last_proc.vba に保存 → Claudeが読み取り・修正
- replace-procedure は _last_proc.vba を自動使用 (--code-file 省略時)
- replace-module は Remove+Import で Attribute を正しく処理

【コマンド一覧】
  list            [excel_file]                     マクロ一覧
  list-modules    [excel_file]                     モジュール一覧
  get             [excel_file] <macro_name>        プロシージャのコード取得
  replace-procedure [excel_file] [--code-file f]  プロシージャを置換
  replace-module  [excel_file] <module> <bas_file> モジュール全体を置換
  export-module   [excel_file] <module>            モジュールを .bas にエクスポート
  diag                                             動作確認

  reorder-macro   <macro> <up|down>                マクロの表示順を入れ替え
  list-shortcuts  [excel_file]                      ショートカットキー一覧

【目コマンド（シート状態の読み取り）】
  read-range     [excel_file] [range] [--formula]  セル値（--formulaで数式）をテキスト格子で読む
  read-selection [excel_file] [--formula]           今選択している範囲を読む
  sheet-info     [excel_file]                       シート構成・使用範囲の一覧
  screenshot     [excel_file] [range] [--out f]    範囲を画像(PNG)で書き出す

【手コマンド（シートの編集・整形・構造操作／開いたままのブックに直接書込）】
  write-range    [excel_file] <range> [値]          値・数式を書込（グリッドは --tsv / _last_values.tsv）
  clear-range    [excel_file] <range>               範囲をクリア（--contents/--formats/--all）
  format-range   [excel_file] <range> [書式opt...]  フォント・色・罫線・書式・列幅等
  sheet          <add|delete|rename|copy|activate|show|hide|very-hide|visibility|tab-color>
  table          <create|list|delete|column|filter|filter-values|filter-clear|filters|sort|sort-multi|ref>
  name           [excel_file] <add|list|delete>     名前付き範囲

  -- 編集の足回り --
  row            <insert|delete> <行番号> [本数]    行の挿入・削除
  col            <insert|delete> <列文字> [本数]    列の挿入・削除
  copy-range     <src> <dst> [--values]            範囲コピー
  fill           <range> [--right]                  オートフィル（既定は下）
  sort           <range> [--key 列][--desc][--header] 並べ替え
  autofilter     [range] [--off]                    オートフィルタ
  -- 検索・置換 --
  find           <文字> [--book][--whole][--formula] セル検索（番地を返す）
  find-replace   <検索> <置換> [range] [--whole]     一括置換
  -- 保存・印刷 --
  save           [excel_file]                       上書き保存
  save-as        <path>                             別名保存
  print-setup    [--area R][--title-rows 1:3]...    印刷設定
  -- 仕上げ --
  cond-format    <range> --gt 100 --bg '#...'        条件付き書式
  hyperlink      <cell> <url> [--text t]             ハイパーリンク
  validation     <range> --list 'A,B,C'              入力規則(ドロップダウン)
  freeze         <cell> | off                        ウィンドウ枠固定
  comment        <cell> <text>                       セルコメント
  -- 重量級 --
  chart          <create|list|delete>                グラフ（column/bar/line/pie/scatter/area）
  chart-config   <set-title|set-type|legend|style|axis-scale|data-labels|add-series|trendline...>  グラフ詳細設定
  pivot          <create|list|delete>                ピボットテーブル（--rows/--cols/--values/--func）
  pivot-field    <list|add-row|add-col|add-value|remove|set-func|sort|group-date|group-numeric...>  フィールド管理
  pivot-calc     <get-data|calc-field|layout|subtotals|grand-totals>  計算フィールド・レイアウト
  slicer         <add|list|delete>                   スライサー（ピボット/テーブルに紐づけ）
  calc-mode      [manual|auto|recalc]                計算モード確認・切替・再計算
  powerquery     <list|refresh|add|edit|delete|load>  PowerQueryの一覧・更新・作成・書換・削除・読込配線(--to sheet|model)
  connection     <list|refresh|delete> [name]        ブック接続の一覧・更新・削除
  datamodel      <list|relation|measure>             データモデル一覧／リレーション・メジャー(DAX)の作成削除
"""

import sys
import os
import re
import shutil
import zlib
import argparse
import time
import datetime
import unicodedata
import pythoncom
import pywintypes
import win32com.client
import win32com.client.dynamic


from vbam_core import *  # noqa: F401,F403
from vbam_vba import *  # noqa: F401,F403
from vbam_view import *  # noqa: F401,F403
from vbam_edit import *  # noqa: F401,F403
from vbam_heavy import *  # noqa: F401,F403

# ================================================================
# エントリポイント
# ================================================================

def build_parser():
    """argparse の構築（main と batch で共用）"""
    parser = argparse.ArgumentParser(
        description="VBAマネージャー (アクティブブック対応版)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python vba_manager.py list                                      # アクティブブックのマクロ一覧
  python vba_manager.py list 秀.xlsm                              # 指定ファイルのマクロ一覧
  python vba_manager.py list-modules                              # モジュール一覧
  python vba_manager.py get 空白行の削除                          # プロシージャ取得 → _last_proc.vba に保存
  python vba_manager.py get shu001 空白行の削除                   # モジュール指定してプロシージャ取得
  python vba_manager.py get アクティブマクロフォーム.CommandButton2_Click  # ドット区切りでも可
  python vba_manager.py replace-procedure                         # _last_proc.vba の内容で置換
  python vba_manager.py replace-procedure --code-file my.vba
  python vba_manager.py replace-module shu001 shu001_new.bas
  python vba_manager.py export-module shu001                      # shu001.bas にエクスポート
""")

    sub = parser.add_subparsers(dest="command")

    # diag
    sub.add_parser("diag")

    # setup-check（導入セルフ診断・初心者が最初に打つ1コマンド）
    p = sub.add_parser("setup-check", help="導入セルフ診断（Python/pywin32/Excel/VBOM信頼設定を○×表示）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # list-open
    p = sub.add_parser("list-open")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # list [excel_file]
    p = sub.add_parser("list")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--standard", action="store_true", help="標準モジュールのみを抽出")
    p.add_argument("--detail", action="store_true", help="所属モジュール・行数・先頭コメント付きで表示")
    p.add_argument("--module", dest="module_opt", default=None, help="対象モジュールを限定")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--personal", action="store_true", help="個人用マクロブック (PERSONAL.XLSB) を対象にする")
    p.add_argument("--addin", nargs="?", const=True, default=False,
                   help="アドインブック (.xlam/.xla) を対象にする。複数ロード時は名前(一部可)を指定")
    p.add_argument("--all", action="store_true", help="開いているすべてのブック・アドインを対象にする")

    # list-modules [excel_file]
    p = sub.add_parser("list-modules")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--personal", action="store_true", help="個人用マクロブック (PERSONAL.XLSB) を対象にする")
    p.add_argument("--addin", nargs="?", const=True, default=False,
                   help="アドインブック (.xlam/.xla) を対象にする。複数ロード時は名前(一部可)を指定")
    p.add_argument("--all", action="store_true", help="開いているすべてのブック・アドインを対象にする")

    # get [excel_file] <macro_name> [...]
    p = sub.add_parser("get")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="保存先ファイル（省略時は _last_proc.vba。参照用コピーを残したいときに）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # replace-procedure [excel_file] [code_file] [--code-file file] [--module name]
    p = sub.add_parser("replace-procedure")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--code-file", dest="code_file_opt", default=None)
    p.add_argument("--module", dest="module_opt", default=None,
                   help="適用先モジュール名を指定（同名プロシージャが複数ある場合に必須）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes",
                   help="確認プロンプトをスキップして自動で置換を実行します")
    p.add_argument("--force", action="store_true", dest="force",
                   help="構文エラー警告を無視して強制適用します")

    # add-procedure [excel_file] <module_name> [--code-file f] [-y]
    p = sub.add_parser("add-procedure", help="新規プロシージャをモジュール末尾に追加（コードは _last_proc.vba から）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--code-file", dest="code_file_opt", default=None)
    p.add_argument("-y", "--yes", action="store_true", dest="yes",
                   help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force",
                   help="構文エラー警告・バックアップ失敗を無視して強行")

    # add-module [excel_file] <module_name> [--type std|class|form]
    p = sub.add_parser("add-module", help="新規モジュールを追加（標準/クラス/フォーム）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--type", dest="type_opt", default="std",
                   help="モジュール種別 std|class|form（既定 std）")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # delete-procedure [excel_file] <macro_name> [--module name] [-y]
    p = sub.add_parser("delete-procedure", help="プロシージャを削除（削除コードを表示して確認）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module_opt", default=None,
                   help="対象モジュール名（同名が複数ある場合に必須）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes",
                   help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # docs [excel_file] [--out f.md]
    p = sub.add_parser("docs", help="ブックの構成ドキュメント（取説）を自動生成")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力Markdownパス（省略時は _last_docs.md）")
    p.add_argument("--preview", dest="preview", default=None,
                   help="各シートの先頭N行をMarkdown表で含める")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # checkup(健康診断) [excel_file] [--out f.md]
    p = sub.add_parser("checkup", aliases=["健康診断"],
                       help="ブックの健康診断レポート（総合判定+壊れた参照+シート検査+前回との比較）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力Markdownパス（省略時は _last_checkup.md）")
    p.add_argument("--all", dest="all_books", action="store_true",
                   help="開いている全ブックを一括診断（PERSONAL.XLSB含む）")
    p.add_argument("--history", action="store_true",
                   help="診断せず過去の診断履歴（経過観察）を表で表示")
    p.add_argument("--detail", action="store_true",
                   help="--history で各回の間に起きた所見/マクロの増減も表示")
    p.add_argument("--note", default=None,
                   help="今回の診断にカルテのメモを添付（例: --note \"ボタン18を一語修正\"）")
    p.add_argument("--strict", action="store_true",
                   help="所見が1件でもあれば終了コード1（自動化のゲート用。既定は診断完了=0）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--ack-all", dest="ack_all", action="store_true",
                   help="今回のコード/フォーム所見を全て確認済み（意図的）として登録し、"
                        "以降の所見サマリ・総合判定から除外する"
                        "（存在しないマクロ呼び出し等の致命的所見は対象外）")
    p.add_argument("--show-ack", dest="show_ack", action="store_true",
                   help="診断はせず、確認済み（意図的）所見の一覧を表示")
    p.add_argument("--unack", default=None, metavar="文字列",
                   help="部分一致する確認済み所見を確認済みから外す（見直したくなった時）")

    # call-graph [excel_file] [--macro 名]
    p = sub.add_parser("call-graph", help="マクロの呼び出し関係を解析（未解決Call＝一語バグ検出つき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--macro", dest="macro_opt", default=None,
                   help="このマクロを起点に呼び出しツリーを展開")
    p.add_argument("--mermaid", nargs="?", const="_DEFAULT_", default=None,
                   help="Mermaid図をMarkdownに出力（省略時 _last_callgraph.md。GitHub/Qiitaで描画可）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # impact(影響範囲) [excel_file] <マクロ名>
    p = sub.add_parser("impact", aliases=["影響範囲"],
                       help="マクロ修正前の影響範囲予告（呼び元/呼び先を間接まで一覧）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # grep [excel_file] <pattern>
    p = sub.add_parser("grep", help="全モジュール横断のVBAコード検索")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--regex", action="store_true", help="正規表現として検索")
    p.add_argument("-i", "--ignore-case", dest="ignore_case", action="store_true",
                   help="大文字小文字を区別しない")
    p.add_argument("--module", dest="module_opt", default=None, help="検索対象モジュールを限定")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="表示件数の上限（既定200）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # code-replace [excel_file] <検索> <置換>
    p = sub.add_parser("code-replace", help="全マクロ横断の一括置換（diffプレビュー・バックアップ・確認つき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--regex", action="store_true", help="正規表現として置換")
    p.add_argument("--module", dest="module_opt", default=None, help="対象モジュールを限定")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")

    # list-backups [キーワード] / restore <バックアップファイル>
    p = sub.add_parser("list-backups", help="backups のバックアップ一覧（COM不要）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="表示件数の上限（既定30）")
    p = sub.add_parser("restore", help="モジュールバックアップ(.bas/.frm)を開いているブックへ書き戻す")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # list-shortcuts [excel_file]
    p = sub.add_parser("list-shortcuts")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # run-macro [excel_file] <macro_name> [args...]
    p = sub.add_parser("run-macro", help="Excel内の指定されたマクロを実行します")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--json", action="store_true", help="実行結果をJSON形式で出力")
    p.add_argument("--auto-dialog", dest="auto_dialog", default=None,
                   help="実行中に出るMsgBox/InputBoxを自動応答 ok|cancel|yes|no（既定は応答しない）")

    # test [excel_file] [絞り込み] [--module 名] [--auto-dialog ok] [--json]
    p = sub.add_parser("test", aliases=["テスト"],
                       help="テストSub（名前が「テスト」/test で始まる引数なしSub）を一括実行して成否一覧")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module", default=None, help="対象モジュールを限定")
    p.add_argument("--json", action="store_true", help="結果をJSON形式でも出力")
    p.add_argument("--auto-dialog", dest="auto_dialog", default=None,
                   help="テスト中に出るMsgBox等を自動応答 ok|cancel|yes|no")

    # replace-module [excel_file] <module_name> <bas_file>
    p = sub.add_parser("replace-module")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # delete-module [excel_file] <module_name> [-y]
    p = sub.add_parser("delete-module", help="モジュール丸ごと削除（要約表示→確認→バックアップつき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")

    # export-module [excel_file] <module_name>
    p = sub.add_parser("export-module")
    p.add_argument("posargs", nargs="+")

    # export-all [excel_file] [--dir 出力先] [--check]
    p = sub.add_parser("export-all", help="全モジュールを一括エクスポート（1接続）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--dir", dest="dir_opt", default=None, help="出力先フォルダ（省略時はSCRIPTS）")
    p.add_argument("--check", action="store_true", help="書き出した各ファイルに check-bas 相当の検査をかける")

    # reorder-macro <macro_name> <up|down>
    p = sub.add_parser("reorder-macro")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--force", action="store_true",
                   help="バックアップが取れなくても実行する（未保存の新規ブック等）")

    # --- 目コマンド ---
    # read-range [excel_file] [range ...] [--formula] [--tsv [f]] [--width N] [--json]
    p = sub.add_parser("read-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--formula", action="store_true",
                   help="計算結果でなく数式(.Formula)を表示する")
    p.add_argument("--tsv", dest="tsv_out", nargs="?", const="_DEFAULT_", default=None,
                   help="TSVに書き出す（省略時 _last_values.tsv。編集して write-range で書き戻す往復用）")
    p.add_argument("--width", dest="width", default=None,
                   help="列の最大表示幅（既定40。超えた分は…付きで切り詰め）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定。'!'入り・記号入りシート名向け）")

    # read-selection [excel_file] [--formula]
    p = sub.add_parser("read-selection")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--formula", action="store_true",
                   help="計算結果でなく数式(.Formula)を表示する")

    # sheet-info [excel_file] [--preview N]
    p = sub.add_parser("sheet-info")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--preview", dest="preview", default=None,
                   help="各シート使用範囲の先頭N行も表示（ブック俯瞰・1接続）")

    # snapshot [excel_file] [sheet] [--out file] [--sheet NAME] [--max-rows N]
    p = sub.add_parser("snapshot",
                       help="ブック(または1シート)を意味構造JSONに畳む＝開いたままブックLMの下ごしらえ")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力JSONパス（省略時は _last_snapshot.json）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（1シートだけ畳む。posargでも可）")
    p.add_argument("--max-rows", dest="max_rows", default=None,
                   help="1シートあたりのセル読み込み行上限（既定5000。超過は打ち切り注記）")
    p.add_argument("--no-format", dest="no_format", action="store_true",
                   help="書式を採らない（速いが、書式が消えても snapshot-diff で気づけない）")

    # snapshot-diff <before.json> [after.json] [--max N]
    p = sub.add_parser("snapshot-diff",
                       help="2つのsnapshot JSONを比較＝前後差分の検分（COM不要）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--max", dest="max_opt", default=None,
                   help="1分類あたりの表示件数上限（既定20。超過は件数のみ表示）")

    # wiring [excel_file] [--json]
    p = sub.add_parser("wiring", aliases=["配線図"],
                       help="ボタン⇔マクロの配線図（OnAction一覧＋行き先のないボタン検出）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true",
                   help="JSON形式で出力（機械処理用）")

    # screenshot [excel_file] [range] [--out file]
    p = sub.add_parser("screenshot")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力PNGパス（省略時は _last_view.png）")

    # --- 手コマンド (シートの編集・整形・構造操作) ---
    # write-range [excel_file] <range> [値] [--tsv file]
    p = sub.add_parser("write-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--tsv", dest="tsv_opt", default=None,
                   help="グリッドを読み込むTSVファイル（省略時は _last_values.tsv）")
    p.add_argument("--raw", action="store_true",
                   help="数値変換せず文字列として書き込む（セル書式を文字列にする。'007'等の先頭ゼロ保持）")
    p.add_argument("--append", action="store_true",
                   help="使用範囲の最終行の次に書く（rangeは「シート名!列文字」。ログ追記用）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # clear-range [excel_file] <range> [--contents|--formats|--all]
    p = sub.add_parser("clear-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--contents", action="store_true", help="値のみクリア")
    p.add_argument("--formats", action="store_true", help="書式のみクリア")
    p.add_argument("--all", action="store_true", help="すべてクリア（既定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # format-range [excel_file] <range> [書式オプション...]
    p = sub.add_parser("format-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--font")
    p.add_argument("--size")
    p.add_argument("--bold", action="store_true")
    p.add_argument("--unbold", action="store_true")
    p.add_argument("--italic", action="store_true")
    p.add_argument("--color")
    p.add_argument("--bg")
    p.add_argument("--number-format", dest="number_format")
    p.add_argument("--align", choices=['left', 'center', 'right', 'fill', 'justify'])
    p.add_argument("--valign", choices=['top', 'center', 'bottom'])
    p.add_argument("--wrap", action="store_true")
    p.add_argument("--border", choices=['thin', 'medium', 'thick', 'hairline', 'none'])
    p.add_argument("--col-width", dest="col_width")
    p.add_argument("--row-height", dest="row_height")
    p.add_argument("--merge", action="store_true")
    p.add_argument("--unmerge", action="store_true")
    p.add_argument("--autofit", action="store_true")
    p.add_argument("--lock", action="store_true", help="セルをロック（sheet protect 時に有効）")
    p.add_argument("--unlock", action="store_true", help="セルのロック解除（保護中も編集可に）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")

    # sheet [excel_file] <add|delete|rename|copy|activate|show|hide> ...
    p = sub.add_parser("sheet")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--after")
    p.add_argument("--before")
    p.add_argument("--clear", action="store_true", help="tab-color のクリア")
    p.add_argument("--password", default=None, help="protect/unprotect のパスワード")

    # table [excel_file] <create|list|delete> ...
    p = sub.add_parser("table")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--no-headers", dest="no_headers", action="store_true")
    p.add_argument("--at", dest="at", default=None, help="column add の挿入位置(1始まり)")
    p.add_argument("--desc", dest="desc", action="store_true", help="sort を降順に")
    p.add_argument("--tsv", dest="tsv_out", nargs="?", const="_DEFAULT_", default=None,
                   help="table read の結果をTSVに書き出す（省略時 _last_values.tsv）")

    # name [excel_file] <add|list|delete> ...
    p = sub.add_parser("name")
    p.add_argument("posargs", nargs="*")

    # --- 手コマンド 第2弾 ---
    # a. 編集の足回り
    p = sub.add_parser("row")          # row <insert|delete> <行番号> [本数]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet", default=None,
                   help="対象シート名（省略時はアクティブシート）")
    p = sub.add_parser("col")          # col <insert|delete> <列文字> [本数]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet", default=None,
                   help="対象シート名（省略時はアクティブシート）")
    p = sub.add_parser("copy-range")   # copy-range <src> <dst> [--values]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--values", action="store_true", help="値のみ貼り付け")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="コピー元シート名（srcと分離指定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="コピー元にシート名だけ（使用範囲全域）を許可する")
    p = sub.add_parser("fill")         # fill <range> [--right]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--right", action="store_true", help="右方向にフィル（既定は下）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p = sub.add_parser("sort")         # sort <range> [--key 列] [--desc] [--header|--no-header]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--key", help="並べ替えキー列（列文字）")
    p.add_argument("--desc", action="store_true", help="降順")
    p.add_argument("--header", action="store_true", help="先頭行を見出しとして扱う")
    p.add_argument("--no-header", dest="no_header", action="store_true", help="見出しなし")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p = sub.add_parser("autofilter")   # autofilter [range] [--off]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--off", action="store_true", help="オートフィルタを解除")

    # b. 検索・置換
    p = sub.add_parser("find")         # find <文字> [--book] [--whole] [--formula]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--book", action="store_true", help="全シート横断で検索")
    p.add_argument("--whole", action="store_true", help="完全一致")
    p.add_argument("--formula", action="store_true", help="数式も検索対象にする")
    p.add_argument("--max", dest="max_hits", type=int, default=None,
                   help="表示件数の上限（既定200）")
    p = sub.add_parser("find-replace") # find-replace <検索> <置換> [range] [--whole]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--whole", action="store_true", help="完全一致のみ置換")
    p.add_argument("--match-case", dest="match_case", action="store_true",
                   help="大文字小文字を区別する（既定は区別しない）")
    p.add_argument("--wildcard", action="store_true",
                   help="検索文字列の * ? をワイルドカードとして扱う（既定は文字どおり）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # c. 保存・印刷まわり
    p = sub.add_parser("save")         # save [excel_file]
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("save-as")      # save-as [excel_file] <path> [--overwrite]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--overwrite", action="store_true", help="出力先が既存でも上書きする")
    p = sub.add_parser("export-pdf")   # export-pdf <出力.pdf> [--sheet|--range]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet_opt", default=None, help="このシートだけをPDF化")
    p.add_argument("--range", dest="range_opt", default=None, help='この範囲だけをPDF化（例 "集計!A1:H50"）')
    p.add_argument("--overwrite", action="store_true", help="出力先が既存でも上書きする")
    p = sub.add_parser("print-setup")  # print-setup [opts]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--area", help="印刷範囲（例 A1:H50）")
    p.add_argument("--title-rows", dest="title_rows", help="印刷タイトル行（例 1:3）")
    p.add_argument("--title-cols", dest="title_cols", help="印刷タイトル列（例 A:B）")
    p.add_argument("--landscape", action="store_true", help="横向き")
    p.add_argument("--portrait", action="store_true", help="縦向き")
    p.add_argument("--fit-wide", dest="fit_wide", help="横N ページに収める")
    p.add_argument("--fit-tall", dest="fit_tall", help="縦N ページに収める")
    p.add_argument("--zoom", help="拡大縮小率(%%)")
    p.add_argument("--center-h", dest="center_h", action="store_true", help="水平中央")
    p.add_argument("--center-v", dest="center_v", action="store_true", help="垂直中央")

    # d. 仕上げ・見た目
    p = sub.add_parser("cond-format")  # cond-format <range> --gt 100 --bg '#...'
    p.add_argument("posargs", nargs="*")
    p.add_argument("--gt"); p.add_argument("--lt")
    p.add_argument("--ge"); p.add_argument("--le")
    p.add_argument("--eq"); p.add_argument("--ne")
    p.add_argument("--between", nargs=2, metavar=("V1", "V2"))
    p.add_argument("--formula", dest="formula_opt", default=None,
                   help='数式ベースのルール（例 --formula "=B2>AVERAGE($B$2:$B$20)"）')
    p.add_argument("--bg"); p.add_argument("--color")
    p.add_argument("--bold", action="store_true")
    p.add_argument("--clear", action="store_true", help="条件付き書式を全削除")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート全域（シート名だけの指定）を明示的に許可する")
    p = sub.add_parser("hyperlink")    # hyperlink <cell> <url> [--text t] / --remove / --list
    p.add_argument("posargs", nargs="*")
    p.add_argument("--text", help="表示文字")
    p.add_argument("--remove", action="store_true", help="ハイパーリンク削除")
    p.add_argument("--list", dest="list_links", nargs="?", const="__ACTIVE__", default=None,
                   help="シート内の全ハイパーリンクを一覧（--list シート名 で対象指定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="--remove でシート全域（シート名だけの指定）を明示的に許可する")
    p = sub.add_parser("validation")   # validation <range> --list 'A,B,C' / --clear
    p.add_argument("posargs", nargs="*")
    p.add_argument("--list", help="ドロップダウン候補（カンマ区切り）")
    p.add_argument("--clear", action="store_true", help="入力規則を削除")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート全域（シート名だけの指定）を明示的に許可する")
    p = sub.add_parser("freeze")       # freeze <cell> / freeze off
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("comment")      # comment <cell> <text> / --remove
    p.add_argument("posargs", nargs="*")
    p.add_argument("--remove", action="store_true", help="コメント削除")

    # 重量級(1) chart <create|list|delete>
    p = sub.add_parser("chart")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--type", help="column|bar|line|pie|scatter|area|doughnut")
    p.add_argument("--title", help="グラフタイトル")
    p.add_argument("--at", help="左上を合わせるセル")
    p.add_argument("--name", help="グラフ名")
    p.add_argument("--width", help="幅(pt)")
    p.add_argument("--height", help="高さ(pt)")

    # 重量級(1b) chart-config <action> <chart名> ...
    p = sub.add_parser("chart-config")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--min"); p.add_argument("--max")
    p.add_argument("--major"); p.add_argument("--minor")
    p.add_argument("--value", action="store_true"); p.add_argument("--percent", action="store_true")
    p.add_argument("--category", action="store_true"); p.add_argument("--series", action="store_true")
    p.add_argument("--position")
    p.add_argument("--series-name", dest="series_name")
    p.add_argument("--category-range", dest="category_range")
    p.add_argument("--marker-style", dest="marker_style")
    p.add_argument("--marker-size", dest="marker_size")
    p.add_argument("--marker-fg", dest="marker_fg")
    p.add_argument("--marker-bg", dest="marker_bg")
    p.add_argument("--invert", action="store_true")
    p.add_argument("--name")

    # 重量級(2) pivot <create|list|delete>
    p = sub.add_parser("pivot")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--rows", help="行フィールド（カンマ区切り）")
    p.add_argument("--cols", help="列フィールド（カンマ区切り）")
    p.add_argument("--values", help="値フィールド（カンマ区切り）")
    p.add_argument("--func", help="集計方法 sum|count|average|max|min（既定 sum）")
    p.add_argument("--sheet", help="出力シート名（無ければ作成）")
    p.add_argument("--at", help="出力先セル（同シート内に置く場合）")
    p.add_argument("--name", help="ピボットテーブル名")

    # 重量級(2b) pivot-field <action> <pivot> <field> ...
    p = sub.add_parser("pivot-field")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--func", help="add-value/set-func の集計 sum|count|average|max|min")
    p.add_argument("--name", help="add-value の表示名")

    # 重量級(2c) pivot-calc <action> <pivot> ...
    p = sub.add_parser("pivot-calc")
    p.add_argument("posargs", nargs="*")

    # 重量級(3) slicer <add|list|delete>
    p = sub.add_parser("slicer")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--at", help="左上を合わせるセル")
    p.add_argument("--name", help="スライサー名")

    # calc-mode [manual|auto|recalc]
    p = sub.add_parser("calc-mode")
    p.add_argument("posargs", nargs="*")

    # 重量級(4) powerquery <list|refresh|add|delete>
    p = sub.add_parser("powerquery")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--m-file", dest="m_file", default=None, help="add 用 M式ファイル（省略時 _last_query.m）")
    p.add_argument("--m", dest="m_opt", default=None, help="add 用 M式をインライン指定")
    p.add_argument("--desc", default=None, help="クエリの説明")
    p.add_argument("--to", dest="to", default=None, help="load 用 読み込み先: sheet|model")
    p.add_argument("--sheet", dest="sheet", default=None, help="load --to sheet の出力先シート（省略時アクティブ）")
    p.add_argument("--at", dest="at", default=None, help="load --to sheet の左上セル（省略時 A1）")
    p.add_argument("--force", action="store_true",
                   help="load --to model: 既にモデルに載っていても作り直す（メジャー/リレーションは失われる）")

    # 重量級(5) connection <list|refresh|delete> / datamodel [list]
    p = sub.add_parser("connection")
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("datamodel")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--dax", dest="dax", default=None, help="measure add 用 DAX式をインライン指定")
    p.add_argument("--dax-file", dest="dax_file", default=None, help="measure add 用 DAXファイル（省略時 _last_dax.dax）")
    p.add_argument("--desc", dest="desc", default=None, help="measure の説明")
    p.add_argument("--format", dest="format", default=None,
                   help="measure の書式: general|whole|decimal|currency|percent|scientific（既定 general）")
    p.add_argument("--decimals", dest="decimals", default=None, help="小数桁数（decimal/currency/percent/scientific、既定2）")
    p.add_argument("--thousands", dest="thousands", action="store_true", help="桁区切りを使う（whole/decimal/percent）")
    p.add_argument("--symbol", dest="symbol", default=None, help="通貨コード（currency、例: USD/JPY/EUR。グリフ$¥は不可。無効なら既定）")

    # printer-list
    p = sub.add_parser("printer-list")
    p.add_argument("posargs", nargs="*")

    # printer-setup
    p = sub.add_parser("printer-setup")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--printer", help="対象プリンター名（省略時はActivePrinterまたは既定プリンター）")
    p.add_argument("--duplex", choices=['simplex', 'vertical', 'horizontal'], help="両面印刷（simplex:片面, vertical:長辺, horizontal:短辺）")
    p.add_argument("--color", choices=['mono', 'color'], help="カラーモード（mono:モノクロ, color:カラー）")
    p.add_argument("--orientation", choices=['portrait', 'landscape'], help="用紙の向き（portrait:縦, landscape:横）")

    # check [excel_file]
    p = sub.add_parser("check", help="全モジュールの構文チェックと診断を実行します")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # check-bas <file.bas> [--fix]  (COM不要・取り込み前の単体検査)
    p = sub.add_parser("check-bas", help="取り込み前に .bas を単体検査（文字コード/改行二重化/重複）。COM不要・複数可")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--fix", action="store_true", help="改行二重化を CP932 のまま自動修正する")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # batch <コマンドファイル|->  （1接続・1プロセスでコマンド列を実行）
    p = sub.add_parser("batch", help="コマンド列を1回の接続で連続実行（ファイル or 標準入力 '-'）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--keep-going", dest="keep_going", action="store_true",
                   help="途中の失敗で止まらず最後まで実行する")

    # shell（対話セッション。接続を張ったままコマンドを打ち続ける）
    sub.add_parser("shell", help="対話セッション（接続維持のREPL。2コマンド目から再接続なし）")

    return parser


def cmd_batch(args):
    """コマンド列を1プロセス・1COM接続で連続実行: batch <file|->

    各行は通常のCLI引数列そのもの（例: `get shu003 空白行の削除`）。
    空行と # 始まりは無視。get_workbook の接続キャッシュにより全行が同じ
    COM接続を使い回すため、「1コマンド毎の再接続で数分」級の一括作業が
    数秒に縮む。各行の実行は既存コマンドの機械的な再生のみ（判断はしない）。
    """
    import shlex
    src = args.posargs[0] if args.posargs else None
    if not src:
        print("使い方: batch <コマンドファイル|->   （- で標準入力から読む）")
        print("  例: get shu003 マクロA")
        print("      replace-procedure -y")
        return False
    if src == '-':
        # パイプ経由の入力はロケール(CP932)で誤読されるため UTF-8 に固定する（shell と同じ対処）
        if not sys.stdin.isatty():
            try:
                sys.stdin.reconfigure(encoding='utf-8')
            except Exception:
                pass
        text = sys.stdin.read()
    else:
        path = smart_path_resolve(src)
        if not path or not os.path.exists(path):
            print(f"エラー: コマンドファイルが見つかりません: {src}")
            return False
        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                text = f.read()
        except UnicodeDecodeError:
            # メモ帳等で CP932 (Shift-JIS) 保存されたコマンドファイルも受け付ける
            with open(path, 'r', encoding='cp932') as f:
                text = f.read()

    parser = build_parser()
    table = _command_table()
    keep_going = getattr(args, 'keep_going', False)
    total = ok_n = 0
    for lineno, line in enumerate(text.splitlines(), 1):
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        total += 1
        print(f"----- [batch:{lineno}] {line} -----")
        try:
            # Windows パスの \ をエスケープ扱いしない（クォートは通常どおり効く）
            lex = shlex.shlex(line, posix=True)
            lex.whitespace_split = True
            lex.escape = ''
            # shlex 既定のコメント文字 '#' を無効化。行頭 # は上で処理済みで、
            # 行中の # を生かすと「テスト#1」「--bg #FF0000」の # 以降が黙って消える
            lex.commenters = ''
            tokens = list(lex)
        except ValueError as e:
            print(f"[batch:{lineno}] 引数の解析に失敗: {e}")
            if keep_going:
                continue
            print("[batch] 停止（--keep-going で続行可）")
            return False
        try:
            sub_args, unknown = parser.parse_known_args(tokens)
        except SystemExit as e:
            if e.code in (0, None):
                # 行内の -h/--help はヘルプ表示済み。エラーではない
                ok_n += 1
                continue
            print(f"[batch:{lineno}] 引数エラー")
            if keep_going:
                continue
            print("[batch] 停止（--keep-going で続行可）")
            return False
        unknown = [u for u in unknown if u not in ("--visible", "-v")]
        if unknown:
            print(f"[batch:{lineno}] 不明な引数/オプション: {' '.join(unknown)}")
            if keep_going:
                continue
            print("[batch] 停止（--keep-going で続行可）")
            return False
        if not sub_args.command or sub_args.command in ('batch', 'shell'):
            print(f"[batch:{lineno}] このコマンドは batch 内で実行できません")
            if keep_going:
                continue
            return False
        try:
            res = table[sub_args.command](sub_args)
        except SystemExit as e:
            # reorder-macro 等は sys.exit で終了コードを返すため、ここで吸収する
            # （code=None の素の sys.exit() は正常終了。shell 側の判定と揃える）
            res = (e.code in (0, None))
        except Exception as e:
            print(f"[batch:{lineno}] エラー: {e}")
            res = False
        if res is not False:
            ok_n += 1
        elif not keep_going:
            print(f"[batch] {lineno}行目で失敗したため停止（--keep-going で続行可）")
            print(f"===== batch 結果: {ok_n}/{total} 成功 =====")
            return False
    print(f"===== batch 完了: {ok_n}/{total} 成功 =====")
    return ok_n == total


def cmd_shell(args):
    """対話セッション: shell

    接続を張ったままコマンドを打ち続ける REPL。batch のファイル版に対する対話版で、
    get_workbook の接続キャッシュにより2コマンド目からは COM 再接続なしで動く
    （1コマンド約1秒 → 体感即応）。exit / quit / Ctrl+C で終了。
    """
    import shlex
    # パイプ/リダイレクト経由の入力はロケール(CP932)で誤読されるため UTF-8 に固定する
    # （対話（コンソール直打ち）は Windows のコンソールAPIが処理するので触らない）
    if not sys.stdin.isatty():
        try:
            sys.stdin.reconfigure(encoding='utf-8')
        except Exception:
            pass
    parser = build_parser()
    table = _command_table()
    hist_cmds = []
    print("===== vba_manager 対話セッション =====")
    print("  コマンドをそのまま入力（例: list / get マクロ名 / read-range A1:D10）")
    print("  help で使い方、history で履歴（!番号 で再実行・!! は直前）、exit で終了。")
    while True:
        try:
            line = input("vba> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n終了します。")
            break
        if not line or line.startswith('#'):
            continue
        if line.lower() in ('exit', 'quit', 'q'):
            print("終了します。")
            break
        if line.lower() == 'help':
            parser.print_help()
            continue
        if line.lower() == 'history':
            for i, h in enumerate(hist_cmds, 1):
                print(f"  {i}: {h}")
            if not hist_cmds:
                print("  (まだ履歴はありません)")
            continue
        if line.startswith('!'):
            ref = line[1:].strip()
            if ref == '!':
                idx = len(hist_cmds)
            else:
                try:
                    idx = int(ref)
                except ValueError:
                    idx = 0
            if not (1 <= idx <= len(hist_cmds)):
                print("履歴にありません。history で番号を確認して !番号 で再実行（!! は直前）")
                continue
            line = hist_cmds[idx - 1]
            print(f"vba> {line}")
        hist_cmds.append(line)
        try:
            lex = shlex.shlex(line, posix=True)
            lex.whitespace_split = True
            lex.escape = ''                      # Windows パスの \ をエスケープ扱いしない
            lex.commenters = ''                  # 行中の # をコメント扱いしない（#FF0000 等が消える）
            tokens = list(lex)
        except ValueError as e:
            print(f"引数の解析に失敗: {e}")
            continue
        try:
            sub_args, unknown = parser.parse_known_args(tokens)
        except SystemExit:
            continue                             # 引数エラーは argparse が表示済み
        unknown = [u for u in unknown if u not in ("--visible", "-v")]
        if unknown:
            print(f"不明な引数/オプション: {' '.join(unknown)}")
            continue
        if not sub_args.command or sub_args.command in ('shell',):
            print("このコマンドはセッション内で実行できません")
            continue
        try:
            table[sub_args.command](sub_args)
        except SystemExit as e:
            if e.code not in (0, None):
                print(f"（終了コード {e.code}）")
        except KeyboardInterrupt:
            print("（中断しました）")
        except Exception as e:
            print(f"エラー: {e}")
    return True


def _command_table():
    """コマンド名→実装の対応表（main と batch で共用）"""
    return {
        "check":             cmd_check,
        "check-bas":         cmd_check_bas,
        "diag":              cmd_diag,
        "setup-check":       cmd_setup_check,
        "list-open":         cmd_list_open,
        "list":              cmd_list,
        "list-modules":      cmd_list_modules,
        "get":               cmd_get,
        "replace-procedure": cmd_replace_procedure,
        "add-procedure":     cmd_add_procedure,
        "add-module":        cmd_add_module,
        "delete-procedure":  cmd_delete_procedure,
        "grep":              cmd_grep,
        "code-replace":      cmd_code_replace,
        "docs":              cmd_docs,
        "call-graph":        cmd_call_graph,
        "checkup":           cmd_checkup,
        "健康診断":            cmd_checkup,
        "test":              cmd_test,
        "テスト":              cmd_test,
        "impact":            cmd_impact,
        "影響範囲":            cmd_impact,
        "replace-module":    cmd_replace_module,
        "delete-module":     cmd_delete_module,
        "export-module":     cmd_export_module,
        "export-all":        cmd_export_all,
        "list-backups":      cmd_list_backups,
        "restore":           cmd_restore,
        "reorder-macro":     cmd_reorder_macro,
        "list-shortcuts":    cmd_list_shortcuts,
        "run-macro":         cmd_run_macro,
        "read-range":        cmd_read_range,
        "read-selection":    cmd_read_selection,
        "sheet-info":        cmd_sheet_info,
        "snapshot":          cmd_snapshot,
        "snapshot-diff":     cmd_snapshot_diff,
        "wiring":            cmd_wiring,
        "配線図":              cmd_wiring,
        "screenshot":        cmd_screenshot,
        "write-range":       cmd_write_range,
        "clear-range":       cmd_clear_range,
        "format-range":      cmd_format_range,
        "sheet":             cmd_sheet,
        "table":             cmd_table,
        "name":              cmd_name,
        "row":               cmd_row,
        "col":               cmd_col,
        "copy-range":        cmd_copy_range,
        "fill":              cmd_fill,
        "sort":              cmd_sort,
        "autofilter":        cmd_autofilter,
        "find":              cmd_find,
        "find-replace":      cmd_find_replace,
        "save":              cmd_save,
        "save-as":           cmd_save_as,
        "export-pdf":        cmd_export_pdf,
        "print-setup":       cmd_print_setup,
        "cond-format":       cmd_cond_format,
        "hyperlink":         cmd_hyperlink,
        "validation":        cmd_validation,
        "freeze":            cmd_freeze,
        "comment":           cmd_comment,
        "chart":             cmd_chart,
        "chart-config":      cmd_chart_config,
        "pivot-field":       cmd_pivot_field,
        "pivot-calc":        cmd_pivot_calc,
        "pivot":             cmd_pivot,
        "slicer":            cmd_slicer,
        "calc-mode":         cmd_calc_mode,
        "powerquery":        cmd_powerquery,
        "connection":        cmd_connection,
        "datamodel":         cmd_datamodel,
        "printer-list":      cmd_printer_list,
        "printer-setup":     cmd_printer_setup,
        "batch":             cmd_batch,
        "shell":             cmd_shell,
    }


def main():
    setup_encoding()
    parser = build_parser()
    args, unknown = parser.parse_known_args()

    # 未知オプションの黙殺はタイポを事故に変える（例: clear-range --content が
    # 「値のみクリア」でなく既定の全消し Clear() に化ける）。グローバルの
    # --visible/-v だけ許容し、それ以外の残留はエラーで止める。
    unknown = [u for u in unknown if u not in ("--visible", "-v")]
    if unknown:
        print(f"エラー: 不明な引数/オプションです: {' '.join(unknown)}")
        print("  タイプミスの可能性があります。--help で正しいオプションを確認してください。")
        sys.exit(1)

    cmds = _command_table()

    if args.command in cmds:
        ok = False
        try:
            try:
                ok = cmds[args.command](args)
            except SystemExit:
                raise
            except Exception as e:
                print(f"エラー: {e}")
                sys.exit(1)
        finally:
            cleanup_excel()
        # 明示的に False を返したコマンドは失敗(1)、それ以外は成功(0)
        sys.exit(0 if ok is not False else 1)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
