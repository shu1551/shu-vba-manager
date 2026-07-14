# -*- coding: utf-8 -*-
"""vbam_edit.py — vba_manager 分割パート: 「手」コマンド（write/clear/format/sheet/table/検索置換/保存印刷/仕上げ）

vba_manager.py から機械分割（2026-07-12）。単体で実行せず、vba_manager.py 経由で使う。
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
# ================================================================
# 「手」コマンド (シートの編集・整形・構造操作)
#   ※ アクティブ(開いたまま)のブックに COM で直接書き込む。
#      Excel MCP と違いブックを閉じる必要がない代わりに、
#      プログラム経由の変更は Excel の Undo 履歴を消す。
#   ※ 既定では保存しない。Excelで確認後に手動保存するか、
#      保存せず閉じれば変更を破棄できる（=Undo代わりの逃げ道）。
# ================================================================

_LAST_QUERY_FILE  = os.path.join(SCRIPT_DIR, '_last_query.m')       # powerquery add のM式入力
_LAST_DAX_FILE    = os.path.join(SCRIPT_DIR, '_last_dax.dax')       # datamodel measure add のDAX入力

# 配置・罫線の定数 (xl定数の実値)
_XL_ALIGN_H = {'left': -4131, 'center': -4108, 'right': -4152,
               'fill': 5, 'justify': -4130}
_XL_ALIGN_V = {'top': -4160, 'center': -4108, 'bottom': -4107}
_XL_BORDER_WEIGHT = {'hairline': 1, 'thin': 2, 'medium': -4138, 'thick': 4}


def _hex_to_excel_color(hexstr):
    """#RRGGBB / RRGGBB → Excel の BGR 整数"""
    s = hexstr.lstrip('#').strip()
    if len(s) != 6:
        raise ValueError(f"色は #RRGGBB 形式で指定してください: {hexstr}")
    r = int(s[0:2], 16); g = int(s[2:4], 16); b = int(s[4:6], 16)
    return r + g * 256 + b * 65536




class _alerts_off:
    """DisplayAlerts を一時的に切り、抜けるとき「元の値」に戻す with ブロック。

    無条件に True へ戻すと、batch や MCP の 1接続セッションで呼び出し元が意図的に
    False にしていた設定まで書き換えてしまい、後続の操作で確認ダイアログが出て
    無言ハングする（xl.Run はダイアログが閉じるまで戻らない）。変更前の値を覚えて復元する。
    """

    def __init__(self, xl):
        self.xl = xl
        self.prev = True

    def __enter__(self):
        try:
            self.prev = self.xl.DisplayAlerts
        except Exception:
            self.prev = True          # 読めない場合のみ Excel 既定（True）を採用
        self.xl.DisplayAlerts = False
        return self.xl

    def __exit__(self, exc_type, exc, tb):
        try:
            self.xl.DisplayAlerts = self.prev
        except Exception as ex:
            # 戻せないと DisplayAlerts=False のままの Excel が残る＝以後の上書き確認・
            # シート削除確認が全部抑止された状態でユーザーが使い続ける（開いたまま運用では実害）。
            # 保護の再適用失敗と同じく、黙らずに報告する
            print(f"⚠ DisplayAlerts を元に戻せませんでした（{self.prev} に戻す予定でした）: {ex}",
                  file=sys.stderr)
            print("  この Excel では確認ダイアログが抑止されたままです。"
                  "Excel を開き直すか、手動で設定を確認してください。", file=sys.stderr)
        return False


def _read_tsv_grid(path, raw=False):
    """TSV(タブ区切り)をセル値の 2次元タプルに変換（行の長さは不揃いのまま返す）。

    以前は短い行を None で最大列数まで詰めていたが、None 代入は既存セルの
    クリアとして作用し「触れないつもりの右側セル」を消すため、詰め物はしない。
    矩形化の要否は書き込み側（cmd_write_range）が判断する。
    """
    with open(path, 'r', encoding='utf-8-sig') as f:
        text = f.read()
    text = text.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n')
    if text == '':
        return ()
    return tuple(tuple((c if raw else _coerce_cell(c)) for c in line.split('\t'))
                 for line in text.split('\n'))


@protect_safe
def cmd_write_range(args):
    """セル範囲に値・数式を書き込む (read-range の対)"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: write-range [excel_file] <range> [値]")
        print("  単一値はインライン、グリッドは --tsv <file> か _last_values.tsv から読み込み")
        print("  '='始まりは数式として書き込み")
        return False
    spec = rest[0]
    inline_value = rest[1] if len(rest) >= 2 else None
    tsv_opt = getattr(args, 'tsv_opt', None)
    if _reject_extra_args(rest, 2, '使い方: write-range [excel_file] <range> [値]'):
        return False

    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    append = getattr(args, 'append', False)
    if append:
        # --append: spec は「シート名!列文字」または「列文字」。使用範囲の最終行の
        # 次の行から書く（自動の最終行判定を使うため、書き込み先番地を必ず明示表示する）
        # シート名は spec の「!」指定と --sheet のどちらでも指定できる（非append経路と
        # 揃える。--sheet を無視するとアクティブシートに書いて別シートを汚す）
        sheet_part = None
        col_part = spec
        if '!' in spec:
            sheet_part, col_part = spec.split('!', 1)
            if sheet_opt and sheet_opt != sheet_part:
                print(f"エラー: シート指定が食い違っています（'{spec}' と --sheet {sheet_opt}）")
                return False
        elif sheet_opt:
            sheet_part = sheet_opt
        if sheet_part is not None:
            sheet_part = sheet_part.strip("'")
            ws = None
            for sh in wb.Worksheets:
                if sh.Name == sheet_part:
                    ws = sh
                    break
            if ws is None:
                print(f"エラー: シート '{sheet_part}' が見つかりません")
                return False
        else:
            ws = wb.ActiveSheet
        if not re.fullmatch(r'[A-Za-z]{1,3}', col_part or ''):
            print("エラー: --append の range は「シート名!列文字」（例: ログ!A）で指定してください")
            return False
        ur = ws.UsedRange
        next_row = ur.Row + ur.Rows.Count if ur is not None else 1
        # 空シートでも UsedRange は $A$1 を返すため、そのままだと2行目から始まる
        try:
            if (ur is not None and ur.Rows.Count == 1 and ur.Columns.Count == 1
                    and ur.Row == 1 and ur.Column == 1
                    and ur.Cells(1, 1).Value is None):
                next_row = 1
        except Exception:
            pass
        rng = ws.Range(f"{col_part.upper()}{next_row}")
        print(f"追記位置: {ws.Name}!{rng.Address}（使用範囲の最終行の次）")
    else:
        whole = _whole_sheet_spec(wb, spec, sheet_opt)
        if whole is not None:
            print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
            print(f"  全セルが同じ値で上書きされる危険があるため、write-range では")
            print(f"  セル/範囲を明示してください（例: \"{whole}!A1\"）")
            return False
        ws, rng = _resolve_range(xl, wb, spec, sheet_opt)

    raw = getattr(args, 'raw', False)

    # 書き込みは Worksheet_Change 等のイベントマクロを同期発火させる。そのマクロが
    # MsgBox を出すと rng.Value 代入がそこでブロックし、閉じるまでコマンドが無言で
    # ハングする（2026-07-11 実害）。安全解除の監視を書き込みの間だけ常設する。
    _dlg_watcher = _start_dialog_watcher(xl)
    try:
        if inline_value is not None:
            # インライン単一値: 範囲全体に同じ値 (数式可)
            if raw:
                rng.NumberFormat = "@"   # 文字列として保持（"007" 等の先頭ゼロを守る）
                rng.Value = inline_value
            else:
                rng.Value = _coerce_cell(inline_value)
            print(f"書き込み: {ws.Name}!{rng.Address} ← {inline_value}")
        else:
            path = (smart_path_resolve(tsv_opt) if tsv_opt else _LAST_VALUES_FILE)
            if not path or not os.path.exists(path):
                print(f"エラー: TSVが見つかりません: {tsv_opt or _LAST_VALUES_FILE}")
                print("  単一値ならインラインで: write-range A1 \"値\"")
                return False
            grid = _read_tsv_grid(path, raw=raw)
            if not grid:
                print("エラー: TSVが空です")
                return False
            nrows = len(grid)
            lens = {len(r) for r in grid}
            top = ws.Cells(rng.Row, rng.Column)
            if nrows == 1 and len(grid[0]) == 1:
                if raw:
                    top.NumberFormat = "@"
                top.Value = grid[0][0]
                print(f"書き込み: {ws.Name}!{top.Address} ← {grid[0][0]}")
            elif len(lens) == 1:
                ncols = len(grid[0])
                target = ws.Range(top, ws.Cells(rng.Row + nrows - 1,
                                                rng.Column + ncols - 1))
                if raw:
                    target.NumberFormat = "@"
                target.Value = grid
                print(f"書き込み: {ws.Name}!{target.Address} ← TSV {nrows}行 x {ncols}列")
            else:
                # 行の長さが不揃い: 矩形化して None を書くと右側の既存セルが消えるため、
                # 行ごとに実際の長さぶんだけ書き込む
                print(f"⚠ TSVの行の長さが不揃いです（{min(lens)}〜{max(lens)}列）。"
                      "行ごとに書き込み、短い行の右側セルには触れません。")
                for i, row in enumerate(grid):
                    if not row:
                        continue
                    r_tgt = ws.Range(ws.Cells(rng.Row + i, rng.Column),
                                     ws.Cells(rng.Row + i, rng.Column + len(row) - 1))
                    if raw:
                        r_tgt.NumberFormat = "@"
                    r_tgt.Value = (row,)
                print(f"書き込み: {ws.Name}!{top.Address} 起点 ← TSV {nrows}行（不揃い）")
    finally:
        _dlg_watcher.stop()

    note = _dialog_watcher_note(_dlg_watcher, None)
    if note:
        print(note, file=sys.stderr)
    print("（保存はしていません。Excelで確認後に保存してください）")
    return True


@protect_safe
@dialog_safe
def cmd_clear_range(args):
    """セル範囲をクリア (既定: すべて)"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: clear-range [excel_file] <range> [--contents|--formats|--all]")
        return False
    spec = rest[0]
    if _reject_extra_args(rest, 1, '範囲は「シート名!範囲」の単一引数で指定してください'
                                   '（例: clear-range "Sheet1!A1:B2" --contents）'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, spec, sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域を本当にクリアするなら --whole-sheet を付けてください。")
        print(f"  範囲を消すつもりなら「シート名!範囲」で指定してください（例: \"{whole}!A1:B2\"）")
        return False
    ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
    if getattr(args, 'contents', False):
        rng.ClearContents(); what = "値"
    elif getattr(args, 'formats', False):
        rng.ClearFormats(); what = "書式"
    else:
        rng.Clear(); what = "すべて"
    print(f"クリア({what}): {ws.Name}!{rng.Address}")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_format_range(args):
    """セル範囲に書式・整形を適用"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: format-range [excel_file] <range> [オプション...]")
        print("  --font 名 --size N --bold --unbold --italic")
        print("  --color '#RRGGBB' --bg '#RRGGBB' --number-format 書式")
        print("  --align left|center|right --valign top|center|bottom --wrap")
        print("  --border thin|medium|thick|hairline|none")
        print("  --col-width N --row-height N --merge --unmerge --autofit")
        return False
    spec = rest[0]
    if _reject_extra_args(rest, 1, '使い方: format-range [excel_file] <range> [オプション...]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, spec, sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域が同じ書式で塗られ、--merge 併用なら全域結合で値が消える危険があります。")
        print(f"  範囲を明示するか、本当に全域なら --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
    applied = []

    if getattr(args, 'font', None):
        rng.Font.Name = args.font; applied.append(f"font={args.font}")
    if getattr(args, 'size', None):
        rng.Font.Size = float(args.size); applied.append(f"size={args.size}")
    if getattr(args, 'bold', False):
        rng.Font.Bold = True; applied.append("bold")
    if getattr(args, 'unbold', False):
        rng.Font.Bold = False; applied.append("unbold")
    if getattr(args, 'italic', False):
        rng.Font.Italic = True; applied.append("italic")
    if getattr(args, 'color', None):
        rng.Font.Color = _hex_to_excel_color(args.color); applied.append(f"color={args.color}")
    if getattr(args, 'bg', None):
        rng.Interior.Color = _hex_to_excel_color(args.bg); applied.append(f"bg={args.bg}")
    if getattr(args, 'number_format', None):
        rng.NumberFormatLocal = args.number_format; applied.append(f"numfmt={args.number_format}")
    if getattr(args, 'align', None):
        rng.HorizontalAlignment = _XL_ALIGN_H[args.align]; applied.append(f"align={args.align}")
    if getattr(args, 'valign', None):
        rng.VerticalAlignment = _XL_ALIGN_V[args.valign]; applied.append(f"valign={args.valign}")
    if getattr(args, 'wrap', False):
        rng.WrapText = True; applied.append("wrap")
    if getattr(args, 'border', None):
        if args.border == 'none':
            rng.Borders.LineStyle = -4142            # xlNone
            applied.append("border=none")
        else:
            rng.Borders.LineStyle = 1                # xlContinuous
            rng.Borders.Weight = _XL_BORDER_WEIGHT.get(args.border, 2)
            applied.append(f"border={args.border}")
    if getattr(args, 'col_width', None) is not None:
        rng.ColumnWidth = float(args.col_width); applied.append(f"col-width={args.col_width}")
    if getattr(args, 'row_height', None) is not None:
        rng.RowHeight = float(args.row_height); applied.append(f"row-height={args.row_height}")
    if getattr(args, 'merge', False):
        # 複数の値を含む範囲の Merge は Excel が確認ダイアログを出し、
        # CLI が無言で応答待ちブロックする。値の個数を先に数えて、
        # 消える場合は警告した上で DisplayAlerts を切って実行する（左上の値が残る）
        n_vals = 0
        try:
            n_vals = int(xl.WorksheetFunction.CountA(rng))
        except Exception:
            pass
        if n_vals > 1:
            print(f"⚠ 範囲に値が{n_vals}個あります。結合により左上以外の値は消えます。")
        with _alerts_off(xl):
            rng.Merge()
        applied.append("merge")
    if getattr(args, 'unmerge', False):
        rng.UnMerge(); applied.append("unmerge")
    if getattr(args, 'autofit', False):
        rng.Columns.AutoFit(); applied.append("autofit")
    if getattr(args, 'lock', False):
        rng.Locked = True; applied.append("lock（シート保護時に有効）")
    if getattr(args, 'unlock', False):
        rng.Locked = False; applied.append("unlock（シート保護時に有効）")

    if not applied:
        print("書式オプションが指定されていません。--bold --bg '#FFFF00' などを指定してください。")
        return False
    print(f"書式適用: {ws.Name}!{rng.Address}  [{', '.join(applied)}]")
    print("（保存はしていません）")
    return True


@protect_safe
@dialog_safe
def cmd_sheet(args):
    """シート操作: add/delete/rename/copy/activate/show/hide"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: sheet [excel_file] <add|delete|rename|copy|activate|show|hide> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    def find_sheet(name):
        for sh in wb.Sheets:
            if sh.Name == name:
                return sh
        return None

    if action == 'add':
        new_name = rest[1] if len(rest) >= 2 else None
        before = getattr(args, 'before', None)
        after = getattr(args, 'after', None)
        # --before/--after の対象が実在しないと find_sheet が None になり、
        # 無言でアクティブシート手前に追加されてしまうため先に検証する
        if before and find_sheet(before) is None:
            print(f"エラー: --before のシート '{before}' が見つかりません"); return False
        if after and find_sheet(after) is None:
            print(f"エラー: --after のシート '{after}' が見つかりません"); return False
        # 名前重複は Add 後の rename で例外→既定名シートの残骸になるため先に検証
        if new_name and find_sheet(new_name) is not None:
            print(f"エラー: シート '{new_name}' は既に存在します"); return False
        if before:
            sh = wb.Sheets.Add(find_sheet(before))
        elif after:
            sh = wb.Sheets.Add(None, find_sheet(after))
        else:
            sh = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
        if new_name:
            sh.Name = new_name
        print(f"シート追加: {sh.Name}")
    elif action == 'delete':
        if len(rest) < 2:
            print("使い方: sheet delete <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        with _alerts_off(xl):
            sh.Delete()
        print(f"シート削除: {rest[1]}")
    elif action == 'rename':
        if len(rest) < 3:
            print("使い方: sheet rename <old> <new>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        sh.Name = rest[2]
        print(f"シート名変更: {rest[1]} → {rest[2]}")
    elif action == 'copy':
        if len(rest) < 2:
            print("使い方: sheet copy <name> [newname]"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        # 名前重複は Copy 後の rename で例外になり、「Sheet1 (2)」等の複製シートが
        # 残骸として残る（add と同じく先に弾く）
        if len(rest) >= 3 and find_sheet(rest[2]) is not None:
            print(f"エラー: シート '{rest[2]}' は既に存在します"); return False
        sh.Copy(None, sh)
        newsh = wb.ActiveSheet
        if len(rest) >= 3:
            newsh.Name = rest[2]
        print(f"シート複製: {rest[1]} → {newsh.Name}")
    elif action == 'activate':
        if len(rest) < 2:
            print("使い方: sheet activate <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        sh.Activate()
        print(f"アクティブ化: {rest[1]}")
    elif action in ('protect', 'unprotect'):
        if len(rest) < 2:
            print(f"使い方: sheet {action} <name> [--password パスワード]"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        pw = getattr(args, 'password', None)
        if action == 'protect':
            sh.Protect(Password=pw) if pw else sh.Protect()
            print(f"シート保護: {rest[1]}" + ("（パスワード付き）" if pw else ""))
            print("  ※ format-range --lock/--unlock で設定した Locked がここで効きます")
        else:
            try:
                sh.Unprotect(Password=pw) if pw else sh.Unprotect()
            except Exception as e:
                print(f"エラー: 保護解除に失敗しました（パスワード違いの可能性）: {e}")
                return False
            # 保護ガード（protect_safe）は入口で全保護シートを一時解除し、出口で
            # 記録どおり再保護する。このコマンドは「保護を外すこと自体が目的」なので、
            # 記録から落としておかないと出口で保護が戻り、成功表示のまま無言で効かない
            forget_protection(sh)
            print(f"シート保護解除: {rest[1]}")
    elif action in ('show', 'hide', 'very-hide'):
        if len(rest) < 2:
            print(f"使い方: sheet {action} <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        vis = {'show': -1, 'hide': 0, 'very-hide': 2}[action]   # xlVisible/-Hidden/-VeryHidden
        sh.Visible = vis
        label = {'show': '表示', 'hide': '非表示', 'very-hide': '完全非表示(VBAのみ解除可)'}[action]
        print(f"シート{label}: {rest[1]}")
    elif action == 'visibility':
        if len(rest) < 2:
            print("使い方: sheet visibility <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        label = {-1: 'visible（表示）', 0: 'hidden（非表示）',
                 2: 'veryhidden（完全非表示）'}.get(int(sh.Visible), str(sh.Visible))
        print(f"表示状態: {rest[1]} = {label}")
        return True
    elif action == 'tab-color':
        if len(rest) < 2:
            print("使い方: sheet tab-color <name> [#RRGGBB | R G B | --clear]"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        if getattr(args, 'clear', False):
            sh.Tab.ColorIndex = -4142             # xlColorIndexNone
            print(f"タブ色クリア: {rest[1]}")
        elif len(rest) >= 5 and all(x.isdigit() for x in rest[2:5]):
            r, g, b = int(rest[2]), int(rest[3]), int(rest[4])
            sh.Tab.Color = r + g * 256 + b * 65536
            print(f"タブ色設定: {rest[1]} = RGB({r},{g},{b})")
        elif len(rest) >= 3:
            sh.Tab.Color = _hex_to_excel_color(rest[2])
            print(f"タブ色設定: {rest[1]} = {rest[2]}")
        else:
            if int(sh.Tab.ColorIndex) == -4142:
                print(f"タブ色: {rest[1]} = （未設定）")
            else:
                c = int(sh.Tab.Color)
                r = c & 255; g = (c >> 8) & 255; b = (c >> 16) & 255
                print(f"タブ色: {rest[1]} = #{r:02X}{g:02X}{b:02X} (R={r},G={g},B={b})")
            return True
    else:
        print(f"未知のアクション: {action}")
        return False

    print("（保存はしていません）")
    return True


@protect_safe
def cmd_table(args):
    """テーブル(ListObject)操作: create/list/delete"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: table [excel_file] <create|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
            for lo in sh.ListObjects:
                cnt += 1
                print(f"[{sh.Name}] {lo.Name}  範囲={lo.Range.Address}")
        if cnt == 0:
            print("テーブルはありません。")
        return True

    if action == 'create':
        if len(rest) < 2:
            print("使い方: table create <range> [name] [--no-headers]"); return False
        ws, rng = _resolve_range(xl, wb, rest[1])
        has_headers = 2 if getattr(args, 'no_headers', False) else 1  # xlNo=2 / xlYes=1
        lo = ws.ListObjects.Add(1, rng, None, has_headers)             # xlSrcRange=1
        if len(rest) >= 3:
            lo.Name = rest[2]
        print(f"テーブル作成: [{ws.Name}] {lo.Name}  範囲={lo.Range.Address}")
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: table delete <name>"); return False
        name = rest[1]
        for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
            for lo in sh.ListObjects:
                if lo.Name == name:
                    lo.Unlist()                    # テーブル解除 (データは残す)
                    print(f"テーブル解除: [{sh.Name}] {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: テーブル '{name}' が見つかりません")
        return False

    # ---- 以降は <table名> を rest[1] に取る列・フィルタ・ソート操作 ----
    def _find_lo(name):
        for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
            for lo in sh.ListObjects:
                if lo.Name == name:
                    return sh, lo
        return None, None

    def _col_field(lo, col_name):
        """テーブル内の列番号(1始まり)を名前から得る。無ければ None。"""
        for i in range(1, lo.ListColumns.Count + 1):
            if lo.ListColumns.Item(i).Name == col_name:
                return i
        return None

    if action == 'column':
        # table column <add|remove|rename|format> <table> ...
        sub = rest[1].lower() if len(rest) >= 2 else ''
        tname = rest[2] if len(rest) >= 3 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません（table list で確認）。"); return False
        if sub == 'add':
            col_name = rest[3] if len(rest) >= 4 else None
            pos = getattr(args, 'at', None)
            lc = lo.ListColumns.Add(int(pos)) if pos else lo.ListColumns.Add()
            if col_name:
                lc.Name = col_name
            print(f"列追加: {tname}[{lc.Name}]（位置 {lc.Index}）")
            print("（保存はしていません）"); return True
        if sub == 'remove':
            col_name = rest[3] if len(rest) >= 4 else None
            if _col_field(lo, col_name) is None:
                print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
            lo.ListColumns.Item(col_name).Delete()
            print(f"列削除: {tname}[{col_name}]"); print("（保存はしていません）"); return True
        if sub == 'rename':
            if len(rest) < 5:
                print("使い方: table column rename <table> <旧列> <新列>"); return False
            old, new = rest[3], rest[4]
            if _col_field(lo, old) is None:
                print(f"エラー: 列 '{old}' が見つかりません。"); return False
            lo.ListColumns.Item(old).Name = new
            print(f"列名変更: {tname}[{old}] → [{new}]"); print("（保存はしていません）"); return True
        if sub == 'format':
            col_name = rest[3] if len(rest) >= 4 else None
            if _col_field(lo, col_name) is None:
                print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
            lc = lo.ListColumns.Item(col_name)
            if len(rest) >= 5:                       # set
                lc.DataBodyRange.NumberFormat = rest[4]
                print(f"列書式設定: {tname}[{col_name}] = {rest[4]}")
                print("（保存はしていません）"); return True
            else:                                    # get
                try:
                    fmt = lc.DataBodyRange.Cells(1, 1).NumberFormat
                except Exception:
                    fmt = '(取得不可)'
                print(f"列書式: {tname}[{col_name}] = {fmt}"); return True
        print("使い方: table column <add|remove|rename|format> <table> ...")
        return False

    if action in ('filter', 'filter-values', 'filter-clear', 'filters'):
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。"); return False
        if action == 'filter-clear':
            try:
                if lo.AutoFilter is not None:
                    lo.AutoFilter.ShowAllData()
                print(f"フィルタ解除: {tname}")
            except Exception:
                print(f"フィルタは設定されていません: {tname}")
            print("（保存はしていません）"); return True
        if action == 'filters':
            af = lo.AutoFilter
            if af is None:
                print(f"フィルタなし: {tname}"); return True
            print(f"--- {tname} のフィルタ ---")
            any_on = False
            for i in range(1, lo.ListColumns.Count + 1):
                fl = af.Filters.Item(i)
                try:
                    on = fl.On
                except Exception:
                    on = False
                if on:
                    any_on = True
                    try:
                        c1 = fl.Criteria1
                    except Exception:
                        c1 = '(?)'
                    print(f"  {lo.ListColumns.Item(i).Name}: {c1}")
            if not any_on:
                print("  (フィルタ条件なし)")
            return True
        # filter / filter-values は列指定が必要
        col_name = rest[2] if len(rest) >= 3 else None
        field = _col_field(lo, col_name)
        if field is None:
            print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
        if action == 'filter':
            crit = rest[3] if len(rest) >= 4 else None
            if not crit:
                print("使い方: table filter <table> <列> <条件>（例: \">100\" \"=Active\"）"); return False
            lo.Range.AutoFilter(Field=field, Criteria1=crit)
            print(f"フィルタ適用: {tname}[{col_name}] {crit}")
            print("（保存はしていません）"); return True
        if action == 'filter-values':
            vals = rest[3:]
            if not vals:
                print("使い方: table filter-values <table> <列> 値1 値2 ..."); return False
            lo.Range.AutoFilter(Field=field, Criteria1=list(vals), Operator=7)   # xlFilterValues
            print(f"フィルタ適用(値): {tname}[{col_name}] {vals}")
            print("（保存はしていません）"); return True

    if action in ('sort', 'sort-multi'):
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。"); return False
        so = lo.Sort
        so.SortFields.Clear()
        if action == 'sort':
            col_name = rest[2] if len(rest) >= 3 else None
            if _col_field(lo, col_name) is None:
                print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
            order = 2 if getattr(args, 'desc', False) else 1   # xlDescending/xlAscending
            so.SortFields.Add(lo.ListColumns.Item(col_name).Range, 0, order)
            so.Apply()
            print(f"ソート: {tname} {col_name} {'降順' if order == 2 else '昇順'}")
            print("（保存はしていません）"); return True
        else:  # sort-multi  col:asc col:desc ...
            specs = rest[2:]
            if not specs:
                print("使い方: table sort-multi <table> 列:asc 列:desc ..."); return False
            applied = []
            for spec in specs:
                if ':' in spec:
                    cn, od = spec.rsplit(':', 1)
                else:
                    cn, od = spec, 'asc'
                if _col_field(lo, cn) is None:
                    print(f"エラー: 列 '{cn}' が見つかりません。"); return False
                order = 2 if od.lower().startswith('d') else 1
                so.SortFields.Add(lo.ListColumns.Item(cn).Range, 0, order)
                applied.append(f"{cn}{'↓' if order == 2 else '↑'}")
            so.Apply()
            print(f"複数ソート: {tname} {' / '.join(applied)}")
            print("（保存はしていません）"); return True

    if action == 'read':
        # テーブル名で直接読む（従来は table list で番地を得て read-range する2段＝2接続だった）
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。（table list で確認）"); return False
        rng = lo.Range
        print(f"テーブル: {tname}   [{sh.Name}!{rng.Address}]")
        print("=" * 60)
        print(_values_to_grid(rng))
        print("=" * 60)
        tsv_out = getattr(args, 'tsv_out', None)
        if tsv_out is not None:
            path = _LAST_VALUES_FILE if tsv_out == '_DEFAULT_' else os.path.abspath(tsv_out)
            rows = _range_values_2d(rng)
            # セル内改行(Alt+Enter)・タブは TSV の行/列区切りと衝突する。
            # 検査せずに書き戻しコマンドまで案内すると、改行1つで TSV の行数が増え、
            # write-range の「不揃い」経路に落ちて以降の全行が1行ずつ下へズレて
            # 上書きされる（＝無警告のデータ破壊）。read-range 側と同じ警告を出す
            dirty = []
            for ri, r in enumerate(rows):
                for ci, v in enumerate(r):
                    if isinstance(v, str) and ('\n' in v or '\r' in v or '\t' in v):
                        dirty.append(f"{_col_letter(rng.Column + ci)}{rng.Row + ri}")
            if dirty:
                shown = ", ".join(dirty[:8]) + ("" if len(dirty) <= 8 else f" …他{len(dirty) - 8}件")
                print(f"⚠ 警告: セル内に改行/タブを含むセルが {len(dirty)}件あります: {shown}")
                print("  このTSVをそのまま write-range で書き戻すと行・列がずれます。")
                print("  該当セルは手で編集するか、write-range の対象から外してください。")
            with open(path, 'w', encoding='utf-8-sig', newline='') as f:
                f.write('\n'.join('\t'.join(_cell_str(v) for v in r) for r in rows) + '\n')
            print(f"TSV書き出し: {path}  ({len(rows)}行)")
            print(f"  編集後の書き戻し: py vba_manager.py write-range \"{sh.Name}!{_col_letter(rng.Column)}{rng.Row}\"")
            print(f"  （\"007\" 等の先頭ゼロを数値化させたくない場合は --raw を付ける）")
        return True

    if action == 'ref':
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。"); return False
        col_name = rest[2] if len(rest) >= 3 else None
        if col_name:
            print(f"構造化参照: {tname}[{col_name}]")
        else:
            cols = [lo.ListColumns.Item(i).Name for i in range(1, lo.ListColumns.Count + 1)]
            print(f"構造化参照: {tname}[#All] / 列: {', '.join(cols)}")
        return True

    print(f"未知のアクション: {action}")
    return False


@protect_safe
def cmd_name(args):
    """名前付き範囲操作: add/list/delete"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: name [excel_file] <add|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for nm in wb.Names:
            cnt += 1
            try:
                refers = nm.RefersTo
            except Exception:
                refers = '(?)'
            print(f"{nm.Name}  →  {refers}")
        if cnt == 0:
            print("名前付き範囲はありません。")
        return True

    if action == 'add':
        if len(rest) < 3:
            print("使い方: name add <name> <range>"); return False
        nm_name = rest[1]
        ws, rng = _resolve_range(xl, wb, rest[2])
        # rng.Address は既定で絶対参照 ($A$2)。pywin32 ではプロパティなので引数なしで使う
        refers = "='" + ws.Name.replace("'", "''") + "'!" + rng.Address
        wb.Names.Add(nm_name, refers)
        print(f"名前付き範囲を追加: {nm_name} → {refers}")
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: name delete <name>"); return False
        nm_name = rest[1]
        # 完全一致を優先。シートスコープ名（'Sheet1'!名前）の末尾一致は
        # 同名が複数シートにあると最初の1個を消す取り違えになるため、
        # 複数一致ならエラーで止めて候補を出す。
        exact = [nm for nm in wb.Names if nm.Name == nm_name]
        if exact:
            exact[0].Delete()
            print(f"名前付き範囲を削除: {nm_name}")
            print("（保存はしていません）")
            return True
        suffix = [nm for nm in wb.Names if nm.Name.split('!')[-1] == nm_name]
        if len(suffix) == 1:
            actual = suffix[0].Name
            suffix[0].Delete()
            print(f"名前付き範囲を削除: {actual}")
            print("（保存はしていません）")
            return True
        if len(suffix) > 1:
            print(f"エラー: 名前 '{nm_name}' はシート違いで複数あります。完全名で指定してください:")
            for nm in suffix:
                print(f"  {nm.Name}")
            return False
        print(f"エラー: 名前 '{nm_name}' が見つかりません")
        return False

    print(f"未知のアクション: {action}")
    return False


# ================================================================
# 「手」コマンド 第2弾 (編集の足回り / 検索置換 / 保存印刷 / 仕上げ)
#   ※ いずれもアクティブ(開いたまま)のブックに COM で直接作用。
#      save 系を除き既定では保存しない。
# ================================================================

def _col_num(s):
    """列文字(A,B,..,AA) を列番号(1始まり)に変換"""
    n = 0
    for ch in s.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


# ---- a. 編集の足回り ----

def _sheet_or_active(wb, sheet_name):
    """--sheet 指定があればそのシート、なければアクティブシートを返す。
    指定シートが見つからなければエラー表示して None。"""
    if not sheet_name:
        return wb.ActiveSheet
    for sh in wb.Worksheets:
        if sh.Name == sheet_name:
            return sh
    print(f"エラー: シート '{sheet_name}' が見つかりません")
    return None


@protect_safe
@dialog_safe
def cmd_row(args):
    """行の挿入・削除: row <insert|delete> <行番号> [本数]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: row <insert|delete> <行番号> [本数]")
        return False
    action = rest[0].lower()
    start = int(rest[1])
    count = int(rest[2]) if len(rest) >= 3 else 1
    # 本数・行番号の下限検査。count=0 だと Rows("10:9") のような逆順アドレスになり、
    # Excel がこれを正規化するため delete では 9〜10 行＝隣の行まで消える
    # （バッチや計算値から 0 が渡ると無警告でデータが飛ぶ）
    if count < 1:
        print(f"エラー: 本数は1以上で指定してください: {count}")
        return False
    if start < 1:
        print(f"エラー: 行番号は1以上で指定してください: {start}")
        return False
    xl, wb = get_workbook(target_file)
    ws = _sheet_or_active(wb, getattr(args, 'sheet', None))
    if ws is None:
        return False
    if action == 'delete':
        # 破壊操作は実行前に対象を明示する（対象取り違え事故の防止）
        print(f"対象シート: {ws.Name}（{wb.Name}）")
    rng = ws.Rows(f"{start}:{start + count - 1}")
    if action == 'insert':
        rng.Insert()
        print(f"行挿入: {ws.Name} {start}行目に {count}行")
    elif action == 'delete':
        rng.Delete()
        print(f"行削除: {ws.Name} {start}〜{start + count - 1}行")
    else:
        print(f"未知のアクション: {action}（insert|delete）"); return False
    print("（保存はしていません）")
    return True


@protect_safe
@dialog_safe
def cmd_col(args):
    """列の挿入・削除: col <insert|delete> <列文字> [本数]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: col <insert|delete> <列文字> [本数]")
        return False
    action = rest[0].lower()
    start = rest[1]
    # _col_num は英字以外も黙って数値化する。str.isalpha() は '名' のような日本語も
    # True になるため isascii() と併せて弾く（通すと巨大な列番号になり、
    # delete では見当違いの列が消える＝取り返しがつかない）
    if not (start.isascii() and start.isalpha()):
        print(f"エラー: 列は列文字（A〜XFD）で指定してください: '{start}'")
        return False
    count = int(rest[2]) if len(rest) >= 3 else 1
    # 本数の下限検査。count=0 だと Columns("C:B") のような逆順アドレスになり、
    # Excel がこれを正規化するため delete では B〜C 列＝隣の列まで消える
    if count < 1:
        print(f"エラー: 本数は1以上で指定してください: {count}")
        return False
    end = _col_letter(_col_num(start) + count - 1)
    xl, wb = get_workbook(target_file)
    ws = _sheet_or_active(wb, getattr(args, 'sheet', None))
    if ws is None:
        return False
    if action == 'delete':
        # 破壊操作は実行前に対象を明示する（対象取り違え事故の防止）
        print(f"対象シート: {ws.Name}（{wb.Name}）")
    rng = ws.Columns(f"{start}:{end}")
    if action == 'insert':
        rng.Insert()
        print(f"列挿入: {ws.Name} {start}列に {count}列")
    elif action == 'delete':
        rng.Delete()
        print(f"列削除: {ws.Name} {start}〜{end}列")
    else:
        print(f"未知のアクション: {action}（insert|delete）"); return False
    print("（保存はしていません）")
    return True


@protect_safe
@dialog_safe
def cmd_copy_range(args):
    """範囲コピー: copy-range <src> <dst> [--values]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: copy-range <src> <dst> [--values]")
        return False
    if _reject_extra_args(rest, 2, '使い方: copy-range <src> <dst> [--values]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, rest[0], sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: コピー元 '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域が複写先を上書きする危険があるため、範囲を明示するか --whole-sheet を付けてください。")
        return False
    # コピー先も同じガード。シート名だけ渡すと UsedRange 全域が貼り付け先になり、
    # コピー元が単一セル等だと全域タイル上書きになる（2026-07-09 再点検で発見）
    whole_d = _whole_sheet_spec(wb, rest[1])
    if whole_d is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: コピー先 '{rest[1]}' はシート '{whole_d}' の使用範囲全域を指します。")
        print(f"  全域がタイル状に上書きされる危険があるため、貼り付け先の左上セルを明示するか --whole-sheet を付けてください。")
        return False
    ws_s, rng_s = _resolve_range(xl, wb, rest[0], sheet_opt)
    ws_d, rng_d = _resolve_range(xl, wb, rest[1])
    if getattr(args, 'values', False):
        try:
            rng_s.Copy()
            rng_d.PasteSpecial(-4163)          # xlPasteValues
        finally:
            # 失敗しても Excel にコピーの点線（CutCopyMode）を残さない
            try:
                xl.CutCopyMode = False
            except Exception:
                pass
        print(f"コピー(値のみ): {ws_s.Name}!{rng_s.Address} → {ws_d.Name}!{rng_d.Address}")
    else:
        rng_s.Copy(rng_d)
        print(f"コピー(書式・式込): {ws_s.Name}!{rng_s.Address} → {ws_d.Name}!{rng_d.Address}")
    print("（保存はしていません）")
    return True


@protect_safe
@dialog_safe
def cmd_fill(args):
    """オートフィル: fill <range> [--right]（既定は下方向）"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: fill <range> [--right]   範囲の先頭セルを残りに複写")
        return False
    if _reject_extra_args(rest, 1, '使い方: fill <range> [--right]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, rest[0], sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  先頭行/列が全域に複写される危険があるため、範囲を明示するか --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0], sheet_opt)
    if getattr(args, 'right', False):
        rng.FillRight(); direction = "右"
    else:
        rng.FillDown(); direction = "下"
    print(f"フィル({direction}): {ws.Name}!{rng.Address}")
    print("（保存はしていません）")
    return True


@protect_safe
@dialog_safe
def cmd_sort(args):
    """並べ替え: sort <range> [--key 列文字] [--desc] [--header|--no-header]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: sort <range> [--key 列文字] [--desc] [--header|--no-header]")
        return False
    if _reject_extra_args(rest, 1, '使い方: sort <range> [--key 列文字] [--desc] [--header|--no-header]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, rest[0], sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域を並べ替えるなら --whole-sheet を付けてください（--header の明示も推奨）。")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0], sheet_opt)
    keycol = getattr(args, 'key', None)
    if keycol and not (keycol.isascii() and keycol.isalpha()):
        # _col_num は英字以外も黙って数値化してしまい、"A1" が K列扱いになる。
        # なお str.isalpha() は '名' のような日本語も True になるため isascii() が必須
        # （'名' を通すと _col_num が巨大な列番号を返し、意図しない列でソートされる）
        print(f"エラー: --key は列文字（A〜XFD）で指定してください: '{keycol}'")
        return False
    key_idx = _col_num(keycol) if keycol else rng.Column
    keycell = ws.Cells(rng.Row, key_idx)
    order = 2 if getattr(args, 'desc', False) else 1       # xlDescending=2 / xlAscending=1
    if getattr(args, 'header', False):
        header = 1                                          # xlYes
    elif getattr(args, 'no_header', False):
        header = 2                                          # xlNo
    else:
        header = 0                                          # xlGuess
    # Orientation / MatchCase / OrderCustom は Excel が「前回の並べ替え設定」をシートに
    # 保存して引き継ぐ仕様。未指定だと手動の列単位ソート等が引き継がれ、行方向のつもりが
    # 列方向に並べ替わる事故になるため必ず明示する（xlTopToBottom=1）。
    # OrderCustom も同族で、UI で一度「ユーザー設定リスト」（曜日順・部署順など）を
    # 使うとその順が残り、五十音順のつもりがカスタム順で並ぶ（OrderCustom=1＝通常順）
    rng.Sort(Key1=keycell, Order1=order, Header=header,
             Orientation=1, MatchCase=False, OrderCustom=1)
    print(f"並べ替え: {ws.Name}!{rng.Address}  キー列={keycol or _col_letter(rng.Column)}  "
          f"{'降順' if order == 2 else '昇順'}")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_autofilter(args):
    """オートフィルタ: autofilter [range] [--off]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if _reject_extra_args(rest, 1, '使い方: autofilter [range] [--off]'):
        return False
    xl, wb = get_workbook(target_file)
    if getattr(args, 'off', False):
        # 位置引数（範囲/シート名）があればそのシートを解除対象にする。
        # 黙って ActiveSheet に落とすと、指定したつもりの別シートでなく
        # アクティブなシートのフィルタ（絞り込み条件ごと）が消える
        if rest:
            ws, _ = _resolve_range(xl, wb, rest[0])
        else:
            ws = wb.ActiveSheet
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
            print(f"オートフィルタ解除: {ws.Name}")
            print("（保存はしていません）")
        else:
            print(f"オートフィルタは設定されていません: {ws.Name}")
        return True
    spec = rest[0] if rest else None
    ws, rng = _resolve_range(xl, wb, spec)
    if ws.AutoFilterMode:
        print(f"既にオートフィルタが設定されています: {ws.Name}")
    else:
        # win32com の遅延バインディングでは引数なし rng.AutoFilter() が
        # 「AutoFilter メソッドが失敗しました」で落ちる（全省略可能引数を省くとCOMが弾く）。
        # Field:=1 だけ渡すと Criteria なし＝どの列も絞り込まずに範囲全体へフィルタUIを付ける
        # ＝引数なしと同じ「オートフィルタON」になる（実弾で確認済み）。
        rng.AutoFilter(1)
        print(f"オートフィルタ設定: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
    return True


# ---- b. 検索・置換 ----

def cmd_find(args):
    """セル検索: find <文字> [--book] [--whole] [--formula]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: find <文字> [--book(全シート)] [--whole(完全一致)] [--formula(式も検索)]")
        return False
    needle = rest[0]
    xl, wb = get_workbook(target_file)
    # グラフシートには UsedRange が無く例外で検索全体が落ちるため Worksheets 限定
    sheets = list(wb.Worksheets) if getattr(args, 'book', False) else [wb.ActiveSheet]
    look_in = -4123 if getattr(args, 'formula', False) else -4163  # xlFormulas / xlValues
    look_at = 1 if getattr(args, 'whole', False) else 2            # xlWhole / xlPart
    total = 0
    # `or 200` だと --max 0（件数だけ見たい）が偽値で既定に化けるため is None 判定
    mh = getattr(args, 'max_hits', None)
    try:
        max_hits = 200 if mh is None else int(mh)
    except (TypeError, ValueError):
        print("エラー: --max は数値で指定してください")
        return False
    if max_hits < 0:
        print("エラー: --max は 0 以上で指定してください（0 は件数のみ表示）")
        return False
    for ws in sheets:
        try:
            rng = ws.UsedRange
        except Exception:
            # アクティブがグラフシート等だと UsedRange 自体が例外になる
            print(f"（'{ws.Name}' はワークシートではないためスキップ）")
            continue
        try:
            cell = rng.Find(What=needle, LookIn=look_in, LookAt=look_at, MatchCase=False)
        except Exception:
            cell = None
        first = None
        while cell is not None:
            addr = cell.Address
            if first is None:
                first = addr
            elif addr == first:
                break
            total += 1
            if total <= max_hits:
                # 「シート名!$A$1」はそのまま write-range 等の range 引数に貼れる形
                print(f"{ws.Name}!{addr}: {cell.Value}")
            cell = rng.FindNext(cell)
    if total == 0:
        print(f"'{needle}' は見つかりませんでした。")
    else:
        if total > max_hits:
            print(f"…他 {total - max_hits}件（--max で上限変更可）")
        print(f"--- {total}件 ヒット ---")
    return True


@protect_safe
@dialog_safe
def cmd_find_replace(args):
    """一括置換: find-replace <検索> <置換> [range] [--whole] [--wildcard]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: find-replace <検索> <置換> [range] [--whole] [--wildcard]")
        return False
    needle, repl = rest[0], rest[1]
    spec = rest[2] if len(rest) >= 3 else None
    if _reject_extra_args(rest, 3, '使い方: find-replace <検索> <置換> [range] [--whole] [--match-case] [--wildcard]'):
        return False
    # Excel の Find/Replace は * ? を常にワイルドカード解釈する。
    # 素通しすると `find-replace "*" "×"` が全非空セルの丸ごと置換になるため、
    # 既定は ~ エスケープで文字どおりに扱い、--wildcard で明示オプトインする
    if not getattr(args, 'wildcard', False):
        escaped = re.sub(r'([~*?])', r'~\1', needle)
        if escaped != needle:
            print("（* ? ~ は文字どおりに置換します。パターンとして使うなら --wildcard）")
        needle = escaped
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    if spec or sheet_opt:
        ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
    else:
        ws = wb.ActiveSheet
        try:
            rng = ws.UsedRange
        except Exception:
            print(f"エラー: アクティブシート '{ws.Name}' は置換できる"
                  "ワークシートではありません（グラフシート等）。")
            return False
    look_at = 1 if getattr(args, 'whole', False) else 2
    match_case = getattr(args, 'match_case', False)
    # Range.Replace は置換件数を返さないため、置換前にヒットセル数を数える
    # （LookIn は Replace と同じ数式(-4123)で揃える）
    count = 0
    first = None
    # SearchFormat / MatchByte も Sort の Orientation と同じ「省略すると前回値を
    # 引き継ぐ」族。UI で「書式を指定して検索」した後だと Find が書式条件つきで走り、
    # 0件になって「見つかりませんでした（置換なし）」と正常終了する＝無言失敗。
    # MatchByte も残ると日本語シートで半角/全角の区別が勝手に付く
    try:
        xl.FindFormat.Clear()
        xl.ReplaceFormat.Clear()
    except Exception:
        pass
    cell = rng.Find(What=needle, LookAt=look_at, LookIn=-4123, MatchCase=match_case,
                    SearchFormat=False, MatchByte=False)
    while cell is not None:
        addr = cell.Address
        if first is None:
            first = addr
        elif addr == first:
            break
        count += 1
        cell = rng.FindNext(cell)
    if count == 0:
        print(f"'{needle}' は {ws.Name}!{rng.Address} に見つかりませんでした（置換なし）")
        return True
    rng.Replace(What=needle, Replacement=repl, LookAt=look_at, MatchCase=match_case,
                SearchFormat=False, ReplaceFormat=False, MatchByte=False)
    print(f"置換: {ws.Name}!{rng.Address}  '{needle}' → '{repl}'  （{count}セルにヒット）")
    print("（保存はしていません）")
    return True


# ---- c. 保存・印刷まわり ----

@dialog_safe
def cmd_save(args):
    """上書き保存: save [excel_file]"""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    with _alerts_off(xl):
        wb.Save()
    print(f"保存しました: {wb.FullName}")
    return True


def cmd_save_as(args):
    """別名保存: save-as [excel_file] <path>（省略時はアクティブブックを対象）"""
    # 他コマンドと同じ流儀: 引数2つなら第1引数を対象ブック、第2引数を出力パスとする。
    # （以前は rest[0] を無条件に出力パスにしていたため、
    #   `save-as 既存ブック.xlsx 新名.xlsx` で既存ブックが無言上書きされる罠があった）
    rest = list(args.posargs)
    if not rest:
        print("使い方: save-as [excel_file] <出力path> [--overwrite]")
        return False
    if len(rest) >= 3:
        print("エラー: 引数が多すぎます。使い方: save-as [excel_file] <出力path> [--overwrite]")
        return False
    if len(rest) == 2:
        target_file, out_arg = rest[0], rest[1]
    else:
        target_file, out_arg = None, rest[0]
    out = os.path.abspath(out_arg)

    FMT = {'.xlsx': 51, '.xlsm': 52, '.xlsb': 50, '.xls': 56,
           '.csv': 6, '.txt': -4158}
    ext = os.path.splitext(out)[1].lower()
    if ext not in FMT:
        # 未知拡張子を黙って xlsx にフォールバックすると「中身xlsxの .pdf」等の壊れファイルになる
        print(f"エラー: 対応していない拡張子です: '{ext or '(なし)'}'")
        print(f"  対応: {' '.join(sorted(FMT))}")
        return False
    fmt = FMT[ext]

    if os.path.exists(out) and not getattr(args, 'overwrite', False):
        print(f"エラー: 出力先が既に存在します: {out}")
        print("  上書きするなら --overwrite を付けてください。")
        return False

    xl, wb = get_workbook(target_file)
    try:
        old_full = wb.FullName
    except Exception:
        old_full = None
    src_ext = os.path.splitext(wb.Name)[1].lower()
    if src_ext in ('.xlsm', '.xlsb', '.xls') and ext == '.xlsx':
        # DisplayAlerts=False で Excel の警告が出ないため、こちらで明示する
        print("⚠ 注意: マクロ付きブックを .xlsx で保存するため、VBAマクロは保存されません。")
    if ext in ('.csv', '.txt'):
        try:
            n_sheets = wb.Sheets.Count
        except Exception:
            n_sheets = 1
        if n_sheets > 1:
            # これも DisplayAlerts=False で Excel 側の警告が抑止されるため明示する
            print(f"⚠ 注意: {ext} はアクティブシート1枚しか保存されません"
                  f"（このブックは {n_sheets} シート）。")
    with _alerts_off(xl):
        wb.SaveAs(out, FileFormat=fmt)
    # batch/shell/MCP の1接続セッションでは接続キャッシュが旧パスキーのまま残り、
    # 旧名を指定した後続コマンドが改名後のブックに当たる（対象取り違え）。
    # SaveAs 成功時にキャッシュのキーを新パスへ付け替える
    if old_full:
        old_key = old_full.lower()
        if old_key != out.lower() and old_key in _wb_cache:
            _wb_cache[out.lower()] = _wb_cache.pop(old_key)
    print(f"別名保存しました: {out}")
    print("  （以後、開いているブックの保存先はこの新パスになります）")
    return True


def cmd_export_pdf(args):
    """PDF出力: export-pdf [excel_file] <出力.pdf> [--sheet 名 | --range "シート!範囲"]

    ExportAsFixedFormat による出力。既定はブック全体、--sheet で1シート、
    --range で範囲のみ。ブック自体は変更しない（保存フラグも汚さない）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: export-pdf [excel_file] <出力.pdf> [--sheet 名 | --range \"シート!A1:H50\"] [--overwrite]")
        return False
    out = os.path.abspath(rest[0])
    if _reject_extra_args(rest, 1, '出力パスは1つだけ指定してください'):
        return False
    if not out.lower().endswith('.pdf'):
        print(f"エラー: 出力は .pdf で指定してください: {out}")
        return False
    if os.path.exists(out) and not getattr(args, 'overwrite', False):
        print(f"エラー: 出力先が既に存在します: {out}")
        print("  上書きするなら --overwrite を付けてください。")
        return False
    sheet_opt = getattr(args, 'sheet_opt', None)
    range_opt = getattr(args, 'range_opt', None)
    if sheet_opt and range_opt:
        print("エラー: --sheet と --range は同時に指定できません")
        return False

    xl, wb = get_workbook(target_file)
    if range_opt:
        ws, rng = _resolve_range(xl, wb, range_opt)
        rng.ExportAsFixedFormat(0, out)               # 0 = xlTypePDF
        scope = f"範囲 {ws.Name}!{rng.Address}"
    elif sheet_opt:
        ws = None
        for sh in wb.Worksheets:
            if sh.Name == sheet_opt:
                ws = sh
                break
        if ws is None:
            print(f"エラー: シート '{sheet_opt}' が見つかりません")
            return False
        ws.ExportAsFixedFormat(0, out)
        scope = f"シート '{ws.Name}'"
    else:
        wb.ExportAsFixedFormat(0, out)
        scope = "ブック全体"
    print(f"PDF出力: {scope} → {out}")
    return True


@protect_safe
def cmd_print_setup(args):
    """印刷設定: print-setup [--area R] [--title-rows 1:3] [--title-cols A:B] ..."""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    ws = wb.ActiveSheet
    ps = ws.PageSetup
    applied = []

    if getattr(args, 'area', None):
        ps.PrintArea = ws.Range(args.area).Address
        applied.append(f"area={args.area}")
    if getattr(args, 'title_rows', None):
        a, b = (args.title_rows.split(':') + [args.title_rows])[:2]
        ps.PrintTitleRows = f"${a}:${b}"
        applied.append(f"title-rows={args.title_rows}")
    if getattr(args, 'title_cols', None):
        a, b = (args.title_cols.split(':') + [args.title_cols])[:2]
        ps.PrintTitleColumns = f"${a}:${b}"
        applied.append(f"title-cols={args.title_cols}")
    if getattr(args, 'landscape', False):
        ps.Orientation = 2; applied.append("landscape")
    if getattr(args, 'portrait', False):
        ps.Orientation = 1; applied.append("portrait")
    if getattr(args, 'fit_wide', None) is not None:
        ps.Zoom = False; ps.FitToPagesWide = int(args.fit_wide)
        applied.append(f"fit-wide={args.fit_wide}")
    if getattr(args, 'fit_tall', None) is not None:
        ps.Zoom = False; ps.FitToPagesTall = int(args.fit_tall)
        applied.append(f"fit-tall={args.fit_tall}")
    if getattr(args, 'zoom', None) is not None:
        ps.Zoom = int(args.zoom); applied.append(f"zoom={args.zoom}")
    if getattr(args, 'center_h', False):
        ps.CenterHorizontally = True; applied.append("center-h")
    if getattr(args, 'center_v', False):
        ps.CenterVertically = True; applied.append("center-v")

    if not applied:
        print("オプションが指定されていません。--area / --title-rows / --landscape など。")
        return False
    print(f"印刷設定: {ws.Name}  [{', '.join(applied)}]")
    print("（保存はしていません）")
    return True


def cmd_printer_list(args):
    """プリンター一覧およびアクティブプリンターを取得"""
    target_file, _ = parse_target_and_rest(args.posargs)
    
    # Excelのアクティブプリンターを取得
    active_printer = None
    try:
        xl, wb = get_workbook(target_file)
        active_printer = xl.ActivePrinter
        print(f"現在のアクティブプリンター: {active_printer}")
    except Exception as e:
        print(f"警告: Excelからアクティブプリンターを取得できませんでした ({e})")

    # WMI経由でOSにインストールされているプリンター一覧を取得
    import win32com.client
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        printers = wmi.InstancesOf("Win32_Printer")
        print("\nインストールされているプリンター一覧:")
        for printer in printers:
            name = printer.Name
            status = " (選択中)" if active_printer and name in active_printer else ""
            print(f"  - {name}{status}")
    except Exception as e:
        print(f"エラー: インストールされているプリンター一覧を取得できませんでした ({e})")
        return False
    return True


_PRINTER_DUPLEX = {'simplex': 1, 'vertical': 2, 'horizontal': 3}
_PRINTER_COLOR = {'mono': 1, 'color': 2}
_PRINTER_ORIENT = {'portrait': 1, 'landscape': 2}


def cmd_printer_setup(args):
    """プリンターの詳細設定（両面印刷・カラー等）を変更・表示

    ※ 他の「手」コマンドと違い、書き込み先はブックではなく **OS のプリンター設定** そのもの
      （win32print.SetPrinter）。保存せず閉じて破棄する逃げ道が無く、即時・不可逆で
      他のアプリの印刷にも影響する。オプション無しで呼べば現在の構成の表示だけ。
    """
    # 対象プリンター名の決定
    printer_name = getattr(args, 'printer', None)
    if not printer_name:
        # Excelが起動していればそのアクティブプリンター名を使用、さもなくばデフォルトプリンター
        try:
            target_file, _ = parse_target_and_rest(args.posargs)
            xl, wb = get_workbook(target_file)
            raw_printer = xl.ActivePrinter
            if " on " in raw_printer:
                printer_name = raw_printer.split(" on ")[0]
            else:
                printer_name = raw_printer
        except Exception:
            import win32print
            printer_name = win32print.GetDefaultPrinter()

    if not printer_name:
        print("エラー: 対象プリンターが特定できません。--printer で指定してください。")
        return False

    # 辞書に無い値（--duplex bogus 等）は、以前は無言で捨てられ applied が空のまま
    # 「現在のプリンター構成」を表示して成功扱い＝変更依頼が消えていた。先に弾く。
    wants = []                       # [(名前, 設定するDevMode属性, 値)]
    for opt_name, opt_attr, table in (('--duplex', 'Duplex', _PRINTER_DUPLEX),
                                      ('--color', 'Color', _PRINTER_COLOR),
                                      ('--orientation', 'Orientation', _PRINTER_ORIENT)):
        raw = getattr(args, opt_name.lstrip('-').replace('-', '_'), None)
        if not raw:
            continue
        val = table.get(str(raw).lower())
        if val is None:
            print(f"エラー: {opt_name} に指定できない値です: '{raw}'")
            print(f"  受け付ける値: {' | '.join(table)}")
            return False
        wants.append((f"{opt_name.lstrip('-')}={raw}", opt_attr, val))

    print(f"対象プリンター: {printer_name}")
    if wants:
        # 他コマンドの「（保存はしていません）」と真逆＝逃げ道が無いことを明示する
        print("⚠ 注意: これは Windows のプリンター設定そのものを書き換えます"
              "（即時反映・元に戻す操作なし・Excel 以外の印刷にも影響します）。")

    import win32print
    try:
        # 設定変更に必要なアクセス権を指定 (PRINTER_ACCESS_ADMINISTER=4, PRINTER_ACCESS_USE=8)
        access = 4 | 8
        handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": access})
    except Exception:
        try:
            handle = win32print.OpenPrinter(printer_name)
        except Exception as e:
            print(f"エラー: プリンターを開けませんでした ({e})")
            return False

    try:
        info = win32print.GetPrinter(handle, 2)
        devmode = info["pDevMode"]
        if devmode is None:
            print("エラー: プリンターの構成情報 (DevMode) を取得できませんでした。")
            return False

        # 値の妥当性は上で検査済み（未知の値はここへ来ない＝無言で捨てない）
        applied = []
        for label, attr, val in wants:
            setattr(devmode, attr, val)
            applied.append(label)

        if applied:
            win32print.SetPrinter(handle, 2, info, 0)
            print(f"プリンター設定更新完了: [{', '.join(applied)}]")
        else:
            # 現在の設定を表示
            duplex_names = {1: '片面 (simplex)', 2: '両面/長辺綴じ (vertical)', 3: '両面/短辺綴じ (horizontal)'}
            color_names = {1: 'モノクロ (mono)', 2: 'カラー (color)'}
            orient_names = {1: '縦向き (portrait)', 2: '横向き (landscape)'}
            
            d_val = duplex_names.get(devmode.Duplex, f"不明({devmode.Duplex})")
            c_val = color_names.get(devmode.Color, f"不明({devmode.Color})")
            o_val = orient_names.get(devmode.Orientation, f"不明({devmode.Orientation})")
            
            print(f"現在のプリンター構成:")
            print(f"  - 両面印刷: {d_val}")
            print(f"  - カラー　: {c_val}")
            print(f"  - 用紙向き: {o_val}")
            
    except Exception as e:
        print(f"エラー: プリンター設定の変更に失敗しました ({e})")
        return False
    finally:
        win32print.ClosePrinter(handle)
    return True


# ---- d. 仕上げ・見た目 ----

_XL_COND_OP = {'gt': 5, 'lt': 6, 'eq': 3, 'ne': 4, 'ge': 7, 'le': 8, 'between': 1}


@protect_safe
def cmd_cond_format(args):
    """条件付き書式(セルの値): cond-format <range> --gt 100 --bg '#FFC7CE'"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: cond-format <range> [--gt|--lt|--ge|--le|--eq|--ne 値 | --between v1 v2]")
        print("         [--bg '#RRGGBB'] [--color '#RRGGBB'] [--bold] [--clear]")
        return False
    spec = rest[0]
    if _reject_extra_args(rest, 1, '使い方: cond-format [excel_file] <range> [--gt 値 ...|--clear]'):
        return False
    xl, wb = get_workbook(target_file)
    whole = _whole_sheet_spec(wb, spec)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  範囲を明示するか、本当に全域なら --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, spec)

    if getattr(args, 'clear', False):
        rng.FormatConditions.Delete()
        print(f"条件付き書式を全削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True

    formula = getattr(args, 'formula_opt', None)
    if formula:
        # 数式ベースのルール（xlExpression=2）。'=' 始まりの数式が TRUE のセルに書式。
        # 相対参照は範囲の左上セル基準（Excel の条件付き書式と同じ規約）
        if not formula.startswith('='):
            formula = '=' + formula
        # xlExpression は名前付き引数だと DISP_E_PARAMNOTOPTIONAL になるため位置渡し
        # （Operator は式タイプでは無視されるがスロットとして必要。1=xlBetween をダミーに）
        fc = rng.FormatConditions.Add(2, 1, formula)
    else:
        op = None; f1 = None; f2 = None
        for name in ('gt', 'lt', 'eq', 'ne', 'ge', 'le'):
            v = getattr(args, name, None)
            if v is not None:
                op = _XL_COND_OP[name]; f1 = str(v); break
        if op is None and getattr(args, 'between', None):
            op = _XL_COND_OP['between']; f1, f2 = args.between[0], args.between[1]
        if op is None:
            print("比較条件がありません。--gt 100 か --formula \"=数式\" を指定してください。")
            return False

        if f2 is not None:
            fc = rng.FormatConditions.Add(Type=1, Operator=op, Formula1=f1, Formula2=f2)
        else:
            fc = rng.FormatConditions.Add(Type=1, Operator=op, Formula1=f1)

    if getattr(args, 'bg', None):
        fc.Interior.Color = _hex_to_excel_color(args.bg)
    if getattr(args, 'color', None):
        fc.Font.Color = _hex_to_excel_color(args.color)
    if getattr(args, 'bold', False):
        fc.Font.Bold = True
    print(f"条件付き書式を追加: {ws.Name}!{rng.Address}")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_hyperlink(args):
    """ハイパーリンク: hyperlink <cell> <url> [--text 表示文字] / --remove / --list

    url を省略して単セルを指定すると、そのセルのリンクを表示する（取得）。
    --list でシート内の全ハイパーリンクを一覧する。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)

    list_opt = getattr(args, 'list_links', None)
    if list_opt is not None:
        # シート内の全リンク。シート名は --list シート名 か 位置引数、無ければアクティブ
        want = None
        if list_opt != "__ACTIVE__":
            want = list_opt
        elif rest:
            want = rest[0]
        if want:
            ws = None
            for sh in wb.Worksheets:
                if sh.Name == want:
                    ws = sh
                    break
            if ws is None:
                # 黙ってアクティブシートに落とすと「指定シートは0件」と誤読される
                print(f"エラー: シート '{want}' が見つかりません")
                return False
        else:
            ws = wb.ActiveSheet
        cnt = ws.Hyperlinks.Count
        print(f"--- {ws.Name} のハイパーリンク（{cnt}件） ---")
        for i in range(1, cnt + 1):
            hl = ws.Hyperlinks.Item(i)
            try:
                addr = hl.Range.Address
            except Exception:
                addr = '(図形)'
            sub = f"#{hl.SubAddress}" if getattr(hl, 'SubAddress', '') else ''
            print(f"  {ws.Name}!{addr}: {hl.Address or ''}{sub}")
        return True

    if not rest:
        print("使い方: hyperlink <cell> <url> [--text 表示文字]")
        print("       hyperlink <cell>            # そのセルのリンクを表示")
        print("       hyperlink <cell> --remove   # 削除")
        print("       hyperlink --list [シート名]  # シート内の全リンク一覧")
        return False
    if getattr(args, 'remove', False):
        # シート名だけの指定は使用範囲全域のリンク一括削除（書式も戻る）になるためガード
        whole = _whole_sheet_spec(wb, rest[0])
        if whole is not None and not getattr(args, 'whole_sheet', False):
            print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
            print(f"  セル/範囲を明示するか、全リンクを消すなら --whole-sheet を付けてください。")
            return False
    ws, rng = _resolve_range(xl, wb, rest[0])
    if getattr(args, 'remove', False):
        rng.Hyperlinks.Delete()
        print(f"ハイパーリンク削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True
    if len(rest) < 2:
        # 取得モード: そのセルのリンクを表示
        cell = rng.Cells(1, 1)
        if cell.Hyperlinks.Count == 0:
            print(f"{ws.Name}!{cell.Address}: ハイパーリンクはありません")
        else:
            hl = cell.Hyperlinks.Item(1)
            sub = f"#{hl.SubAddress}" if getattr(hl, 'SubAddress', '') else ''
            print(f"{ws.Name}!{cell.Address}: {hl.Address or ''}{sub}"
                  f"  表示=「{cell.Value}」")
        return True
    url = rest[1]
    cell = rng.Cells(1, 1)
    ws.Hyperlinks.Add(Anchor=cell, Address=url)
    # TextToDisplay は環境により効かないので、表示文字は明示的にセル値で上書き
    if getattr(args, 'text', None):
        cell.Value = args.text
    print(f"ハイパーリンク追加: {ws.Name}!{cell.Address} → {url}")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_validation(args):
    """入力規則(ドロップダウン): validation <range> --list 'A,B,C' / --clear"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: validation <range> --list 'A,B,C'  /  validation <range> --clear")
        return False
    if _reject_extra_args(rest, 1, "使い方: validation <range> --list 'A,B,C'  /  validation <range> --clear"):
        return False
    xl, wb = get_workbook(target_file)
    # --clear も設定パス（先に既存規則を Delete する）も破壊的なので、
    # シート名だけの指定＝使用範囲全域は明示なしでは拒否する
    whole = _whole_sheet_spec(wb, rest[0])
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  範囲を明示するか、本当に全域なら --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0])
    if getattr(args, 'clear', False):
        rng.Validation.Delete()
        print(f"入力規則を削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True
    lst = getattr(args, 'list', None)
    if not lst:
        print("--list 'A,B,C' を指定してください。")
        return False
    rng.Validation.Delete()
    rng.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=lst)  # xlValidateList=3
    rng.Validation.InCellDropdown = True
    print(f"入力規則(リスト)を設定: {ws.Name}!{rng.Address}  [{lst}]")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_freeze(args):
    """ウィンドウ枠固定: freeze <cell> / freeze off"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: freeze <cell>（そのセルの左上で固定）  /  freeze off")
        return False
    xl, wb = get_workbook(target_file)
    ws = wb.ActiveSheet
    ws.Activate()
    if rest[0].lower() == 'off':
        xl.ActiveWindow.FreezePanes = False
        print(f"枠固定を解除: {ws.Name}")
    else:
        ws.Range(rest[0]).Select()
        xl.ActiveWindow.FreezePanes = True
        print(f"枠固定: {ws.Name} {rest[0]} の左上で固定")
    print("（保存はしていません）")
    return True


@protect_safe
def cmd_comment(args):
    """セルコメント: comment <cell> <text> / comment <cell> --remove"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: comment <cell> <text>  /  comment <cell> --remove")
        return False
    xl, wb = get_workbook(target_file)
    # シート名だけの指定は使用範囲の左上（A1とは限らない）に黙って命中するため拒否
    whole = _whole_sheet_spec(wb, rest[0])
    if whole is not None:
        print(f"エラー: '{rest[0]}' はセルではなくシート '{whole}' を指します。"
              f"セルを明示してください（例: {whole}!A1）")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0])
    cell = rng.Cells(1, 1)
    if getattr(args, 'remove', False):
        cell.ClearComments()
        print(f"コメント削除: {ws.Name}!{cell.Address}")
        print("（保存はしていません）")
        return True
    if len(rest) < 2:
        # 引数不足のときに既存コメントを消さないよう、ClearComments はチェックの後
        print("使い方: comment <cell> <text>")
        return False
    cell.ClearComments()
    cell.AddComment(rest[1])
    print(f"コメント追加: {ws.Name}!{cell.Address}")
    print("（保存はしていません）")
    return True




__all__ = [
    '_LAST_DAX_FILE',
    '_LAST_QUERY_FILE',
    '_PRINTER_COLOR',
    '_PRINTER_DUPLEX',
    '_PRINTER_ORIENT',
    '_XL_ALIGN_H',
    '_XL_ALIGN_V',
    '_XL_BORDER_WEIGHT',
    '_XL_COND_OP',
    '_alerts_off',
    '_col_num',
    '_hex_to_excel_color',
    '_read_tsv_grid',
    '_sheet_or_active',
    'cmd_autofilter',
    'cmd_clear_range',
    'cmd_col',
    'cmd_comment',
    'cmd_cond_format',
    'cmd_copy_range',
    'cmd_export_pdf',
    'cmd_fill',
    'cmd_find',
    'cmd_find_replace',
    'cmd_format_range',
    'cmd_freeze',
    'cmd_hyperlink',
    'cmd_name',
    'cmd_print_setup',
    'cmd_printer_list',
    'cmd_printer_setup',
    'cmd_row',
    'cmd_save',
    'cmd_save_as',
    'cmd_sheet',
    'cmd_sort',
    'cmd_table',
    'cmd_validation',
    'cmd_write_range',
]
