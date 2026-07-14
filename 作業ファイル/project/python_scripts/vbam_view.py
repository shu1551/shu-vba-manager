# -*- coding: utf-8 -*-
"""vbam_view.py — vba_manager 分割パート: 「目」コマンド（read-range/sheet-info/screenshot/snapshot 等）

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
# ================================================================
# 「目」コマンド (シート状態の読み取り)
# ================================================================

LAST_VIEW_FILE = os.path.join(SCRIPT_DIR, '_last_view.png')   # screenshot の出力先




def _resolve_range(xl, wb, spec, sheet_name=None):
    """
    範囲指定を (ws, rng) に解決する。
      'A1:D20'         → アクティブシートの範囲
      'Sheet1!A1:D20'  → シート指定の範囲
      'Sheet1'         → そのシートの UsedRange
      None / ''        → アクティブシートの UsedRange
    sheet_name（--sheet オプション）が来たら spec はアドレスのみとして扱う。
    「シート!範囲」一本槍だと、'!' を含むシート名（Excelでは合法）や
    記号入り日本語シート名のクォートで詰むための分離指定の口。
    """
    if sheet_name:
        ws = None
        for sh in wb.Sheets:
            if sh.Name == sheet_name:
                ws = sh
                break
        if ws is None:
            raise Exception(f"シート '{sheet_name}' が見つかりません")
        if not spec:
            return ws, ws.UsedRange
        return ws, ws.Range(spec)

    if not spec:
        ws = wb.ActiveSheet
        return ws, ws.UsedRange

    if '!' in spec:
        sheet_part, addr = spec.split('!', 1)
        # Excel の数式バー表記（'月次 集計'!A1）のクォートを剥がす（'' は ' に戻す）
        if len(sheet_part) >= 2 and sheet_part.startswith("'") and sheet_part.endswith("'"):
            sheet_part = sheet_part[1:-1].replace("''", "'")
        ws = wb.Sheets(sheet_part)
        if not addr:
            return ws, ws.UsedRange
        return ws, ws.Range(addr)

    # シート名そのものなら UsedRange
    for sh in wb.Sheets:
        if sh.Name == spec:
            return sh, sh.UsedRange

    # それ以外はアクティブシートのアドレスとして扱う
    ws = wb.ActiveSheet
    return ws, ws.Range(spec)


def _whole_sheet_spec(wb, spec, sheet_name=None):
    """spec がシート全域(UsedRange)に解決される形ならシート名を返す（破壊系コマンドのガード用）。

    「シート名だけ」「末尾!」「空」の spec は _resolve_range で UsedRange 全域になる。
    読み取り系では便利だが、clear/fill/sort/write 等の破壊系では
    範囲指定ミス1つで全域破壊になるため、明示指定(--whole-sheet)なしでは拒否する。
    """
    if sheet_name:
        return sheet_name if not spec else None
    if not spec:
        return wb.ActiveSheet.Name
    if '!' in spec:
        sheet_part, addr = spec.split('!', 1)
        return sheet_part if not addr else None
    for sh in wb.Sheets:
        if sh.Name == spec:
            return spec
    return None






def _disp_width(s):
    """全角文字を2幅として数えた表示幅"""
    w = 0
    for ch in s:
        w += 2 if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else 1
    return w


def _disp_truncate(s, width):
    """表示幅 width に収まるよう切り詰める。

    切れたことが分かるよう末尾に '…' を付ける（黙って切ると、欠けた値を
    全文と誤読して write で書き戻す事故の芽になる）。全文が要るときは
    read-range --width で広げるか --tsv で書き出す。
    """
    if _disp_width(s) <= width:
        return s
    lim = max(width - 2, 1)     # '…' は全角幅2として確保
    out = []
    w = 0
    for ch in s:
        cw = 2 if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else 1
        if w + cw > lim:
            break
        out.append(ch)
        w += cw
    return ''.join(out) + '…'


def _disp_pad(s, width, right=False):
    """表示幅基準で width までスペース埋め（right=Trueで右寄せ）"""
    pad = width - _disp_width(s)
    if pad <= 0:
        return s
    return (' ' * pad + s) if right else (s + ' ' * pad)




def _merged_areas_in_range(rng, cap=8000):
    """range 内の結合セル領域を "A3:I8" 形式のアドレス一覧で返す（重複なし）。

    戻り値は (areas, skipped_total)。
      ・結合が1つも無い範囲は rng.MergeCells が False を返すので即 ([], None)（最速パス）。
      ・cap を超えるセル数の範囲は走査せず (None, total) を返す（巨大UsedRangeの暴走防止）。
    MergeArea の矩形は Python 側の seen 集合に算術で畳み、余分な COM 呼び出しを避ける。
    神エクセルの二重構造（文字はA1・見た目はI列まで結合）を読み解くための素。
    """
    try:
        if rng.MergeCells is False:      # 範囲内に結合が皆無 → 走査不要
            return [], None
    except Exception:
        pass                              # None(混在)や例外は通常走査へ
    try:
        total = int(rng.Cells.Count)
    except Exception:
        return [], None
    if total > cap:
        return None, total
    ws = rng.Worksheet
    r0, c0 = rng.Row, rng.Column
    nr, nc = rng.Rows.Count, rng.Columns.Count
    seen = set()
    areas = []
    for i in range(nr):
        for j in range(nc):
            key = (r0 + i, c0 + j)
            if key in seen:
                continue
            try:
                cell = ws.Cells(r0 + i, c0 + j)
                if cell.MergeCells:
                    ma = cell.MergeArea
                    mr, mc = ma.Row, ma.Column
                    mrc, mcc = ma.Rows.Count, ma.Columns.Count
                    areas.append(
                        f"{_col_letter(mc)}{mr}:{_col_letter(mc + mcc - 1)}{mr + mrc - 1}")
                    for ii in range(mrc):        # 結合矩形を丸ごと走査済みにする
                        for jj in range(mcc):
                            seen.add((mr + ii, mc + jj))
                else:
                    seen.add(key)
            except Exception:
                seen.add(key)
    return areas, None


def _values_to_grid(rng, use_formula=False, max_col_width=40):
    """Range の値を、列文字＋行番号つきのテキスト格子にする

    use_formula=True のときは計算結果ではなく数式(.Formula)を表示する。
    数式のないセルは定数値がそのまま入る（write-range の .Value と同じ規約）。
    max_col_width を超える列は '…' 付きで切り詰める（--width で変更可）。
    """
    raw = rng.Formula if use_formula else rng.Value
    if raw is None:
        return "(空の範囲です)"

    # 単一セル
    if not isinstance(raw, tuple):
        a1 = f"{_col_letter(rng.Column)}{rng.Row}"
        return f"{a1}: {_cell_str(raw)}"

    # tuple-of-tuples へ正規化
    rows = []
    for row in raw:
        rows.append(list(row) if isinstance(row, tuple) else [row])
    if not rows:
        return "(空の範囲です)"

    start_row = rng.Row
    start_col = rng.Column
    ncols = max(len(r) for r in rows)

    str_rows = [[_cell_str(v) for v in row] + [''] * (ncols - len(row)) for row in rows]
    headers = [_col_letter(start_col + j) for j in range(ncols)]

    rownum_w = len(str(start_row + len(str_rows) - 1))
    col_w = []
    for j in range(ncols):
        w = _disp_width(headers[j])
        for i in range(len(str_rows)):
            w = max(w, _disp_width(str_rows[i][j]))
        col_w.append(min(w, max_col_width))

    def fmt_row(cells, label):
        parts = [_disp_pad(label, rownum_w, right=True)]
        for j, c in enumerate(cells):
            parts.append(_disp_pad(_disp_truncate(c, col_w[j]), col_w[j]))
        return ' | '.join(parts)

    out = [fmt_row(headers, ''),
           '-' * (rownum_w + sum(col_w) + 3 * ncols)]
    for i, row in enumerate(str_rows):
        out.append(fmt_row(row, str(start_row + i)))
    return '\n'.join(out)


def cmd_read_range(args):
    """シートのセル値をテキスト格子で読み取る（目・テキスト版）。

    複数範囲可（1回のCOM接続でまとめ読み）。--tsv で _last_values.tsv に
    書き出せば「読む→TSVを編集→write-range で書き戻す」の往復が
    get→_last_proc.vba→replace-procedure と同じ型になる。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    specs = rest if rest else [None]
    use_formula = getattr(args, 'formula', False)
    try:
        # `or 40` だと --width 0 が偽値で既定値に化ける（指定が無言で消える）ため is None 判定
        w_opt = getattr(args, 'width', None)
        width = 40 if w_opt is None else int(w_opt)
    except (TypeError, ValueError):
        print("エラー: --width は数値で指定してください")
        return False
    if width < 1:
        print("エラー: --width は 1 以上で指定してください")
        return False
    tsv_out = getattr(args, 'tsv_out', None)
    if tsv_out is not None and len(specs) > 1:
        print("エラー: --tsv は範囲1つのときだけ使えます")
        return False

    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    blocks = [(_resolve_range(xl, wb, spec, sheet_opt)) for spec in specs]

    if getattr(args, 'json', False):
        import json
        out = []
        for ws, rng in blocks:
            rows = _range_values_2d(rng, use_formula)
            entry = {"sheet": ws.Name, "address": rng.Address,
                     "ref": f"{ws.Name}!{rng.Address}",
                     "rows": [[_cell_str(v) for v in r] for r in rows]}
            areas, skipped = _merged_areas_in_range(rng)
            if skipped:
                entry["merged_skipped_cells"] = skipped   # cap超で未走査
            elif areas:
                entry["merged"] = areas
            out.append(entry)
        print(json.dumps({"success": True, "file": wb.Name, "ranges": out},
                         ensure_ascii=False), file=sys.stdout)
        return True

    for ws, rng in blocks:
        mode = "（数式表示）" if use_formula else ""
        # 末尾の [シート名!番地] はそのまま次コマンドの range 引数に貼れる形
        print(f"シート: {ws.Name}   範囲: {rng.Address}{mode}   [{ws.Name}!{rng.Address}]")
        print("=" * 60)
        print(_values_to_grid(rng, use_formula, max_col_width=width))
        print("=" * 60)
        areas, skipped = _merged_areas_in_range(rng)
        if skipped:
            print(f"結合セル: 未走査（{skipped}セルはcap超・範囲を絞れば表示）")
        elif areas:
            shown = areas[:30]
            more = f"  …他{len(areas) - 30}件" if len(areas) > 30 else ""
            print(f"結合セル {len(areas)}件: " + ", ".join(shown) + more)

    if tsv_out is not None:
        path = _LAST_VALUES_FILE if tsv_out == '_DEFAULT_' else os.path.abspath(tsv_out)
        ws, rng = blocks[0]
        rows = _range_values_2d(rng, use_formula)
        # セル内改行(Alt+Enter)・タブは TSV の行/列区切りと衝突し、
        # そのまま write-range で書き戻すと格子がずれて無警告のデータ破壊になる。
        # 値は変えず、該当セルを名指しで警告する
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
        lines = ['\t'.join(_cell_str(v) for v in r) for r in rows]
        with open(path, 'w', encoding='utf-8-sig', newline='') as f:
            f.write('\n'.join(lines) + '\n')
        print(f"TSV書き出し: {path}  ({len(rows)}行 x {max(len(r) for r in rows)}列)")
        print(f"  編集後の書き戻し: py vba_manager.py write-range \"{ws.Name}!{_col_letter(rng.Column)}{rng.Row}\"")
        print(f"  （\"007\" 等の先頭ゼロを数値化させたくない場合は --raw を付ける）")
    return True


def cmd_read_selection(args):
    """ユーザーが今選択している範囲を読み取る"""
    target_file, _ = parse_target_and_rest(args.posargs)
    use_formula = getattr(args, 'formula', False)
    xl, wb = get_workbook(target_file)
    sel = xl.Selection
    if sel is None:
        print("選択範囲がありません。")
        return False
    try:
        wsname = sel.Worksheet.Name
    except Exception:
        wsname = "(不明)"
    try:
        addr = sel.Address
    except Exception:
        addr = "(範囲ではありません)"
    mode = "（数式表示）" if use_formula else ""
    print(f"シート: {wsname}   選択範囲: {addr}{mode}")
    print("=" * 60)
    try:
        print(_values_to_grid(sel, use_formula))
    except Exception as e:
        print(f"(値を読めませんでした: {e})")
    print("=" * 60)
    return True


def cmd_sheet_info(args):
    """ブックのシート構成・使用範囲を表示（見取り図）。

    --preview N で各シート使用範囲の先頭N行も格子表示（初見ブックの俯瞰が
    1コマンド1接続で済む。従来は sheet-info + シート毎の read-range で N+1 接続）。
    """
    target_file, _ = parse_target_and_rest(args.posargs)
    try:
        preview = int(getattr(args, 'preview', None) or 0)
    except (TypeError, ValueError):
        print("エラー: --preview は数値で指定してください")
        return False
    xl, wb = get_workbook(target_file)
    active = wb.ActiveSheet.Name
    print(f"ブック: {wb.Name}")
    print(f"シート数: {wb.Sheets.Count}   アクティブ: {active}")
    print("-" * 60)
    for sh in wb.Sheets:
        mark = '*' if sh.Name == active else ' '
        try:
            ur = sh.UsedRange
            dims = f"{ur.Rows.Count}行 x {ur.Columns.Count}列  ({ur.Address})"
        except Exception:
            ur = None
            dims = "(空)"
        vis = '' if sh.Visible == -1 else '  [非表示]'
        print(f"{mark} {sh.Name}: {dims}{vis}")
        if ur is not None:
            areas, skipped = _merged_areas_in_range(ur)
            if skipped:
                print(f"    結合: 未走査（{skipped}セル・大）")
            elif areas:
                shown = areas[:12]
                more = f"  …他{len(areas) - 12}件" if len(areas) > 12 else ""
                print(f"    結合 {len(areas)}件: " + ", ".join(shown) + more)
        if preview > 0 and ur is not None:
            try:
                nrows = min(preview, ur.Rows.Count)
                head = ur.Worksheet.Range(ur.Cells(1, 1), ur.Cells(nrows, ur.Columns.Count))
                grid = _values_to_grid(head)
                print('    ' + grid.replace('\n', '\n    '))
            except Exception as e:
                print(f"    (先頭行を読めませんでした: {e})")
    print("-" * 60)
    return True


def _shape_text(shp):
    """図形の表示文字を複数方式で拾う。

    Formsコントロール(ボタン等)は TextFrame.Characters().Text、
    AutoShape等は TextFrame2.TextRange.Text に文字が入る（相手で口が違う）。
    先に非空を返した方を採用。どちらも取れなければ None。
    """
    for getter in (
        lambda: shp.TextFrame.Characters().Text,
        lambda: shp.TextFrame2.TextRange.Text,
    ):
        try:
            t = getter()
            if t:
                return t
        except Exception:
            continue
    return None


def _collect_shapes(shapes, out, in_group=None):
    """Shapes を平坦なリストに集める。グループ(msoGroup=6)は中のボタンも展開して拾う。

    ファイル一覧等ではボタンがグループにまとめられており、展開しないと
    中の「抽出開始」「選択抽出」等が丸ごと落ちる（実測で判明）。
    """
    for shp in shapes:
        s = {}
        try:
            s["name"] = shp.Name
        except Exception:
            pass
        try:
            typ = int(shp.Type)
            s["type"] = typ
        except Exception:
            typ = None
        txt = _shape_text(shp)
        if txt:
            s["text"] = txt
        try:
            s["l"] = round(float(shp.Left))
            s["t"] = round(float(shp.Top))
        except Exception:
            pass
        try:
            oa = shp.OnAction
            if oa:
                s["onaction"] = oa.split('!')[-1].strip("'\" ")
                if '!' in oa:
                    # 「'ブック名'!マクロ名」のブック修飾。剥がして捨てると
                    # アドイン先の配線を wiring が「存在しない」と誤判定する
                    book = oa.rsplit('!', 1)[0].strip("'\" ")
                    if book:
                        s["onaction_book"] = book
        except Exception:
            pass
        if in_group:
            s["group"] = in_group
        out.append(s)
        if typ == 6:                       # msoGroup → 中身を1段展開
            try:
                _collect_shapes(shp.GroupItems, out, in_group=s.get("name"))
            except Exception:
                pass


def cmd_snapshot(args):
    """アクティブブック(または1シート)を意味構造JSONに畳む＝開いたままブックLMの下ごしらえ。

    セル(疎)＋結合(merged)＋図形/ボタン(text・座標・OnAction)＋テーブル(ListObject)を
    1ファイルに束ねる。晴美さんのExStruct extract を「開いてるブック相手・その場」で焼き直したもの。
    表かどうかの"推定"はしない（機械的事実だけ吐き、意味付けは読み手のAIがやる＝この子の設計思想）。
    """
    import json
    target_file, rest = parse_target_and_rest(args.posargs)
    only_sheet = rest[0] if rest else getattr(args, 'sheet_opt', None)
    out_path = os.path.abspath(getattr(args, 'out_opt', None) or _LAST_SNAPSHOT_FILE)
    try:
        max_rows = int(getattr(args, 'max_rows', None) or 5000)
    except (TypeError, ValueError):
        print("エラー: --max-rows は数値で指定してください")
        return False

    xl, wb = get_workbook(target_file)
    active = wb.ActiveSheet.Name

    if only_sheet:
        targets = [sh for sh in wb.Sheets if sh.Name == only_sheet]
        if not targets:
            print(f"シート '{only_sheet}' が見つかりません")
            return False
    else:
        targets = list(wb.Sheets)

    sheets = {}
    for sh in targets:
        info = {}
        try:
            ur = sh.UsedRange
            nr, nc = ur.Rows.Count, ur.Columns.Count
            info["dims"] = f"{nr}行 x {nc}列"
            info["used"] = ur.Address
        except Exception:
            ur, nr, nc = None, 0, 0
            info["dims"] = "(空)"
        info["visible"] = (sh.Visible == -1)

        # --- セル（疎：空セル・空行は落とす。読み込みも max_rows 行までに絞る）---
        cells = []
        if ur is not None and nr > 0:
            r0, c0 = ur.Row, ur.Column
            read_rows = min(nr, max_rows)
            try:
                sub = sh.Range(ur.Cells(1, 1), ur.Cells(read_rows, nc))
                raw = sub.Value
            except Exception:
                raw = None
            if raw is None:
                grid = []
            elif not isinstance(raw, tuple):
                grid = [(raw,)]
            else:
                grid = [r if isinstance(r, tuple) else (r,) for r in raw]
            for i, row in enumerate(grid):
                cmap = {}
                for j, v in enumerate(row):
                    if v is None or v == '':
                        continue
                    cmap[_col_letter(c0 + j)] = _cell_str(v)
                if cmap:
                    cells.append({"r": r0 + i, "c": cmap})
            if nr > max_rows:
                # first_row / last_read_row は「絶対行番号」。セルは r0+i の絶対行で
                # 格納されるので、read_rows（＝読んだ行数）だけを持たせると
                # snapshot-diff 側が行数と絶対行番号を取り違える
                # （UsedRange が1行目から始まらないシートで、読み込み済みの実差分まで
                #   捨てて「差分なし」と嘘をつく。2026-07-14 実弾で確認）
                info["cells_truncated"] = {
                    "read_rows": max_rows, "total_rows": nr,
                    "first_row": r0, "last_read_row": r0 + read_rows - 1,
                }
        info["cells"] = cells

        # --- 結合（①の素を流用）---
        if ur is not None:
            areas, skipped = _merged_areas_in_range(ur)
            if skipped:
                info["merged_skipped_cells"] = skipped
            elif areas:
                info["merged"] = areas

        # --- 図形／ボタン（text・座標・実行マクロ。グループは1段展開）---
        shapes = []
        try:
            _collect_shapes(sh.Shapes, shapes)
        except Exception as e:
            # 黙って落とすと「図形なし」と見分けがつかず、後の snapshot-diff で
            # 図形の消失/追加を取り違える。読めなかった事実を JSON とサマリに残す
            info["shapes_error"] = str(e)
        if shapes:
            info["shapes"] = shapes

        # --- テーブル（正式な ListObject だけ・推定はしない）---
        tables = []
        try:
            for lo in sh.ListObjects:
                tables.append({"name": lo.Name, "address": lo.Range.Address})
        except Exception:
            pass
        if tables:
            info["tables"] = tables

        sheets[sh.Name] = info

    doc = {"success": True, "book": wb.Name, "active_sheet": active, "sheets": sheets}
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)

    # 端末にはサマリだけ（本体JSONは大きくなり得るのでファイルへ）
    print(f"スナップショット: {wb.Name}  → {out_path}")
    print("-" * 60)
    for name, info in sheets.items():
        parts = [info.get("dims", "")]
        parts.append(f"セル{len(info.get('cells', []))}行")
        if "cells_truncated" in info:
            parts.append(f"(全{info['cells_truncated']['total_rows']}行→{info['cells_truncated']['read_rows']}打切)")
        if "merged" in info:
            parts.append(f"結合{len(info['merged'])}")
        elif "merged_skipped_cells" in info:
            parts.append("結合未走査(大)")
        if "shapes" in info:
            parts.append(f"図形{len(info['shapes'])}")
        if "shapes_error" in info:
            parts.append("図形読取不可")
        if "tables" in info:
            parts.append(f"表{len(info['tables'])}")
        mark = '*' if name == active else ' '
        print(f"{mark} {name}: " + "  ".join(p for p in parts if p))
    print("-" * 60)
    # 図形を列挙できなかったシートは警告として明示する（「図形0」に見せない）
    shape_err = [(n, i["shapes_error"]) for n, i in sheets.items() if "shapes_error" in i]
    if shape_err:
        print("⚠ 次のシートは図形を読み取れませんでした（このJSONでは図形なしと区別できません）:")
        for n, e in shape_err:
            print(f"   {n}: {e}")
        print("-" * 60)
    print("このJSONを read して質問すれば、開いたままブックLMになる。")
    return True


def _sheet_first_row(info):
    """シートの UsedRange 先頭行（絶対行番号）。取れなければ None。

    新しい snapshot は cells_truncated.first_row を持つ。古い JSON にはないので
    used のアドレス（"$A$200:$C$699" 等）から復元する。
    """
    t = info.get('cells_truncated') or {}
    fr = t.get('first_row')
    try:
        fr = int(fr)
        if fr > 0:
            return fr
    except (TypeError, ValueError):
        pass
    m = re.search(r'\$?[A-Z]{1,3}\$?(\d+)', str(info.get('used') or ''))
    return int(m.group(1)) if m else None


def _cell_compare_limit(old_info, new_info):
    """snapshot 2つのシート情報から「セルを比較してよい行の上限（絶対行番号）」を返す。

    snapshot は --max-rows で読み込む行を打ち切る（cells_truncated に記録）。
    打ち切りは「そこから下は読んでいない」だけで「空だった」ではないため、
    打ち切り行より下を比較すると片側だけ空＝疑似差分になる。

    ※ 返すのは「絶対行番号」であって行数ではない。セルは r0+i の絶対行で格納
    されるため、read_rows（読んだ行数）をそのまま上限にすると単位が食い違い、
    UsedRange が1行目から始まらないシートでは読み込み済みの実差分まで捨てて
    「差分なし」と嘘をつく（2026-07-14 実弾で確認）。
    上限は last_read_row（＝first_row + read_rows - 1）で判定する。
    """
    limits = []
    for info in (old_info, new_info):
        t = info.get('cells_truncated') or {}
        if not t:
            continue
        last = t.get('last_read_row')
        try:
            last = int(last)
        except (TypeError, ValueError):
            last = None
        if last is None:
            # 古い snapshot（last_read_row を持たない）は used から復元する
            fr = _sheet_first_row(info)
            try:
                rr = int(t.get('read_rows'))
            except (TypeError, ValueError):
                rr = None
            last = (fr + rr - 1) if (fr is not None and rr) else None
        if last is not None and last > 0:
            limits.append(last)
    return min(limits) if limits else None


def cmd_snapshot_diff(args):
    """2つの snapshot JSON を機械的に比較: snapshot-diff <before.json> [after.json]

    checkup を前後に挟む型の「シート側」版＝マクロや手作業が実際に何を変えたかを
    セル・結合・図形・テーブル単位の事実で示す（COM 不要の純粋処理）。
    after 省略時は _last_snapshot.json（直近の snapshot）と比較する。
    評価はしない＝差分という事実だけ並べ、意味付けは読み手がやる（snapshot と同じ思想）。
    """
    import json
    if not args.posargs:
        print("使い方: snapshot-diff <before.json> [after.json]")
        print("  after 省略時は _last_snapshot.json（直近の snapshot）と比較します")
        return False
    old_path = os.path.abspath(args.posargs[0])
    new_path = (os.path.abspath(args.posargs[1]) if len(args.posargs) >= 2
                else _LAST_SNAPSHOT_FILE)
    try:
        # `or 20` だと --max 0（件数だけ見たい）が偽値で既定に化けるため is None 判定
        m_opt = getattr(args, 'max_opt', None)
        max_show = 20 if m_opt is None else int(m_opt)
    except (TypeError, ValueError):
        print("エラー: --max は数値で指定してください")
        return False
    if max_show < 0:
        print("エラー: --max は 0 以上で指定してください（0 は件数のみ表示）")
        return False

    docs = []
    for path in (old_path, new_path):
        if not os.path.exists(path):
            print(f"エラー: ファイルがありません: {path}")
            return False
        try:
            # utf-8-sig: PowerShell の Out-File 等が付ける BOM も受け入れる（BOM無しも可）
            with open(path, 'r', encoding='utf-8-sig') as f:
                d = json.load(f)
        except Exception as e:
            print(f"エラー: JSON を読めません: {path} ({e})")
            return False
        if not isinstance(d, dict) or 'sheets' not in d:
            print(f"エラー: snapshot 形式ではありません（'sheets' がない）: {path}")
            return False
        docs.append(d)
    old, new = docs

    def clip(v, n=40):
        s = str(v).replace('\r', ' ').replace('\n', ' ')
        return s if len(s) <= n else s[:n] + '…'

    def cell_map(info):
        out = {}
        for row in info.get('cells', ()):
            for col, v in row.get('c', {}).items():
                out[(row['r'], col)] = v
        return out

    def addr(key):
        return f"{key[1]}{key[0]}"

    def cell_order(key):
        return (key[0], len(key[1]), key[1])   # 行→列文字（桁→辞書順）で安定表示

    def show(label, items, fmt):
        print(f"  {label}: {len(items)}件")
        for it in items[:max_show]:
            print(f"    {fmt(it)}")
        if len(items) > max_show:
            print(f"    … 他 {len(items) - max_show}件（--max で表示数変更可）")

    print("===== スナップショット差分 =====")
    print(f"  旧: {old.get('book', '?')}  ({old_path})")
    print(f"  新: {new.get('book', '?')}  ({new_path})")
    if old.get('book') != new.get('book'):
        print("  ※ ブック名が異なります（別ブック同士の比較）")

    old_sheets, new_sheets = old['sheets'], new['sheets']
    diff_sheets = 0
    merged_note_hidden = []   # 結合未走査で比較できず、かつ他に差分がなく非表示のシート
    trunc_note_hidden = []    # セル打ち切りで一部しか比較できず、かつ他に差分がなく非表示のシート

    for name in [n for n in old_sheets if n not in new_sheets]:
        diff_sheets += 1
        print(f"\n- シート削除: {name}")
    for name in [n for n in new_sheets if n not in old_sheets]:
        diff_sheets += 1
        info = new_sheets[name]
        print(f"\n+ シート追加: {name}  ({info.get('dims', '?')})")

    for name in [n for n in old_sheets if n in new_sheets]:
        oi, ni = old_sheets[name], new_sheets[name]

        oc, nc_ = cell_map(oi), cell_map(ni)
        # 片側だけ --max-rows で打ち切った snapshot をそのまま比べると、打ち切り行より
        # 下は「片側だけ空」＝実際は無変更なのに丸ごとセル追加/削除に化ける。
        # 結合（merged_skipped_cells）で比較を降りるのと同じ扱いで、比較する行を
        # 両者の read_rows の小さい方までに切り詰める（切り詰めた事実は必ず表示する）。
        cell_limit = _cell_compare_limit(oi, ni)
        if cell_limit is not None:
            oc = {k: v for k, v in oc.items() if k[0] <= cell_limit}
            nc_ = {k: v for k, v in nc_.items() if k[0] <= cell_limit}
        added = sorted([k for k in nc_ if k not in oc], key=cell_order)
        removed = sorted([k for k in oc if k not in nc_], key=cell_order)
        changed = sorted([k for k in oc if k in nc_ and oc[k] != nc_[k]],
                         key=cell_order)

        merged_unscanned = bool(oi.get('merged_skipped_cells')
                                or ni.get('merged_skipped_cells'))
        if merged_unscanned:
            # 片側でも結合未走査なら差分は出せない（空 vs 実データの疑似差分を出さない）
            m_add, m_del = [], []
        else:
            om = set(oi.get('merged') or [])
            nm = set(ni.get('merged') or [])
            m_add, m_del = sorted(nm - om), sorted(om - nm)

        def shape_map(info):
            out, dup = {}, set()
            for s in info.get('shapes', ()):
                nm2 = s.get('name', '(無名)')
                if nm2 in out:
                    dup.add(nm2)          # Excel は図形名の重複を許す＝黙って落とさず注記する
                out.setdefault(nm2, s)
            return out, dup
        os_, odup = shape_map(oi)
        ns_, ndup = shape_map(ni)
        shape_dup = odup | ndup
        s_add = sorted([n2 for n2 in ns_ if n2 not in os_])
        s_del = sorted([n2 for n2 in os_ if n2 not in ns_])
        s_chg = []
        for n2 in sorted(set(os_) & set(ns_)):
            a, b = os_[n2], ns_[n2]
            fields = []
            if a.get('text') != b.get('text'):
                fields.append(f"文字 '{clip(a.get('text'))}'→'{clip(b.get('text'))}'")
            if a.get('onaction') != b.get('onaction'):
                fields.append(f"OnAction {a.get('onaction')}→{b.get('onaction')}")
            if (a.get('l'), a.get('t')) != (b.get('l'), b.get('t')):
                fields.append(f"位置 ({a.get('l')},{a.get('t')})→({b.get('l')},{b.get('t')})")
            if fields:
                s_chg.append((n2, fields))

        def table_map(info):
            return {t['name']: t.get('address') for t in info.get('tables', ())}
        ot, nt = table_map(oi), table_map(ni)
        t_add = sorted([n2 for n2 in nt if n2 not in ot])
        t_del = sorted([n2 for n2 in ot if n2 not in nt])
        t_chg = sorted([n2 for n2 in ot if n2 in nt and ot[n2] != nt[n2]])

        has_diff = any((added, removed, changed, m_add, m_del,
                        s_add, s_del, s_chg, t_add, t_del, t_chg,
                        oi.get('dims') != ni.get('dims')))
        if not has_diff:
            # シート自体を表示しない場合も、比較できなかった事実は落とさない
            if merged_unscanned:
                merged_note_hidden.append(name)
            if cell_limit is not None:
                trunc_note_hidden.append(f"{name}（{cell_limit}行まで）")
            continue
        diff_sheets += 1
        print(f"\n* シート: {name}")
        if oi.get('dims') != ni.get('dims'):
            print(f"  使用範囲: {oi.get('dims')} → {ni.get('dims')}")
        if cell_limit is not None:
            total = None
            for _i in (oi, ni):
                t = _i.get('cells_truncated') or {}
                if t.get('total_rows'):
                    total = max(total or 0, int(t['total_rows']))
            tail = f"（全{total}行）" if total else ""
            print(f"  ※ 打ち切られた snapshot のため、セルは {cell_limit}行までしか比較していません"
                  f"{tail}＝それ以降の行の変更は検出できません")
        if merged_unscanned:
            print("  ※ 結合セル未走査の snapshot（範囲が大）＝結合の差分は比較していない")
        if shape_dup:
            print(f"  ※ 同名の図形が複数: {', '.join(sorted(shape_dup))}"
                  "（名前単位の比較のため2つ目以降は対象外）")
        if changed:
            show("セル変更", changed,
                 lambda k: f"{addr(k)}: '{clip(oc[k])}' → '{clip(nc_[k])}'")
        if added:
            show("セル追加", added, lambda k: f"{addr(k)}: '{clip(nc_[k])}'")
        if removed:
            show("セル削除", removed, lambda k: f"{addr(k)}: '{clip(oc[k])}'")
        if m_add:
            show("結合追加", m_add, lambda a: a)
        if m_del:
            show("結合解除", m_del, lambda a: a)
        if s_add:
            show("図形追加", s_add,
                 lambda n2: n2 + (f"「{clip(ns_[n2].get('text'))}」" if ns_[n2].get('text') else ""))
        if s_del:
            show("図形削除", s_del,
                 lambda n2: n2 + (f"「{clip(os_[n2].get('text'))}」" if os_[n2].get('text') else ""))
        if s_chg:
            show("図形変更", s_chg, lambda it: f"{it[0]}: " + " / ".join(it[1]))
        if t_add:
            show("テーブル追加", t_add, lambda n2: f"{n2} ({nt[n2]})")
        if t_del:
            show("テーブル削除", t_del, lambda n2: f"{n2} ({ot[n2]})")
        if t_chg:
            show("テーブル範囲変更", t_chg, lambda n2: f"{n2}: {ot[n2]} → {nt[n2]}")

    if merged_note_hidden:
        print(f"\n※ 結合セル未走査の snapshot のため、次のシートは結合の差分を比較できていません: "
              f"{', '.join(merged_note_hidden)}")
    if trunc_note_hidden:
        print(f"\n※ 打ち切られた snapshot のため、次のシートはセルを途中までしか比較できていません: "
              f"{', '.join(trunc_note_hidden)}")
    print("\n" + "-" * 60)
    if diff_sheets:
        print(f"差分あり: シート{diff_sheets}枚に変更")
    else:
        print("差分なし（2つのスナップショットは一致）")
    return True


def cmd_wiring(args):
    """ボタン⇔マクロの配線図: wiring [excel_file] [--json]

    シート上の図形/ボタンに登録された OnAction を全部拾い（グループは1段展開）、
    ブック内のマクロ名簿と突き合わせる。行き先のないボタン＝壊れた配線の検出器。
    call-graph（孤立マクロ）・docs（マクロ→ボタン逆引き）と対になる「ボタン側から見た地図」。
    VBA に触れないブック（保護/VBOM未信頼）でも配線一覧だけは出す（実在確認のみ縮退）。
    別ブック修飾（'秀.xlam'!マクロ 等）の配線は名簿の外＝実在確認せず「外部ブック先」として
    別枠で表示し、壊れた配線には数えない（このブックの名簿で×を付けると誤検出になる）。
    終了コード: テキスト表示は壊れた配線ありで 1（lint 型）。--json は常に 0（broken が判定を運ぶ）。
    """
    import json
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, readonly=True)   # 診断は読むだけ

    # マクロ名簿（Sub/Function 名 → 正式名）
    known = {}
    vba_error = None
    proc_pat = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE)
    try:
        for comp in wb.VBProject.VBComponents:
            if int(comp.Type) not in (1, 2, 3, 100):
                continue
            cm = comp.CodeModule
            n = cm.CountOfLines
            code = cm.Lines(1, n) if n else ""
            for m in proc_pat.finditer(code):
                known.setdefault(m.group(1).lower(), m.group(1))
    except Exception as e:
        vba_error = e

    rows = []
    unread_sheets = []          # 図形を列挙できず、配線図から丸ごと落ちたシート
    for sh in wb.Worksheets:
        shapes = []
        try:
            _collect_shapes(sh.Shapes, shapes)
        except Exception as e:
            # 黙って continue すると、そのシートのボタンが1本も出ないまま
            # 「行き先なし 0本」＝健全という誤った安心につながる
            unread_sheets.append((sh.Name, str(e)))
            continue
        for s in shapes:
            macro = s.get('onaction')
            if not macro:
                continue
            # ブック修飾つき配線。自ブック名なら普通に照合、他ブック(アドイン等)なら
            # このブックの名簿では実在確認できない＝外部ブック先として別枠にする
            book = s.get('onaction_book')
            if book and book.lower() == wb.Name.lower():
                book = None
            resolved = None
            if book is None and vba_error is None:
                hit = known.get(macro.lower())
                if hit is None and '.' in macro:
                    # 「モジュール名.マクロ名」形式（同名マクロがあると Excel が
                    # 自動でこの形式にする）は末尾名で解決する
                    hit = known.get(macro.rsplit('.', 1)[-1].lower())
                resolved = hit is not None
            rows.append({'sheet': sh.Name, 'shape': s.get('name', '(図形)'),
                         'text': s.get('text'), 'group': s.get('group'),
                         'macro': macro, 'book': book, 'resolved': resolved})

    broken = [r for r in rows if r['resolved'] is False]
    external = [r for r in rows if r['book']]

    if getattr(args, 'json', False):
        print(json.dumps({"success": True, "book": wb.Name,
                          "vba_readable": vba_error is None,
                          "wires": rows, "broken": len(broken),
                          "external": len(external),
                          "unread_sheets": [{"sheet": n, "error": e}
                                            for n, e in unread_sheets]},
                         ensure_ascii=False))
        # JSON では broken フィールドが判定を運ぶ。ここで exit 1 にすると
        # success:true の JSON に MCP が「失敗しました」を付けて食い違う
        return True

    def label(r):
        t = f"「{str(r['text'])[:30]}」" if r['text'] else ""
        g = f"（グループ {r['group']} 内）" if r['group'] else ""
        return f"{r['shape']}{t}{g}"

    print(f"===== ボタン⇔マクロ配線図: {wb.Name} =====")
    if vba_error is not None:
        print("※ VBA プロジェクトに触れないため実在確認は未実施（配線の一覧のみ）")
        print(f"   詳細: {vba_error}")
    if unread_sheets:
        print("⚠ 次のシートは図形を読み取れず、配線図に含まれていません（未検査）:")
        for n, e in unread_sheets:
            print(f"   {n}: {e}")
    if not rows:
        print("OnAction が登録された図形/ボタンはありません"
              + ("（ただし上記シートは未検査）" if unread_sheets else ""))
        return True

    cur = None
    for r in rows:
        if r['sheet'] != cur:
            cur = r['sheet']
            print(f"\nシート: {cur}")
        if r['book']:
            mark = "（外部ブック先・このブックの名簿では確認できない）"
            dest = f"{r['book']}!{r['macro']}"
        elif r['resolved'] is True:
            mark = "○"
            dest = r['macro']
        elif r['resolved'] is False:
            mark = "×（存在しない）"
            dest = r['macro']
        else:
            mark = "（未確認）"
            dest = r['macro']
        print(f"  {label(r)} → {dest}  {mark}")

    print("\n" + "-" * 60)
    if vba_error is None:
        line = (f"配線 {len(rows)}本 / 実在 {len(rows) - len(broken) - len(external)}"
                f" / 行き先なし {len(broken)}")
        if external:
            line += f" / 外部ブック先 {len(external)}（確認対象外）"
        print(line)
        if broken:
            print("\n⚠ 行き先のないボタン（OnAction 先のマクロが見つからない）:")
            for r in broken:
                print(f"  {r['sheet']} / {label(r)} → {r['macro']}")
    else:
        print(f"配線 {len(rows)}本（実在確認なし）")
    if unread_sheets:
        # 未検査シートがある以上「行き先なし 0本」は健全の証明にならない
        print(f"※ 図形を読み取れなかったシートが {len(unread_sheets)}枚あります"
              f"（{', '.join(n for n, _ in unread_sheets)}）＝この配線図は全数ではありません")
    print("参考: マクロ側から見た地図は call-graph（孤立検出）/ docs（マクロ→ボタン逆引き）")
    return not broken


@protect_safe
def cmd_screenshot(args):
    """範囲を画像(PNG)として書き出す（目・画像版）"""
    target_file, rest = parse_target_and_rest(args.posargs)
    spec = rest[0] if rest else None
    out_path = os.path.abspath(getattr(args, 'out_opt', None) or LAST_VIEW_FILE)

    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, spec)

    # 対象シートをアクティブにすると CopyPicture が安定する
    try:
        ws.Activate()
    except Exception:
        pass

    # Appearance: xlScreen=1 / xlPrinter=2,  Format: xlBitmap=2 / xlPicture(EMF)=-4147
    # ※ chart へ貼る前に cob.Activate() しないと Paste が無反応で白紙PNGになる（要・最重要）
    # ※ 成否は出力サイズではなく「貼り付け後の Shapes 数」で判定する（白紙でもファイルは生成されるため）
    attempts = [(1, 2), (1, -4147), (2, 2)]
    last_err = None
    for appearance, fmt in attempts:
        cob = None
        try:
            rng.CopyPicture(appearance, fmt)
            time.sleep(0.4)
            pythoncom.PumpWaitingMessages()

            cob = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
            cob.Activate()                      # ← これが無いと貼り付かない
            chart = cob.Chart
            time.sleep(0.3)
            chart.Paste()
            time.sleep(0.4)
            pythoncom.PumpWaitingMessages()

            pasted = chart.Shapes.Count          # 1 以上なら貼り付け成功
            if pasted >= 1:
                if os.path.exists(out_path):
                    os.remove(out_path)
                chart.Export(out_path, "PNG")
                cob.Delete()
                cob = None
                if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                    _screenshot_cleanup(xl, ws)
                    print(f"シート: {ws.Name}   範囲: {rng.Address}")
                    print(f"画像保存: {out_path}")
                    print("（注: 一時グラフの作成を伴うため、ブックの Undo 履歴は消えています）")
                    return True
                last_err = "Export に失敗しました"
            else:
                last_err = "クリップボードから貼り付けできませんでした"
        except Exception as e:
            last_err = str(e)
        finally:
            if cob is not None:
                try:
                    cob.Delete()
                except Exception:
                    pass
        time.sleep(0.4)

    _screenshot_cleanup(xl, ws)
    print(f"エラー: スクリーンショットに失敗しました ({last_err})")
    return False


def _screenshot_cleanup(xl, ws):
    """screenshot の後始末（成功・失敗の両経路で共通）。

    ChartObject 操作で選択が動くため A1 に戻し、CopyPicture で
    クリップボードに残った画像もクリアする（copy-range と対称）。
    """
    try:
        ws.Range("A1").Select()
    except Exception:
        pass
    try:
        xl.CutCopyMode = False
    except Exception:
        pass




__all__ = [
    'LAST_VIEW_FILE',
    '_cell_compare_limit',
    '_sheet_first_row',
    '_collect_shapes',
    '_disp_pad',
    '_disp_truncate',
    '_disp_width',
    '_merged_areas_in_range',
    '_resolve_range',
    '_screenshot_cleanup',
    '_shape_text',
    '_values_to_grid',
    '_whole_sheet_spec',
    'cmd_read_range',
    'cmd_read_selection',
    'cmd_screenshot',
    'cmd_sheet_info',
    'cmd_snapshot',
    'cmd_snapshot_diff',
    'cmd_wiring',
]
