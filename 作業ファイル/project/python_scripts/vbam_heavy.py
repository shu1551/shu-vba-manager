# -*- coding: utf-8 -*-
"""vbam_heavy.py — vba_manager 分割パート: 重量級コマンド（チャート/ピボット/スライサー/計算モード/PowerQuery/データモデル）

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
from vbam_edit import *  # noqa: F401,F403
# ================================================================
# 重量級コマンド (1) チャート
# ================================================================

_XL_CHART_TYPE = {
    'column':  51,     # xlColumnClustered
    'bar':     57,     # xlBarClustered
    'line':    4,      # xlLine
    'pie':     5,      # xlPie
    'scatter': -4169,  # xlXYScatter
    'area':    1,      # xlArea
    'doughnut': -4120, # xlDoughnut
}
_XL_CHART_TYPE_NAME = {v: k for k, v in _XL_CHART_TYPE.items()}


@protect_safe
def cmd_chart(args):
    """グラフ操作: chart <create|list|delete> ...

      chart create <data_range> [--type column|bar|line|pie|scatter|area]
                   [--title "見出し"] [--at セル] [--name 名] [--width N --height N]
      chart list
      chart delete <name>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: chart <create|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sh in wb.Worksheets:   # グラフシートは ChartObjects を持たないため除外
            for co in sh.ChartObjects():
                cnt += 1
                try:
                    t = _XL_CHART_TYPE_NAME.get(int(co.Chart.ChartType), co.Chart.ChartType)
                except Exception:
                    t = '?'
                print(f"[{sh.Name}] {co.Name}  type={t}")
        if cnt == 0:
            print("グラフはありません。")
        return True

    if action == 'create':
        if len(rest) < 2:
            print("使い方: chart create <data_range> [--type ...] [--title ...] [--at セル] [--name 名]")
            return False
        ws, rng = _resolve_range(xl, wb, rest[1])

        # 配置: --at 指定があればそのセルの左上、なければデータ範囲の右隣
        at = getattr(args, 'at', None)
        if at:
            anchor = ws.Range(at)
            left, top = anchor.Left, anchor.Top
        else:
            left, top = rng.Left + rng.Width + 10, rng.Top
        width = float(getattr(args, 'width', None) or 360)
        height = float(getattr(args, 'height', None) or 216)

        co = ws.ChartObjects().Add(left, top, width, height)
        ch = co.Chart
        ch.SetSourceData(rng)
        ctype = (getattr(args, 'type', None) or 'column').lower()
        if ctype not in _XL_CHART_TYPE:
            print(f"未知のグラフ種別: {ctype}（{'/'.join(_XL_CHART_TYPE)}）")
            co.Delete()
            return False
        ch.ChartType = _XL_CHART_TYPE[ctype]
        if getattr(args, 'title', None):
            ch.HasTitle = True
            ch.ChartTitle.Text = args.title
        if getattr(args, 'name', None):
            # 名前の重複・禁止文字で失敗すると、自動名（グラフ 1 等）のグラフだけが
            # 残って「名前を付けたはずのグラフが chart list に無い」状態になる。
            # 種別不正の経路と同じく、このコマンドが作ったものは片づけてから失敗を返す。
            try:
                co.Name = args.name
            except Exception as ex:
                print(f"エラー: グラフ名 '{args.name}' を設定できません: {ex}")
                print("  同名のグラフが既にある可能性があります（chart list で確認）。")
                try:
                    co.Delete()
                    print("  作成途中のグラフは片づけました。")
                except Exception:
                    pass
                return False
        print(f"グラフ作成: [{ws.Name}] {co.Name}  種別={ctype}  データ={rng.Address}")
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: chart delete <name>"); return False
        name = rest[1]
        for sh in wb.Worksheets:   # グラフシートは ChartObjects を持たないため除外
            for co in sh.ChartObjects():
                if co.Name == name:
                    co.Delete()
                    print(f"グラフ削除: [{sh.Name}] {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: グラフ '{name}' が見つかりません")
        return False

    print(f"未知のアクション: {action}（create|list|delete）")
    return False


_XL_AXIS = {'category': 1, 'value': 2, 'series': 3, 'secondary': 2}   # xlCategory/xlValue/xlSeriesAxis
_XL_LEGEND_POS = {'bottom': -4107, 'corner': 2, 'top': -4160, 'right': -4152, 'left': -4131}
_XL_TRENDLINE = {'linear': -4132, 'exponential': 5, 'logarithmic': -4133,
                 'movingaverage': 6, 'polynomial': 3, 'power': 4}


@protect_safe
def cmd_chart_config(args):
    """グラフ詳細設定: chart-config <action> <chart名> ...

      set-source <chart> <range>                          データ範囲を再設定
      set-type <chart> <type>                             種別変更(column/bar/line/pie/...)
      set-title <chart> <text>                            グラフタイトル
      set-axis-title <chart> <category|value|secondary> <text>   軸タイトル
      axis-format <chart> <axis> [format]                 軸の表示形式 get/set
      axis-scale <chart> <axis> [--min N --max N --major N --minor N]  軸目盛
      gridlines <chart> <axis> [--major on|off --minor on|off]        目盛線
      legend <chart> <bottom|top|right|left|corner|off>   凡例
      style <chart> <1-48>                                組込スタイル
      placement <chart> <1|2|3>                           1=移動+サイズ/2=移動のみ/3=自由
      data-labels <chart> [--value --percent --category --series --position 位置]
      add-series <chart> <values_range> [--series-name 名] [--category-range 範囲]
      remove-series <chart> <index>
      series-format <chart> <index> [--marker-style N --marker-size N --marker-fg #.. --marker-bg #.. --invert]
      trendline list <chart> <series_index>
      trendline add  <chart> <series_index> <type> [--name 名]
      trendline delete <chart> <series_index> <trendline_index>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: chart-config <action> <chart名> ...（詳細は --help）")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    def find_chart(name):
        for sh in wb.Worksheets:   # グラフシートは ChartObjects を持たないため除外
            for co in sh.ChartObjects():
                if co.Name == name:
                    return sh, co, co.Chart
        return None, None, None

    # trendline はサブ動詞が rest[1] に来る特例
    if action == 'trendline':
        sub = rest[1].lower() if len(rest) >= 2 else ''
        cname = rest[2] if len(rest) >= 3 else None
        sh, co, ch = find_chart(cname) if cname else (None, None, None)
        if not ch:
            print(f"エラー: グラフ '{cname}' が見つかりません（chart list で確認）。"); return False
        try:
            sidx = int(rest[3]) if len(rest) >= 4 else 1
        except ValueError:
            print("series_index は数値で指定してください。"); return False
        s = ch.SeriesCollection(sidx)
        if sub == 'list':
            tls = s.Trendlines()
            print(f"--- {cname} 系列{sidx} の近似曲線 ({tls.Count}) ---")
            for i in range(1, tls.Count + 1):
                tl = tls.Item(i)
                try:
                    nm = tl.Name
                except Exception:
                    nm = f"#{i}"
                print(f"  [{i}] {nm}")
            if tls.Count == 0:
                print("  (なし)")
            return True
        if sub == 'add':
            ttype = (rest[4] if len(rest) >= 5 else 'linear').lower()
            if ttype not in _XL_TRENDLINE:
                print(f"未知の種別: {ttype}（{'/'.join(_XL_TRENDLINE)}）"); return False
            tl = s.Trendlines().Add(_XL_TRENDLINE[ttype])
            if getattr(args, 'name', None):
                tl.Name = args.name
            print(f"近似曲線追加: {cname} 系列{sidx} {ttype}")
            print("（保存はしていません）"); return True
        if sub == 'delete':
            # 削除系のインデックス省略は「黙って#1が消える」事故のもと。明示必須。
            if len(rest) < 5:
                print("使い方: chart-config trendline delete <chart> <series_index> <trendline_index>")
                print("  （削除対象の trendline_index は省略できません。trendline list で確認）")
                return False
            try:
                tidx = int(rest[4])
            except ValueError:
                print("trendline_index は数値で指定してください。"); return False
            s.Trendlines().Item(tidx).Delete()
            print(f"近似曲線削除: {cname} 系列{sidx} #{tidx}")
            print("（保存はしていません）"); return True
        print("使い方: chart-config trendline <list|add|delete> <chart> <series_index> ...")
        return False

    # それ以外は rest[1] が chart 名
    cname = rest[1] if len(rest) >= 2 else None
    sh, co, ch = find_chart(cname) if cname else (None, None, None)
    if not ch:
        print(f"エラー: グラフ '{cname}' が見つかりません（chart list で確認）。"); return False

    def get_axis(axname):
        a = (axname or 'value').lower()
        if a not in _XL_AXIS:
            return None
        if a == 'secondary':
            return ch.Axes(2, 2)            # xlValue, xlSecondary
        return ch.Axes(_XL_AXIS[a])

    if action == 'set-source':
        if len(rest) < 3:
            print("使い方: chart-config set-source <chart> <range>"); return False
        ws, rng = _resolve_range(xl, wb, rest[2])
        ch.SetSourceData(rng)
        print(f"データ範囲再設定: {cname} ← {rng.Address}")
        print("（保存はしていません）"); return True

    if action == 'set-type':
        t = (rest[2] if len(rest) >= 3 else '').lower()
        if t not in _XL_CHART_TYPE:
            print(f"未知の種別: {t}（{'/'.join(_XL_CHART_TYPE)}）"); return False
        ch.ChartType = _XL_CHART_TYPE[t]
        print(f"種別変更: {cname} → {t}"); print("（保存はしていません）"); return True

    if action == 'set-title':
        txt = rest[2] if len(rest) >= 3 else ''
        ch.HasTitle = True
        ch.ChartTitle.Text = txt
        print(f"タイトル設定: {cname} = {txt}"); print("（保存はしていません）"); return True

    if action == 'set-axis-title':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        txt = rest[3] if len(rest) >= 4 else ''
        ax.HasTitle = True
        ax.AxisTitle.Text = txt
        print(f"軸タイトル設定: {cname} {rest[2]} = {txt}"); print("（保存はしていません）"); return True

    if action == 'axis-format':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        if len(rest) >= 4:
            ax.TickLabels.NumberFormat = rest[3]
            print(f"軸表示形式設定: {cname} {rest[2]} = {rest[3]}")
            print("（保存はしていません）"); return True
        else:
            print(f"軸表示形式: {cname} {rest[2]} = {ax.TickLabels.NumberFormat}"); return True

    if action == 'axis-scale':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        changed = []
        for opt, prop in (('min', 'MinimumScale'), ('max', 'MaximumScale'),
                          ('major', 'MajorUnit'), ('minor', 'MinorUnit')):
            v = getattr(args, opt, None)
            if v is not None:
                setattr(ax, prop, float(v)); changed.append(f"{opt}={v}")
        if changed:
            print(f"軸目盛設定: {cname} {rest[2]} [{', '.join(changed)}]")
            print("（保存はしていません）"); return True
        else:
            print(f"軸目盛: {cname} {rest[2]} min={ax.MinimumScale} max={ax.MaximumScale} "
                  f"major={ax.MajorUnit} minor={ax.MinorUnit}"); return True

    if action == 'gridlines':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        mj = getattr(args, 'major', None); mn = getattr(args, 'minor', None)
        if mj is None and mn is None:
            print(f"目盛線: {cname} {rest[2]} major={ax.HasMajorGridlines} minor={ax.HasMinorGridlines}")
            return True
        if mj is not None:
            ax.HasMajorGridlines = (mj.lower() == 'on')
        if mn is not None:
            ax.HasMinorGridlines = (mn.lower() == 'on')
        print(f"目盛線設定: {cname} {rest[2]} major={getattr(args,'major',None)} minor={getattr(args,'minor',None)}")
        print("（保存はしていません）"); return True

    if action == 'legend':
        pos = (rest[2] if len(rest) >= 3 else 'bottom').lower()
        if pos == 'off':
            ch.HasLegend = False
            print(f"凡例: {cname} = 非表示")
        else:
            if pos not in _XL_LEGEND_POS:
                print(f"位置は {'/'.join(_XL_LEGEND_POS)}|off で指定してください。"); return False
            ch.HasLegend = True
            ch.Legend.Position = _XL_LEGEND_POS[pos]
            print(f"凡例: {cname} = {pos}")
        print("（保存はしていません）"); return True

    if action == 'style':
        try:
            sid = int(rest[2]) if len(rest) >= 3 else 1
        except ValueError:
            print("使い方: chart-config style <chart> <1-48の数値>"); return False
        ch.ChartStyle = sid
        print(f"スタイル設定: {cname} = {sid}"); print("（保存はしていません）"); return True

    if action == 'placement':
        try:
            pl = int(rest[2]) if len(rest) >= 3 else 1
        except ValueError:
            print("使い方: chart-config placement <chart> <1|2|3>"); return False
        co.Placement = pl
        names = {1: '移動+サイズ', 2: '移動のみ', 3: '自由配置'}
        print(f"配置方法: {cname} = {pl}（{names.get(pl, pl)}）"); print("（保存はしていません）"); return True

    if action == 'data-labels':
        ch.ApplyDataLabels(
            ShowValue=bool(getattr(args, 'value', False)),
            ShowPercentage=bool(getattr(args, 'percent', False)),
            ShowCategoryName=bool(getattr(args, 'category', False)),
            ShowSeriesName=bool(getattr(args, 'series', False)))
        pos = getattr(args, 'position', None)
        if pos:
            posmap = {'center': -4108, 'insideend': 3, 'outsideend': 2, 'bestfit': 5, 'insidebase': 4}
            if pos.lower() not in posmap:
                print(f"⚠ 未知の位置: {pos}（{'/'.join(posmap)}）位置指定はスキップしました。")
            else:
                # 全系列に適用（以前は系列1のみで、複数系列だと部分適用のまま成功表示だった）
                pos_failed = []
                for si in range(1, ch.SeriesCollection().Count + 1):
                    try:
                        ch.SeriesCollection(si).DataLabels().Position = posmap[pos.lower()]
                    except Exception:
                        pos_failed.append(si)
                if pos_failed:
                    print(f"⚠ 位置指定が適用できなかった系列: {pos_failed}"
                          "（グラフ種別によって位置指定不可の場合があります）")
        print(f"データラベル設定: {cname}"); print("（保存はしていません）"); return True

    if action == 'add-series':
        ws, vrng = _resolve_range(xl, wb, rest[2])
        s = ch.SeriesCollection().NewSeries()
        s.Values = vrng
        if getattr(args, 'series_name', None):
            s.Name = args.series_name
        if getattr(args, 'category_range', None):
            _, crng = _resolve_range(xl, wb, args.category_range)
            s.XValues = crng
        print(f"系列追加: {cname} ← {vrng.Address}"); print("（保存はしていません）"); return True

    if action == 'remove-series':
        # 削除系のインデックス省略は「黙って系列1が消える」事故のもと。明示必須。
        if len(rest) < 3:
            print("使い方: chart-config remove-series <chart> <series_index>")
            print("  （削除対象の series_index は省略できません）")
            return False
        try:
            idx = int(rest[2])
        except ValueError:
            print("series_index は数値で指定してください。"); return False
        ch.SeriesCollection(idx).Delete()
        print(f"系列削除: {cname} #{idx}"); print("（保存はしていません）"); return True

    if action == 'series-format':
        # 省略時は従来どおり系列1に適用する。出力に「#番号」を明示するので
        # 無言の書き換えにはならない（削除系の明示必須とは事情が違う）
        try:
            idx = int(rest[2]) if len(rest) >= 3 else 1
        except ValueError:
            print("series_index は数値で指定してください。"); return False
        s = ch.SeriesCollection(idx)
        ch_list = []
        if getattr(args, 'marker_style', None) is not None:
            s.MarkerStyle = int(args.marker_style); ch_list.append('style')
        if getattr(args, 'marker_size', None) is not None:
            s.MarkerSize = int(args.marker_size); ch_list.append('size')
        if getattr(args, 'marker_fg', None):
            s.MarkerForegroundColor = _hex_to_excel_color(args.marker_fg); ch_list.append('fg')
        if getattr(args, 'marker_bg', None):
            s.MarkerBackgroundColor = _hex_to_excel_color(args.marker_bg); ch_list.append('bg')
        if getattr(args, 'invert', False):
            s.InvertIfNegative = True; ch_list.append('invert')
        if not ch_list:
            # 何も指定が無いのに「系列書式: …[]」と成功表示するのは実態と食い違う
            print("エラー: 変更する書式オプションが指定されていません。")
            print("  --marker-style/--marker-size/--marker-fg/--marker-bg/--invert のいずれかを指定してください。")
            return False
        print(f"系列書式: {cname} #{idx} [{', '.join(ch_list)}]"); print("（保存はしていません）"); return True

    print(f"未知のアクション: {action}")
    return False


# ================================================================
# 重量級コマンド (2) ピボットテーブル
# ================================================================

_XL_PIVOT_FUNC = {'sum': -4157, 'count': -4112, 'average': -4106,
                  'max': -4136, 'min': -4139}


def _unique_sheet_name(wb, base):
    """重複しないシート名を返す"""
    existing = {sh.Name for sh in wb.Sheets}
    if base not in existing:
        return base
    i = 2
    while f"{base}{i}" in existing:
        i += 1
    return f"{base}{i}"


@protect_safe
def cmd_pivot(args):
    """ピボット操作: pivot <create|list|delete> ...

      pivot create <data_range> [--rows F1,F2] [--cols F1] [--values F1,F2]
                   [--func sum|count|average|max|min] [--sheet 出力シート | --at セル] [--name 名]
      pivot list
      pivot delete <name>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: pivot <create|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
            for pt in sh.PivotTables():
                cnt += 1
                print(f"[{sh.Name}] {pt.Name}")
        if cnt == 0:
            print("ピボットテーブルはありません。")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: pivot delete <name>"); return False
        name = rest[1]
        for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
            for pt in sh.PivotTables():
                if pt.Name == name:
                    pt.TableRange2.Clear()
                    print(f"ピボット削除: [{sh.Name}] {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: ピボット '{name}' が見つかりません")
        return False

    if action == 'create':
        if len(rest) < 2:
            print("使い方: pivot create <data_range> [--rows ..][--cols ..][--values ..][--func ..]")
            return False
        ws_s, rng = _resolve_range(xl, wb, rest[1])

        # 出力先の決定: --sheet > --at > 新規シート
        sheet_opt = getattr(args, 'sheet', None)
        at = getattr(args, 'at', None)
        created_sheet = None      # このコマンドが新規に作ったシート（失敗時の後始末用）
        if sheet_opt:
            dws = None
            for sh in wb.Sheets:
                # Excel のシート名重複判定は大文字小文字を区別しない
                if sh.Name.lower() == sheet_opt.lower():
                    dws = sh; break
            if dws is None:
                dws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
                created_sheet = dws
                try:
                    dws.Name = sheet_opt
                except Exception as ex:
                    # 禁止文字(/ 等)・31文字超で失敗すると無名シートが残骸になる
                    print(f"エラー: シート名 '{sheet_opt}' を設定できません: {ex}")
                    xl.DisplayAlerts = False
                    try:
                        dws.Delete()
                    finally:
                        xl.DisplayAlerts = True
                    return False
            dest = dws.Range("A3")
        elif at:
            dest = ws_s.Range(at)
        else:
            dws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
            dws.Name = _unique_sheet_name(wb, "ピボット")
            created_sheet = dws
            dest = dws.Range("A3")

        pt = None
        pc = None
        funcname = (getattr(args, 'func', None) or 'sum').lower()
        values = getattr(args, 'values', None)
        try:
            pc = wb.PivotCaches().Create(1, rng)        # xlDatabase=1
            name = getattr(args, 'name', None)
            if name:
                pt = pc.CreatePivotTable(dest, name)
            else:
                pt = pc.CreatePivotTable(dest)

            def set_fields(spec, orient):
                if not spec:
                    return
                for f in spec.split(','):
                    f = f.strip()
                    if f:
                        pt.PivotFields(f).Orientation = orient

            set_fields(getattr(args, 'rows', None), 1)   # xlRowField
            set_fields(getattr(args, 'cols', None), 2)   # xlColumnField

            func = _XL_PIVOT_FUNC.get(funcname, -4157)
            if values:
                for f in values.split(','):
                    f = f.strip()
                    if f:
                        df = pt.AddDataField(pt.PivotFields(f), f"{funcname}/{f}", func)
        except Exception as ex:
            # 存在しないフィールド名等で途中失敗すると、追加した新規シートと
            # 空ピボットが残骸として残り、再実行のたび「ピボット2/3…」と増殖する。
            # このコマンドが作ったものだけ片づける（既存シートのセルは絶対に消さない）
            print(f"エラー: ピボット作成に失敗しました: {ex}")
            print("  --rows/--cols/--values のフィールド名がデータ範囲の見出しと一致しているか確認してください。")
            try:
                if created_sheet is not None:
                    xl.DisplayAlerts = False
                    try:
                        created_sheet.Delete()
                    finally:
                        xl.DisplayAlerts = True
                    print("  作成途中のシート（このコマンドが作ったもの）は片づけました。")
                elif pt is not None:
                    # TableRange2 はこのピボット自身の出力範囲。作成途中の空ピボットを
                    # 残すと再実行のたび「ピボット2/3…」と残骸が増殖する（pivot delete と
                    # 同じ後始末を自動でやるだけ＝既存セルは巻き込まない）
                    pt.TableRange2.Clear()
                    print("  作成途中の空ピボットは片づけました。")
            except Exception:
                pass
            # PivotCache の後始末。参照が外れたキャッシュは保存時に破棄されるが、
            # Delete を持つバージョンでは明示的に落としておく（孤児キャッシュ対策）。
            try:
                if pc is not None:
                    pc.Delete()
            except Exception:
                pass
            pc = None
            return False

        print(f"ピボット作成: [{pt.Parent.Name}] {pt.Name}  ソース={ws_s.Name}!{rng.Address}")
        print(f"  行={getattr(args,'rows',None) or '-'}  列={getattr(args,'cols',None) or '-'}  "
              f"値={values or '-'}({funcname})")
        print("（保存はしていません）")
        return True

    print(f"未知のアクション: {action}（create|list|delete）")
    return False


def _find_pivot(wb, name):
    for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
        for pt in sh.PivotTables():
            if pt.Name == name:
                return sh, pt
    return None, None


_XL_PIVOT_ORIENT = {'row': 1, 'col': 2, 'column': 2, 'filter': 3, 'page': 3, 'value': 4, 'data': 4, 'hidden': 0}


@protect_safe
def cmd_pivot_field(args):
    """ピボットのフィールド管理: pivot-field <action> <pivot名> <フィールド> ...

      list <pivot>
      add-row|add-col|add-filter <pivot> <field>
      add-value <pivot> <field> [--func sum|count|average|max|min] [--name 表示名]
      remove <pivot> <field>
      set-func <pivot> <field> <func>            データフィールドの集計関数
      set-name <pivot> <field> <表示名>          データフィールドの表示名
      set-format <pivot> <field> <書式コード>    数値書式
      set-filter <pivot> <field> 値1 値2 ...     表示する値を限定
      sort <pivot> <field> <asc|desc>
      group-date <pivot> <field> <days|months|quarters|years>
      group-numeric <pivot> <field> <start> <end> <interval>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: pivot-field <action> <pivot名> <field> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)
    pname = rest[1] if len(rest) >= 2 else None
    sh, pt = _find_pivot(wb, pname) if pname else (None, None)
    if not pt:
        print(f"エラー: ピボット '{pname}' が見つかりません（pivot list で確認）。"); return False

    def pf(field):
        try:
            return pt.PivotFields(field)
        except Exception:
            return None

    if action == 'list':
        print(f"--- {pname} のフィールド ---")
        orient_name = {1: '行', 2: '列', 3: 'フィルタ', 4: '値', 0: '未配置'}
        for i in range(1, pt.PivotFields().Count + 1):
            f = pt.PivotFields().Item(i)
            try:
                o = int(f.Orientation)
            except Exception:
                o = 0
            print(f"  {f.Name}: {orient_name.get(o, o)}")
        return True

    field = rest[2] if len(rest) >= 3 else None
    if action in ('add-row', 'add-col', 'add-filter'):
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Orientation = {'add-row': 1, 'add-col': 2, 'add-filter': 3}[action]
        print(f"フィールド配置: {pname}[{field}] = {action[4:]}"); print("（保存はしていません）"); return True

    if action == 'add-value':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        fn = (getattr(args, 'func', None) or 'sum').lower()
        func = _XL_PIVOT_FUNC.get(fn, -4157)
        cname = getattr(args, 'name', None) or f"{fn}/{field}"
        pt.AddDataField(p, cname, func)
        print(f"値フィールド追加: {pname}[{field}] ({fn}) 表示名={cname}"); print("（保存はしていません）"); return True

    if action == 'remove':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Orientation = 0     # xlHidden
        print(f"フィールド除外: {pname}[{field}]"); print("（保存はしていません）"); return True

    if action == 'set-func':
        fn = (rest[3] if len(rest) >= 4 else 'sum').lower()
        if fn not in _XL_PIVOT_FUNC:
            print(f"未知の関数: {fn}（{'/'.join(_XL_PIVOT_FUNC)}）"); return False
        # データフィールドは表示名で参照されるため DataFields を走査
        target = None
        for i in range(1, pt.DataFields.Count + 1):
            d = pt.DataFields.Item(i)
            if d.Name == field or d.SourceName == field:
                target = d; break
        if target is None:
            print(f"エラー: 値フィールド '{field}' が見つかりません。"); return False
        target.Function = _XL_PIVOT_FUNC[fn]
        print(f"集計関数変更: {pname}[{field}] = {fn}"); print("（保存はしていません）"); return True

    if action == 'set-name':
        newname = rest[3] if len(rest) >= 4 else None
        if not newname:
            print("使い方: pivot-field set-name <pivot> <field> <新しい表示名>")
            return False
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Caption = newname
        print(f"表示名変更: {pname}[{field}] → {newname}"); print("（保存はしていません）"); return True

    if action == 'set-format':
        code = rest[3] if len(rest) >= 4 else None
        if not code:
            print('使い方: pivot-field set-format <pivot> <field> <書式コード（例: "#,##0"）>')
            return False
        for i in range(1, pt.DataFields.Count + 1):
            d = pt.DataFields.Item(i)
            if d.Name == field or d.SourceName == field:
                d.NumberFormat = code
                print(f"値フィールド書式: {pname}[{field}] = {code}"); print("（保存はしていません）"); return True
        print(f"エラー: 値フィールド '{field}' が見つかりません。"); return False

    if action == 'set-filter':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        wanted = set(rest[3:])
        if not wanted:
            print("表示する値を1つ以上指定してください。"); return False
        # 指定値が実在するか先に照合（全部タイポだと Excel が「全項目非表示」を拒否し、
        # 実状態と成功メッセージが食い違うため）
        item_names = []
        for i in range(1, p.PivotItems().Count + 1):
            item_names.append(p.PivotItems().Item(i).Name)
        missing = wanted - set(item_names)
        if missing:
            print(f"エラー: 存在しない値が指定されています: {sorted(missing)}")
            print(f"  このフィールドの値: {item_names}")
            return False
        failed = []
        # Excel は「表示ゼロ項目」を許さない。一巡ループで上から順に
        # Visible = (名前 in wanted) を代入すると、wanted の項目がまだ非表示のまま
        # 最後の表示項目を隠そうとして失敗し、その項目が表示に残る
        # （例: 項目[A,B,C]でA・B表示中に C だけを指定 → B が消せず B と C が表示）。
        # 先に「表示するもの」を出してから、「隠すもの」を隠す2パスにする。
        for i in range(1, p.PivotItems().Count + 1):
            it = p.PivotItems().Item(i)
            if it.Name in wanted:
                try:
                    it.Visible = True
                except Exception:
                    failed.append(it.Name)
        for i in range(1, p.PivotItems().Count + 1):
            it = p.PivotItems().Item(i)
            if it.Name not in wanted:
                try:
                    it.Visible = False
                except Exception:
                    failed.append(it.Name)
        if failed:
            print(f"⚠ 一部の項目の表示切替に失敗しました: {failed}")
            print("  （Excel の制約: 全項目非表示は不可、など。実際の表示状態を確認してください）")
        print(f"値フィルタ: {pname}[{field}] = {sorted(wanted)}"); print("（保存はしていません）"); return True

    if action == 'sort':
        order = (rest[3] if len(rest) >= 4 else 'asc').lower()
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.AutoSort(2 if order.startswith('d') else 1, p.Name)
        print(f"並べ替え: {pname}[{field}] = {'降順' if order.startswith('d') else '昇順'}")
        print("（保存はしていません）"); return True

    if action == 'group-date':
        interval = (rest[3] if len(rest) >= 4 else 'months').lower()
        # Periods: [秒,分,時,日,月,四半期,年]
        flags = {'days': 3, 'months': 4, 'quarters': 5, 'years': 6}
        if interval not in flags:
            print("interval は days|months|quarters|years"); return False
        periods = [False] * 7
        periods[flags[interval]] = True
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.DataRange.Cells(1, 1).Group(Periods=periods)
        print(f"日付グループ化: {pname}[{field}] = {interval}"); print("（保存はしていません）"); return True

    if action == 'group-numeric':
        if len(rest) < 6:
            print("使い方: pivot-field group-numeric <pivot> <field> <start> <end> <interval>"); return False
        try:
            start, end, step = float(rest[3]), float(rest[4]), float(rest[5])
        except ValueError:
            # 非数値をそのまま渡すと ValueError のトレースバックが出るだけで、
            # 何が悪かったのか分からない。使い方を出して止める。
            print("使い方: pivot-field group-numeric <pivot> <field> <start> <end> <interval>")
            print(f"  start/end/interval は数値で指定してください（指定値: {rest[3]} / {rest[4]} / {rest[5]}）")
            return False
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.DataRange.Cells(1, 1).Group(Start=start, End=end, By=step)
        print(f"数値グループ化: {pname}[{field}] = {start}〜{end} 刻み{step}"); print("（保存はしていません）"); return True

    print(f"未知のアクション: {action}")
    return False


@protect_safe
def cmd_pivot_calc(args):
    """ピボットの計算フィールド・レイアウト: pivot-calc <action> <pivot名> ...

      get-data <pivot>                                出力範囲の値を表示
      calc-field create <pivot> <名前> <数式>         計算フィールド作成（=Revenue-Cost 等）
      calc-field list <pivot>
      calc-field delete <pivot> <名前>
      layout <pivot> <compact|tabular|outline>        レポートレイアウト
      subtotals <pivot> <field> <on|off>              小計の表示
      grand-totals <pivot> <rows|cols|both> <on|off>  総計の表示
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: pivot-calc <action> <pivot名> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'calc-field':
        sub = rest[1].lower() if len(rest) >= 2 else ''
        pname = rest[2] if len(rest) >= 3 else None
        sh, pt = _find_pivot(wb, pname) if pname else (None, None)
        if not pt:
            print(f"エラー: ピボット '{pname}' が見つかりません。"); return False
        if sub == 'create':
            if len(rest) < 5:
                print("使い方: pivot-calc calc-field create <pivot> <名前> <数式>"); return False
            cf_name, formula = rest[3], rest[4]
            pt.CalculatedFields().Add(cf_name, formula)
            print(f"計算フィールド作成: {pname}[{cf_name}] = {formula}")
            print("  （値に表示するには pivot-field add-value で追加）")
            print("（保存はしていません）"); return True
        if sub == 'list':
            cfs = pt.CalculatedFields()
            print(f"--- {pname} の計算フィールド ({cfs.Count}) ---")
            for i in range(1, cfs.Count + 1):
                f = cfs.Item(i)
                try:
                    formula = f.Formula
                except Exception:
                    formula = ''
                print(f"  {f.Name} = {formula}")
            if cfs.Count == 0:
                print("  (なし)")
            return True
        if sub == 'delete':
            cf_name = rest[3] if len(rest) >= 4 else None
            if not cf_name:
                print("使い方: pivot-calc calc-field delete <pivot> <計算フィールド名>")
                return False
            try:
                pt.PivotFields(cf_name).Delete()
            except Exception:
                names = []
                try:
                    cfs = pt.CalculatedFields()
                    names = [cfs.Item(i).Name for i in range(1, cfs.Count + 1)]
                except Exception:
                    pass
                print(f"エラー: 計算フィールド '{cf_name}' を削除できません。")
                print("  存在する計算フィールド: " + (", ".join(names) or "(なし)"))
                return False
            print(f"計算フィールド削除: {pname}[{cf_name}]"); print("（保存はしていません）"); return True
        print("使い方: pivot-calc calc-field <create|list|delete> <pivot> ...")
        return False

    pname = rest[1] if len(rest) >= 2 else None
    sh, pt = _find_pivot(wb, pname) if pname else (None, None)
    if not pt:
        print(f"エラー: ピボット '{pname}' が見つかりません。"); return False

    if action == 'get-data':
        rng = pt.TableRange2
        print(f"ピボット出力範囲: {pt.Parent.Name}!{rng.Address}")
        data = rng.Value
        if data is not None:
            for row in data:
                cells = [('' if c is None else str(c)) for c in (row if isinstance(row, tuple) else [row])]
                print("  " + " | ".join(cells))
        return True

    if action == 'layout':
        lay = (rest[2] if len(rest) >= 3 else 'compact').lower()
        laymap = {'compact': 0, 'tabular': 1, 'outline': 2}   # xlCompactRow/xlTabularRow/xlOutlineRow
        if lay not in laymap:
            print("layout は compact|tabular|outline"); return False
        pt.RowAxisLayout(laymap[lay])
        print(f"レイアウト: {pname} = {lay}"); print("（保存はしていません）"); return True

    if action == 'subtotals':
        field = rest[2] if len(rest) >= 3 else None
        onoff = (rest[3] if len(rest) >= 4 else 'on').lower()
        try:
            p = pt.PivotFields(field)
        except Exception:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Subtotals = tuple([onoff == 'on'] + [False] * 11)   # 先頭=自動小計
        print(f"小計: {pname}[{field}] = {onoff}"); print("（保存はしていません）"); return True

    if action == 'grand-totals':
        which = (rest[2] if len(rest) >= 3 else 'both').lower()
        onoff = (rest[3] if len(rest) >= 4 else 'on').lower()
        val = (onoff == 'on')
        if which in ('rows', 'both'):
            pt.RowGrand = val
        if which in ('cols', 'both'):
            pt.ColumnGrand = val
        print(f"総計: {pname} {which} = {onoff}"); print("（保存はしていません）"); return True

    print(f"未知のアクション: {action}")
    return False


# ================================================================
# 重量級コマンド (3) スライサー
# ================================================================

def _find_pivot_or_table(wb, name):
    """名前からピボット or テーブル(ListObject)を探す。戻り値 (obj, kind, sheet) """
    for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
        for pt in sh.PivotTables():
            if pt.Name == name:
                return pt, 'pivot', sh
    for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
        for lo in sh.ListObjects:
            if lo.Name == name:
                return lo, 'table', sh
    return None, None, None


@protect_safe
def cmd_slicer(args):
    """スライサー操作: slicer <add|list|delete> ...

      slicer add <pivot名 or テーブル名> <フィールド> [--at セル] [--name 名]
      slicer list
      slicer delete <name>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: slicer <add|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sc in wb.SlicerCaches:
            for sl in sc.Slicers:
                cnt += 1
                # Slicer.Parent は SlicerCache を返す実装があるため、シート名は Shape 経由で取る
                try:
                    sheet_name = sl.Shape.Parent.Name
                except Exception:
                    try:
                        sheet_name = sl.Parent.Name
                    except Exception:
                        sheet_name = '?'
                print(f"{sl.Name}  (フィールド={sc.SourceName}, シート={sheet_name})")
        if cnt == 0:
            print("スライサーはありません。")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: slicer delete <name>"); return False
        name = rest[1]
        for sc in wb.SlicerCaches:
            for sl in sc.Slicers:
                if sl.Name == name:
                    sl.Delete()
                    print(f"スライサー削除: {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: スライサー '{name}' が見つかりません")
        return False

    if action == 'add':
        if len(rest) < 3:
            print("使い方: slicer add <pivot名 or テーブル名> <フィールド> [--at セル] [--name 名]")
            return False
        src_name, field = rest[1], rest[2]
        src, kind, sh = _find_pivot_or_table(wb, src_name)
        if src is None:
            print(f"エラー: ピボット/テーブル '{src_name}' が見つかりません")
            return False

        # 配置先の座標を先に解決する（Add2 の後に --at が不正で例外になると、
        # スライサー本体のない SlicerCache だけが孤児としてブックに残るため）
        at = getattr(args, 'at', None)
        dws = sh
        if at:
            anchor = dws.Range(at)
            top, left = anchor.Top, anchor.Left
        else:
            top, left = 10.0, 400.0

        sc = wb.SlicerCaches.Add2(src, field)
        try:
            sl = sc.Slicers.Add(SlicerDestination=dws, Caption=field,
                                Top=top, Left=left, Width=144.0, Height=180.0)
        except Exception as ex:
            # Add2 は通ったのに Slicers.Add が失敗すると（保護シート・座標不正等）、
            # スライサー本体のない SlicerCache だけが孤児として残り、以後
            # slicer list にも出ないまま再実行のたびに増える。作ったものは戻す。
            print(f"エラー: スライサーの配置に失敗しました: {ex}")
            try:
                sc.Delete()
                print("  作成途中のスライサーキャッシュは片づけました。")
            except Exception:
                print("  ⚠ スライサーキャッシュの後始末に失敗しました（保存前ならブックを閉じ直すのが確実です）。")
            return False
        # Slicers.Add の Name 引数は効かないことがあるので作成後に明示セット
        req_name = getattr(args, 'name', None)
        if req_name:
            eff_name = req_name.replace(' ', '')
            if eff_name != req_name:
                print(f"⚠ スライサー名のスペースは使えないため除去しました: '{req_name}' → '{eff_name}'")
            try:
                sl.Name = eff_name
            except Exception as ex:
                print(f"⚠ 名前 '{eff_name}' を設定できませんでした（{ex}）。自動名のままです。")
        print(f"スライサー追加: {sl.Name}  ソース={src_name}({kind})  フィールド={field}  シート={dws.Name}")
        print("（保存はしていません）")
        return True

    print(f"未知のアクション: {action}（add|list|delete）")
    return False


# ================================================================
# 計算モード (大量書き込みの高速化)
# ================================================================

def cmd_calc_mode(args):
    """計算モードの確認・切替・再計算

      calc-mode                 現在のモードを表示
      calc-mode manual          手動計算に（大量書込の前に）
      calc-mode auto            自動計算に戻す
      calc-mode recalc          今すぐ再計算（手動中の一括計算）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)

    names = {-4105: '自動 (automatic)', -4135: '手動 (manual)', 2: '半自動 (semiautomatic)'}

    if not rest:
        m = xl.Calculation
        print(f"現在の計算モード: {names.get(m, m)}")
        return True

    sub = rest[0].lower()
    if sub in ('auto', 'automatic'):
        xl.Calculation = -4105
        print("計算モード → 自動")
    elif sub == 'manual':
        xl.Calculation = -4135
        print("計算モード → 手動（書込後は calc-mode recalc / auto で再計算）")
    elif sub in ('recalc', 'now', 'calculate'):
        xl.Calculate()
        print("再計算しました")
    else:
        print(f"未知の指定: {sub}（manual|auto|recalc）")
        return False
    return True


# ================================================================
# 重量級コマンド (4) PowerQuery （一覧・更新・作成・M式書換・削除・読み込み配線）
# ================================================================

def _refresh_connection_sync(cn):
    """接続を「完了まで待って」更新する。

    WorkbookConnection.Refresh は BackgroundQuery=True（Excel が作る Query 接続では
    しばしば既定で True）だと即座に戻り、更新はバックグラウンドで走る。そのまま
    「更新しました」と報告すると、まだ古いデータのままなのに成功と言うことになり、
    直後の read-range が旧データを読む。更新中のエラーも戻り値・例外に出ない。
    そこで一時的に BackgroundQuery=False にして同期実行させ、元に戻す。
    """
    ole = None
    prev = None
    try:
        ole = cn.OLEDBConnection
        prev = bool(ole.BackgroundQuery)
        if prev:
            ole.BackgroundQuery = False
    except Exception:
        ole = None                      # OLEDB でない接続（テキスト等）はそのまま撃つ
    try:
        cn.Refresh()
    finally:
        if ole is not None and prev:
            try:
                ole.BackgroundQuery = prev
            except Exception:
                pass


def _connection_used_by_table(wb, cn_name):
    """接続 cn_name を使っているシートのテーブルがあれば "シート名!テーブル名" を返す。

    この接続を消すとそのテーブルの更新配線が切れる（更新できないテーブルが
    シートに残る）。削除する側は撃つ前に必ずここを通す。
    """
    try:
        for ws_chk in wb.Worksheets:
            for lo_chk in ws_chk.ListObjects:
                try:
                    if lo_chk.QueryTable.WorkbookConnection.Name == cn_name:
                        return f"{ws_chk.Name}!{lo_chk.Name}"
                except Exception:
                    continue     # QueryTable を持たない普通のテーブルはここに来る
    except Exception:
        pass
    return None


def _snapshot_connection(cn):
    """接続の再作成に要る情報を控える（削除をロールバックするため）。

    取れなかった項目は None。最低限 Name と接続文字列が取れなければ復元不能なので
    None を返し、呼び出し側は「消したまま失敗した」ことを明示する。
    """
    snap = {'name': None, 'desc': '', 'conn': None, 'cmd_text': None,
            'cmd_type': None, 'in_model': None}
    try:
        snap['name'] = cn.Name
    except Exception:
        return None
    try:
        snap['desc'] = cn.Description or ''
    except Exception:
        pass                              # 説明は欠けても接続の実体は変わらない
    # 以下は復元の同一性に効く項目。1つでも取れなければ「復元できる」と言ってはいけない。
    # まとめて try に入れると、接続文字列だけ取れて CommandText/CommandType が
    # 欠けたまま初期値で復元し、別物の接続を作って「元に戻しました」と報告してしまう
    # （PowerQuery のモデル接続は CommandType=6。既定の 1 で復元すると配線が変わる）
    try:
        snap['in_model'] = bool(cn.InModel)
    except Exception:
        return None
    try:
        sub = cn.OLEDBConnection            # PowerQuery 接続は OLEDB
    except Exception:
        return None
    for key, get in (('conn', lambda: str(sub.Connection)),
                     ('cmd_text', lambda: sub.CommandText),
                     ('cmd_type', lambda: int(sub.CommandType))):
        try:
            snap[key] = get()
        except Exception:
            return None
    if not snap['conn']:
        return None
    return snap


def _restore_connection(wb, snap):
    """_snapshot_connection で控えた接続を作り直す。成功したら True"""
    try:
        wb.Connections.Add2(snap['name'], snap['desc'], snap['conn'],
                            snap['cmd_text'], snap['cmd_type'],
                            snap['in_model'], False)
        return True
    except Exception:
        return False


@protect_safe
@dialog_safe
def cmd_powerquery(args):
    """PowerQuery: powerquery <list|refresh|add|edit|delete|load> ...

      powerquery list                 クエリと接続の一覧（M式の行数つき）
      powerquery refresh              全クエリ/接続を更新 (RefreshAll)
      powerquery refresh <name>       指定クエリ/接続を更新
      powerquery add <name>           M式から新規クエリ作成（接続のみ）
                                      M式は --m / --m-file / _last_query.m
      powerquery edit <name>          既存クエリのM式を書き換え（M式の指定は add と同じ）
      powerquery delete <name>        クエリを削除
      powerquery load <name> --to sheet [--sheet S] [--at A1]   シートのテーブルに読み込み
      powerquery load <name> --to model                          データモデルに読み込み
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: powerquery <list|refresh [name]|add <name>|delete <name>>")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        # クエリ一覧
        try:
            qs = wb.Queries
            qcount = qs.Count
        except Exception:
            qs = None
            qcount = 0
        if qs and qcount > 0:
            print(f"--- PowerQuery クエリ ({qcount}) ---")
            for i in range(1, qcount + 1):
                q = qs.Item(i)
                try:
                    desc = q.Description or ''
                except Exception:
                    desc = ''
                # M式の行数を補助表示
                try:
                    nlines = len(str(q.Formula).replace('\r\n', '\n').split('\n'))
                except Exception:
                    nlines = '?'
                print(f"  {q.Name}  (M式 {nlines}行)" + (f"  - {desc}" if desc else ""))
        else:
            print("PowerQuery クエリはありません。")
        # 接続一覧（更新対象の確認用）
        try:
            conns = wb.Connections
            ccount = conns.Count
        except Exception:
            ccount = 0
        if ccount > 0:
            print(f"--- 接続 ({ccount}) ---")
            for cn in conns:
                print(f"  {cn.Name}")
        return True

    if action == 'refresh':
        if len(rest) >= 2:
            name = rest[1]
            target_conn = None
            for cn in wb.Connections:
                if cn.Name == name or cn.Name == f"Query - {name}":
                    target_conn = cn
                    break
            if target_conn:
                # 完了を待って更新する（待たないと「更新しました」と言った時点では
                # まだ古いデータのまま＝誤った成功報告になる）
                _refresh_connection_sync(target_conn)
                print(f"更新しました: {target_conn.Name}")
                return True
            # 接続が無い（読み込みなしクエリ等）
            print(f"接続 '{name}' が見つかりません。")
            print("  （読み込みなしクエリは更新対象がありません。powerquery list で名前を確認）")
            return False
        else:
            wb.RefreshAll()
            print("全クエリ/接続を更新しました (RefreshAll)")
            print("  ※ バックグラウンド更新の場合、完了まで数秒かかることがあります。")
            return True

    if action == 'add':
        if len(rest) < 2:
            print("使い方: powerquery add <name> [--m-file f | --m \"M式\"]")
            return False
        name = rest[1]
        # M式の取得: --m インライン > --m-file > _last_query.m
        m_inline = getattr(args, 'm_opt', None)
        if m_inline:
            formula = m_inline
        else:
            mf = getattr(args, 'm_file', None)
            path = smart_path_resolve(mf) if mf else _LAST_QUERY_FILE
            if not path or not os.path.exists(path):
                print(f"エラー: M式ファイルが見つかりません: {mf or _LAST_QUERY_FILE}")
                print("  _last_query.m にM式を書くか、--m-file / --m を指定してください。")
                return False
            formula = read_code_file(path)
        if not formula or not formula.strip():
            print("エラー: M式が空です。")
            return False
        # 重複チェック
        try:
            existing = [wb.Queries.Item(i).Name for i in range(1, wb.Queries.Count + 1)]
        except Exception:
            existing = []
        if name in existing:
            print(f"エラー: クエリ '{name}' は既に存在します（delete してから add）。")
            return False
        desc = getattr(args, 'desc', None) or ''
        wb.Queries.Add(name, formula, desc)
        print(f"クエリ作成: {name}（接続のみ。シート/モデルへの読み込みは別途）")
        print("（保存はしていません）")
        return True

    if action == 'edit':
        if len(rest) < 2:
            print("使い方: powerquery edit <name> [--m-file f | --m \"M式\"]")
            return False
        name = rest[1]
        # M式の取得（add と同じ）: --m > --m-file > _last_query.m
        m_inline = getattr(args, 'm_opt', None)
        if m_inline:
            formula = m_inline
        else:
            mf = getattr(args, 'm_file', None)
            path = smart_path_resolve(mf) if mf else _LAST_QUERY_FILE
            if not path or not os.path.exists(path):
                print(f"エラー: M式ファイルが見つかりません: {mf or _LAST_QUERY_FILE}")
                return False
            formula = read_code_file(path)
        if not formula or not formula.strip():
            print("エラー: M式が空です。")
            return False
        try:
            cnt = wb.Queries.Count
        except Exception:
            cnt = 0
        for i in range(1, cnt + 1):
            q = wb.Queries.Item(i)
            if q.Name == name:
                q.Formula = formula                # WorkbookQuery.Formula は書込可（検証済）
                print(f"クエリ書き換え: {name}")
                print("（保存はしていません。反映には powerquery refresh が必要なことがあります）")
                return True
        print(f"エラー: クエリ '{name}' が見つかりません")
        return False

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: powerquery delete <name>"); return False
        name = rest[1]
        try:
            cnt = wb.Queries.Count
        except Exception:
            cnt = 0
        for i in range(1, cnt + 1):
            if wb.Queries.Item(i).Name == name:
                wb.Queries.Item(i).Delete()
                print(f"クエリ削除: {name}")
                print("（保存はしていません）")
                return True
        print(f"エラー: クエリ '{name}' が見つかりません")
        return False

    if action == 'load':
        # 接続のみクエリを「シートのテーブル」または「データモデル」に読み込む配線。
        #   powerquery load <name> --to sheet  [--sheet S] [--at A1]
        #   powerquery load <name> --to model
        if len(rest) < 2:
            print('使い方: powerquery load <name> --to sheet|model [--sheet S] [--at A1]')
            return False
        name = rest[1]
        to = (getattr(args, 'to', None) or 'sheet').lower()
        # クエリ存在チェック
        try:
            existing = [wb.Queries.Item(i).Name for i in range(1, wb.Queries.Count + 1)]
        except Exception:
            existing = []
        if name not in existing:
            print(f"エラー: クエリ '{name}' が見つかりません（powerquery list で確認）。")
            return False

        # Power Query (Mashup) の OLEDB 接続文字列 — 記録マクロが生成する形に合わせる
        conn_str = ("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
                    f'Location={name};Extended Properties=""')
        cmd_text = f"SELECT * FROM [{name}]"

        if to == 'sheet':
            # 出力先シート（--sheet 省略時はアクティブシート）
            sheet_name = getattr(args, 'sheet', None)
            ws = None
            created_ws = None
            if sheet_name:
                for i in range(1, wb.Worksheets.Count + 1):
                    # Excel のシート名重複判定は大文字小文字を区別しない
                    if wb.Worksheets.Item(i).Name.lower() == sheet_name.lower():
                        ws = wb.Worksheets.Item(i); break
                if ws is None:
                    ws = wb.Worksheets.Add()
                    created_ws = ws
                    try:
                        ws.Name = sheet_name
                    except Exception as ex:
                        # 禁止文字(/ 等)・31文字超で失敗すると無名シートが残骸になる
                        print(f"エラー: シート名 '{sheet_name}' を設定できません: {ex}")
                        xl.DisplayAlerts = False
                        try:
                            ws.Delete()
                        finally:
                            xl.DisplayAlerts = True
                        return False
            else:
                ws = wb.ActiveSheet
            at = getattr(args, 'at', None) or 'A1'
            dest = ws.Range(at)
            # 0 = xlSrcExternal。Source に Mashup の OLEDB 文字列を渡す
            lo = ws.ListObjects.Add(0, conn_str, None, True, dest)
            qt = lo.QueryTable
            qt.CommandType = 2                 # xlCmdSql
            qt.CommandText = cmd_text
            qt.RowNumbers = False
            qt.FillAdjacentFormulas = False
            qt.PreserveFormatting = True
            qt.RefreshOnFileOpen = False
            qt.BackgroundQuery = False
            qt.AdjustColumnWidth = True
            try:
                # QueryTable.Refresh は Boolean を返す。失敗しても例外を投げず
                # False を返す経路があり、捨てると「シートに読み込みました」と
                # 成功報告したまま空／エラー列だけのテーブルが残る
                if qt.Refresh(False) is False:  # BackgroundQuery:=False
                    raise RuntimeError("QueryTable.Refresh が False を返しました"
                                       "（M式の実行時エラーの可能性）")
            except Exception as ex:
                # M式の実行時エラー等で失敗すると、追加済みの ListObject と
                # 自動生成の接続が孤児として残り、再実行のたび「テーブル1/2…」と
                # 増殖する。このコマンドが作ったものだけ片づける
                print(f"エラー: クエリの読み込みに失敗しました: {ex}")
                print("  M式の実行時エラーの可能性があります"
                      "（powerquery list でクエリを確認、powerquery edit で書き換え）。")
                try:
                    wbconn = qt.WorkbookConnection
                except Exception:
                    wbconn = None
                try:
                    lo.Delete()
                except Exception:
                    pass
                try:
                    if wbconn is not None:
                        wbconn.Delete()
                except Exception:
                    pass
                if created_ws is not None:
                    xl.DisplayAlerts = False
                    try:
                        created_ws.Delete()
                    except Exception:
                        pass
                    finally:
                        xl.DisplayAlerts = True
                print("  追加途中のテーブル・接続は片づけました。")
                return False
            # リネームの失敗を握りつぶすと、成功メッセージだけ出て
            # 以後 powerquery refresh <name> で引けない（名前が汎用名のまま）状態になる。
            # 読み込み自体は成功しているので中止はせず、警告として明示する。
            try:
                lo.Name = name
            except Exception as ex:
                print(f"⚠ テーブル名を '{name}' にできませんでした（{ex}）。自動名 '{lo.Name}' のままです。")
            # 既定では「接続」等の汎用名が付く。refresh <name> で引けるよう
            # Excel 標準の "Query - <name>" に揃える。
            conn_name = None
            try:
                wbconn = qt.WorkbookConnection
                if wbconn is not None:
                    wbconn.Name = f"Query - {name}"
                    conn_name = wbconn.Name
            except Exception as ex:
                print(f"⚠ 接続名を 'Query - {name}' にできませんでした（{ex}）。")
                print(f"   powerquery refresh {name} では引けません。"
                      "connection list で実際の接続名を確認し、その名前で refresh してください。")
            print(f"シートに読み込みました: {name} → {ws.Name}!{at}（テーブル: {lo.Name}"
                  + (f", 接続: {conn_name}" if conn_name else "") + "）")
            print("（保存はしていません）")
            return True

        if to == 'model':
            # データモデル（Power Pivot）へ。Queries.Add が作る "Query - name"
            # 接続が残っていると衝突するので、あれば作り直す。
            # ただしその接続がシートのテーブル（--to sheet の読み込み）に使われている
            # 場合、削除するとシート側の更新配線が壊れるため停止する。
            cn_name = f"Query - {name}"
            deleted_snap = None       # 消した既存接続の控え（Add2 失敗時のロールバック用）
            deleted_any = False       # 控えの有無と別に「実際に消したか」を持つ
            delete_failed = None      # 既存接続の Delete が失敗した事実（後段の誤診断防止）
            for cn in list(wb.Connections):
                if cn.Name == cn_name:
                    used_by = _connection_used_by_table(wb, cn_name)
                    if used_by:
                        print(f"エラー: 接続 '{cn_name}' はシートのテーブル {used_by} が使用中です。")
                        print("  削除するとテーブルの更新ができなくなるため中止しました。")
                        print("  モデルにも読み込みたい場合は、シート読み込みを解除してから実行してください。")
                        return False
                    # _connection_used_by_table はシートのテーブルしか見ない。
                    # 既にモデルに載っている接続（InModel）は「未使用」と判定されて
                    # そのまま Delete され、そのテーブルに紐づくメジャー・
                    # リレーションシップが道連れになる。しかも Add2 で作り直すので
                    # 「データモデルに読み込みました」と成功報告してしまう。
                    in_model = False
                    try:
                        in_model = bool(cn.InModel)
                    except Exception:
                        in_model = False
                    if in_model and not getattr(args, 'force', False):
                        print(f"エラー: '{cn_name}' は既にデータモデルに読み込まれています。")
                        print("  作り直すと、このテーブルに紐づくメジャーとリレーションシップが")
                        print("  失われるため中止しました（更新するだけなら powerquery refresh）。")
                        print("  承知のうえで作り直すなら --force を付けてください。")
                        return False
                    # 消す前に中身を控える。Add2 が失敗したときに元へ戻せないと、
                    # 「接続だけ消えてモデルにも載っていない＝クエリが未配線」の
                    # 一番たちの悪い状態で終わるため。
                    old_snap = _snapshot_connection(cn)
                    try:
                        cn.Delete()
                        deleted_any = True
                        deleted_snap = old_snap
                    except Exception as ex_del:
                        # 消せなかった事実は控える。ここで黙ると、後段の Add2 が
                        # 同名衝突で失敗したときに「M式のエラーかも」と誤診断する
                        delete_failed = str(ex_del)
            # Connections.Add2(Name, Description, ConnectionString, CommandText,
            #                  lCmdtype, CreateModelConnection, ImportRelationships)
            # モデル読込は記録マクロ形式に合わせる: CommandText=クエリ名,
            # lCmdtype=6 (xlCmdTableCollection)。これでモデルテーブル名が
            # クエリ名になる（SQL/SELECT形式だと "クエリ" の汎用名になる）。
            try:
                wb.Connections.Add2(cn_name, "", conn_str, name, 6, True, False)
            except Exception as ex:
                # M式の実行時エラー・モデル非対応等で失敗しうる。丸腰で呼ぶと
                # 上で消した既存接続が戻らない。控えた内容で作り直す。
                print(f"エラー: データモデルへの読み込みに失敗しました: {ex}")
                if delete_failed:
                    # 既存の同名接続が消せていない＝Add2 失敗の真因はまず名前衝突。
                    # ここで「M式のエラーかも」と言うと誤誘導になる
                    print(f"  既存の接続 '{cn_name}' を削除できていません（{delete_failed}）。")
                    print("  同名の接続が残っているため作り直せなかった可能性が高いです。")
                else:
                    print("  M式の実行時エラー、またはこのブックがデータモデル非対応の可能性があります。")
                if deleted_snap is None:
                    if deleted_any:
                        # 「何も消していない」と「消したが控えが取れず戻せない」は別物。
                        # 黙って帰ると、接続が消えた事実が一切報告されない
                        print(f"  ⚠ 既存の接続 '{cn_name}' は削除済みで、控えが取れなかったため復旧できません。")
                        print("     クエリ自体は残っています。配線をやり直すには:")
                        print(f"       powerquery load {name} --to sheet   （シートに読み込む場合）")
                        print(f"       powerquery load {name} --to model   （モデルに読み込む場合）")
                    return False
                if _restore_connection(wb, deleted_snap):
                    print(f"  既存の接続 '{cn_name}' は元に戻しました（クエリの配線は元のままです）。")
                else:
                    print(f"  ⚠ 既存の接続 '{cn_name}' を消したまま復旧できませんでした。")
                    print("     クエリ自体は残っています。配線をやり直すには:")
                    print(f"       powerquery load {name} --to sheet   （シートに読み込む場合）")
                    print(f"       powerquery load {name} --to model   （モデルに読み込む場合）")
                    print("     現状は connection list / datamodel list で確認できます。")
                return False
            print(f"データモデルに読み込みました: {name}")
            print("（保存はしていません。datamodel list で確認できます）")
            return True

        print(f"未知の読み込み先: {to}（sheet|model）")
        return False

    print(f"未知のアクション: {action}（list|refresh|add|edit|delete|load）")
    return False


# ================================================================
# 重量級コマンド (5) コネクション / データモデル （管理・読み取り）
# ================================================================

# XlConnectionType: xlConnectionTypeOLEDB=1, ODBC=2, XMLMAP=3, TEXT=4, WEB=5,
#                   DATAFEED=6, MODEL=7, WORKSHEET=8, NOSOURCE=9
_XL_CONN_TYPE = {1: 'OLEDB', 2: 'ODBC', 3: 'XMLMAP', 4: 'TEXT',
                 5: 'WEB', 6: 'DATAFEED', 7: 'MODEL', 8: 'WORKSHEET', 9: 'NOSOURCE'}


@protect_safe
@dialog_safe
def cmd_connection(args):
    """ブック接続の管理: connection <list|refresh|delete> [name]

      connection list                クエリ/外部データ接続の一覧（種別・接続文字列）
      connection refresh [name]      接続を更新（name 省略で全件 RefreshAll）
      connection delete <name>       接続を削除（使用中のテーブルがあれば注記を出す）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: connection <list|refresh|delete> [name]")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        conns = wb.Connections
        n = conns.Count
        if n == 0:
            print("接続はありません。")
            return True
        print(f"--- ブック接続 ({n}) ---")
        for cn in conns:
            try:
                t = _XL_CONN_TYPE.get(int(cn.Type), cn.Type)
            except Exception:
                t = '?'
            print(f"  {cn.Name}  [{t}]")
            try:
                if cn.Description:
                    print(f"      説明: {cn.Description}")
            except Exception:
                pass
            # 接続文字列・コマンド（OLEDB/ODBC）
            try:
                sub = None
                if int(cn.Type) == 1:
                    sub = cn.OLEDBConnection
                elif int(cn.Type) == 2:
                    sub = cn.ODBCConnection
                if sub is not None:
                    cs = str(sub.Connection)
                    print(f"      接続: {cs[:100]}{'…' if len(cs) > 100 else ''}")
            except Exception:
                pass
        return True

    if action == 'refresh':
        if len(rest) >= 2:
            name = rest[1]
            for cn in wb.Connections:
                if cn.Name == name or cn.Name == f"Query - {name}":
                    cn.Refresh()
                    print(f"更新しました: {cn.Name}")
                    return True
            print(f"エラー: 接続 '{name}' が見つかりません")
            return False
        wb.RefreshAll()
        print("全接続を更新しました (RefreshAll)")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: connection delete <name>"); return False
        name = rest[1]
        for cn in wb.Connections:
            if cn.Name == name or cn.Name == f"Query - {name}":
                actual = cn.Name           # Delete 後は参照不可になるので退避
                # 名指しされた接続はそのまま消す（シートのテーブルが使っていても、
                # シートのデータ自体は残る＝更新の配線が外れるだけ。使用中の注記は
                # 出すが、消すか消さないかの判断はツールの仕事ではない）
                used_by = _connection_used_by_table(wb, actual)
                cn.Delete()
                print(f"接続を削除: {actual}")
                if used_by:
                    print(f"  注記: シートのテーブル {used_by} がこの接続を使っていました。"
                          "以後そのテーブルは更新できません（データは残っています）。")
                print("（保存はしていません）")
                return True
        print(f"エラー: 接続 '{name}' が見つかりません")
        return False

    print(f"未知のアクション: {action}（list|refresh|delete）")
    return False


@protect_safe
def cmd_datamodel(args):
    """データモデル: datamodel <list|relation|measure>

      datamodel list   モデルのテーブル・リレーションシップ・メジャーを一覧
      datamodel relation add    <FKテーブル> <FK列> <PKテーブル> <PK列>   リレーション作成
      datamodel relation delete <FKテーブル> <FK列> <PKテーブル> <PK列>   リレーション削除
      datamodel measure add <テーブル> <メジャー名> --dax "式" [--format general|whole|decimal|currency|percent|scientific]
                                                                 [--decimals N] [--thousands] [--symbol JPY]   メジャー(DAX)作成
      datamodel measure delete <メジャー名>                              メジャー削除
      （※ テーブルの追加は powerquery load --to model）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    action = rest[0].lower() if rest else 'list'
    xl, wb = get_workbook(target_file)

    try:
        model = wb.Model
    except Exception:
        print("このブックはデータモデルに対応していません。")
        # list 系は「無い」を正常報告でよいが、追加/削除の要求は実行されて
        # いないので失敗として返す（batch やスクリプト連携で握りつぶさない）
        return action in ('list', 'tables', 'measures', 'relations')

    # --- リレーションシップの作成・削除 ---
    #   datamodel relation add    <FKテーブル> <FK列> <PKテーブル> <PK列>
    #   datamodel relation delete <FKテーブル> <FK列> <PKテーブル> <PK列>
    if action in ('relation', 'rel', 'relationship'):
        sub = rest[1].lower() if len(rest) >= 2 else ''
        if sub in ('add', 'delete'):
            if len(rest) < 6:
                print(f'使い方: datamodel relation {sub} <FKテーブル> <FK列> <PKテーブル> <PK列>')
                print('  FK=多側(参照する側) / PK=一側(参照される側)')
                return False
            fkt_name, fkc_name, pkt_name, pkc_name = rest[2], rest[3], rest[4], rest[5]
        if sub == 'add':
            try:
                fkt = model.ModelTables.Item(fkt_name)
                pkt = model.ModelTables.Item(pkt_name)
            except Exception:
                print(f"エラー: テーブルが見つかりません（{fkt_name} / {pkt_name}）。datamodel list で確認。")
                return False
            try:
                fkc = fkt.ModelTableColumns.Item(fkc_name)
                pkc = pkt.ModelTableColumns.Item(pkc_name)
            except Exception:
                print(f"エラー: 列が見つかりません（{fkt_name}[{fkc_name}] / {pkt_name}[{pkc_name}]）。")
                return False
            model.ModelRelationships.Add(fkc, pkc)
            print(f"リレーション作成: {fkt_name}[{fkc_name}] → {pkt_name}[{pkc_name}]")
            print("（保存はしていません）")
            return True
        if sub == 'delete':
            rels = model.ModelRelationships
            for i in range(1, rels.Count + 1):
                r = rels.Item(i)
                try:
                    if (r.ForeignKeyTable.Name == fkt_name and r.ForeignKeyColumn.Name == fkc_name
                            and r.PrimaryKeyTable.Name == pkt_name and r.PrimaryKeyColumn.Name == pkc_name):
                        r.Delete()
                        print(f"リレーション削除: {fkt_name}[{fkc_name}] → {pkt_name}[{pkc_name}]")
                        print("（保存はしていません）")
                        return True
                except Exception:
                    continue
            print("エラー: 該当するリレーションが見つかりません（datamodel list で確認）。")
            return False
        print('使い方: datamodel relation <add|delete> <FKテーブル> <FK列> <PKテーブル> <PK列>')
        return False

    # --- メジャー(DAX)の作成・削除 ---
    #   datamodel measure add <テーブル> <メジャー名> [--dax "式" | --dax-file f | _last_dax.dax]
    #   datamodel measure delete <メジャー名>
    if action in ('measure', 'measures'):
        sub = rest[1].lower() if len(rest) >= 2 else ''
        if sub == 'add':
            if len(rest) < 4:
                print('使い方: datamodel measure add <テーブル> <メジャー名> --dax "DAX式"')
                print('  DAX は --dax / --dax-file / _last_dax.dax(UTF-8) から取得。')
                print('  ※ 先頭の = は不要。日本語テーブル名は DAX 内でシングルクォート: SUM(\'売上\'[数量])')
                return False
            tbl_name, measure_name = rest[2], rest[3]
            # DAX の取得: --dax インライン > --dax-file > _last_dax.dax
            dax = getattr(args, 'dax', None)
            if not dax:
                df = getattr(args, 'dax_file', None)
                path = smart_path_resolve(df) if df else _LAST_DAX_FILE
                if not path or not os.path.exists(path):
                    print(f"エラー: DAXファイルが見つかりません: {df or _LAST_DAX_FILE}")
                    print("  _last_dax.dax に式を書くか、--dax / --dax-file を指定してください。")
                    return False
                dax = read_code_file(path)
            if not dax or not dax.strip():
                print("エラー: DAX式が空です。")
                return False
            dax = dax.strip()
            if dax.startswith('='):            # Excel数式の癖で = を付けても通るように
                dax = dax[1:].strip()
            try:
                tbl = model.ModelTables.Item(tbl_name)
            except Exception:
                print(f"エラー: テーブル '{tbl_name}' が見つかりません。datamodel list で確認。")
                return False
            # 数値書式（既定 general）。引数付き書式は GetModelFormat* メソッドで取得
            #   （ModelFormat* プロパティは既定値専用で引数を渡せないため）。
            fmt_name = (getattr(args, 'format', None) or 'general').lower()
            dec_arg = getattr(args, 'decimals', None)
            try:
                decimals = int(dec_arg) if dec_arg is not None else 2
            except (TypeError, ValueError):
                print(f"エラー: --decimals は数値で指定してください: '{dec_arg}'")
                return False
            thousands = bool(getattr(args, 'thousands', False))
            symbol = getattr(args, 'symbol', None) or ''
            try:
                if fmt_name == 'general':
                    fmt = model.ModelFormatGeneral
                elif fmt_name in ('whole', 'wholenumber'):
                    fmt = model.GetModelFormatWholeNumber(thousands)
                elif fmt_name in ('decimal', 'decimalnumber'):
                    fmt = model.GetModelFormatDecimalNumber(thousands, decimals)
                elif fmt_name == 'currency':
                    # Symbol は通貨コード（USD/JPY/EUR 等）。グリフ（$ 等）は無効だが
                    # GetModelFormatCurrency では落ちず Add 時に例外になるため、
                    # フォールバックは Add 側で行う。
                    fmt = model.GetModelFormatCurrency(symbol, decimals)
                elif fmt_name in ('percent', 'percentage'):
                    fmt = model.GetModelFormatPercentageNumber(thousands, decimals)
                elif fmt_name in ('scientific', 'sci'):
                    fmt = model.GetModelFormatScientificNumber(decimals)
                else:
                    print(f"エラー: 未知の書式 '{fmt_name}'（general|whole|decimal|currency|percent|scientific）")
                    return False
            except Exception as e:
                print(f"エラー: 書式オブジェクトの取得に失敗: {str(e)[:120]}")
                return False
            desc = getattr(args, 'desc', None) or ''
            try:
                model.ModelMeasures.Add(measure_name, tbl, dax, fmt, desc)
            except Exception as e:
                # currency でグリフ等の無効な通貨コードだと Add 時に例外。
                # 既定の通貨記号で 1 回だけ再試行する。
                if fmt_name == 'currency' and symbol:
                    try:
                        model.ModelMeasures.Add(measure_name, tbl, dax,
                                                model.GetModelFormatCurrency('', decimals), desc)
                        print(f"  注意: 通貨コード '{symbol}' は無効。既定の通貨記号で作成しました（有効例: USD, JPY, EUR）。")
                        print(f"メジャー作成: {tbl_name}[{measure_name}] = {dax}  (書式=currency)")
                        print("（保存はしていません）")
                        return True
                    except Exception as e2:
                        e = e2
                print(f"エラー: メジャー作成に失敗しました: {str(e)[:200]}")
                print("  DAX 構文・テーブル/列名・シングルクォートを確認してください。")
                return False
            print(f"メジャー作成: {tbl_name}[{measure_name}] = {dax}  (書式={fmt_name})")
            print("（保存はしていません）")
            return True
        if sub == 'delete':
            if len(rest) < 3:
                print('使い方: datamodel measure delete <メジャー名>')
                return False
            measure_name = rest[2]
            ms = model.ModelMeasures
            for i in range(1, ms.Count + 1):
                if ms.Item(i).Name == measure_name:
                    ms.Item(i).Delete()
                    print(f"メジャー削除: {measure_name}")
                    print("（保存はしていません）")
                    return True
            print(f"エラー: メジャー '{measure_name}' が見つかりません（datamodel list で確認）。")
            return False
        print('使い方: datamodel measure <add|delete> ...')
        return False

    if action != 'list':
        print(f"未知のアクション: {action}（list|relation|measure）")
        return False

    # テーブル
    try:
        mts = model.ModelTables
        tn = mts.Count
    except Exception:
        mts = None
        tn = 0
    print(f"--- データモデル: テーブル ({tn}) ---")
    for i in range(1, tn + 1):
        mt = mts.Item(i)
        try:
            rc = mt.RecordCount
        except Exception:
            rc = '?'
        print(f"  {mt.Name}  ({rc}行)")
    if tn == 0:
        print("  (なし)")

    # リレーションシップ
    try:
        rels = model.ModelRelationships
        rn = rels.Count
    except Exception:
        rels = None
        rn = 0
    print(f"--- リレーションシップ ({rn}) ---")
    for i in range(1, rn + 1):
        r = rels.Item(i)
        try:
            fkt = r.ForeignKeyTable.Name
            fkc = r.ForeignKeyColumn.Name
            pkt = r.PrimaryKeyTable.Name
            pkc = r.PrimaryKeyColumn.Name
            active = ''
            try:
                active = '' if r.Active else '  (無効)'
            except Exception:
                pass
            print(f"  {fkt}[{fkc}] → {pkt}[{pkc}]{active}")
        except Exception:
            print(f"  (リレーション {i}: 読み取り不可)")
    if rn == 0:
        print("  (なし)")

    # メジャー（対応バージョンのみ）
    try:
        ms = model.ModelMeasures
        mn = ms.Count
        print(f"--- メジャー ({mn}) ---")
        for i in range(1, mn + 1):
            me = ms.Item(i)
            try:
                tbl = me.AssociatedTable.Name
            except Exception:
                tbl = '?'
            print(f"  {me.Name}  (所属={tbl})")
        if mn == 0:
            print("  (なし)")
    except Exception:
        pass

    return True




__all__ = [
    '_XL_AXIS',
    '_XL_CHART_TYPE',
    '_XL_CHART_TYPE_NAME',
    '_XL_CONN_TYPE',
    '_XL_LEGEND_POS',
    '_XL_PIVOT_FUNC',
    '_XL_PIVOT_ORIENT',
    '_XL_TRENDLINE',
    '_connection_used_by_table',
    '_find_pivot',
    '_find_pivot_or_table',
    '_restore_connection',
    '_snapshot_connection',
    '_unique_sheet_name',
    'cmd_calc_mode',
    'cmd_chart',
    'cmd_chart_config',
    'cmd_connection',
    'cmd_datamodel',
    'cmd_pivot',
    'cmd_pivot_calc',
    'cmd_pivot_field',
    'cmd_powerquery',
    'cmd_slicer',
]
