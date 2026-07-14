# -*- coding: utf-8 -*-
r"""
form_layout.py — 宣言的レイアウトで UserForm を組み立てる（座標計算はすべて機械）

「何を並べるか」（判断）は呼び出し側、「どこに置くか」（計算）はここ、の分離。
行構造を書くだけで、ラベル列幅の自動揃え・8ptリズムの余白・ボタンバー右寄せ・
TabIndex の視線順整列・Default/Cancel 設定までを機械計算する。
デザイン規範は excel-userform-builder スキルの「デザイン原則」に従う。

使い方:
    from form_layout import (build_form, preview_layout, generate_vba_stub,
                             row, lbl, txt, combo, lst, chk, opt, opt_group,
                             btn, ok, cancel, button_bar, spacer, heading, frame)

    rows = [
        heading("顧客情報"),
        row(lbl("名前"), txt("txtName")),
        row(lbl("区分"), combo("cmbKind", items=["法人", "個人"])),
        frame("オプション",
              row(opt_group(("optA", "通常"), ("optB", "急ぎ")))),
        spacer(),
        button_bar(ok("btnOK", "登録"), cancel("btnCancel", "閉じる")),
    ]
    preview_layout(rows)                 # Excel 不要のワイヤーフレームPNG（設計の試行錯誤用）
    build_form("F_Input", "データ入力", rows,
               vba_stub=True, png=True)  # 実構築＋イベント雛形注入＋実表示PNG

- 幅未指定の txt/combo/lst は入力列いっぱいに自動ストレッチ（右端が揃う）
- 先頭要素が lbl の行はラベル列として幅を自動で揃え、入力と上下中央合わせ
- button_bar は右下寄せ・全ボタン同サイズ（Windows 規約: キャンセルが右端）
- frame は入れ子1段まで（frame の中に frame は不可）
- 既存フォームを作り直す前に .frm/.frx を backups へ自動退避する
- vba_stub=True で Initialize（items の AddItem・先頭入力へ SetFocus）と
  各ボタンの Click 雛形を機械生成して注入（vba_file 指定があればそちら優先）
- 自由配置（カレンダー格子等）はこのモジュールでなく form_builder の add_* + Grid を使う
"""
import os
import sys
import time
import unicodedata
import zlib

from form_builder import FormBuilder, check_control_name

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', 'backups'))
PREVIEW_FILE = os.path.join(SCRIPT_DIR, '_last_layout_preview.png')
STUB_FILE = os.path.join(SCRIPT_DIR, '_last_form_code.vba')

# ================================================================
# スタイル定数（8の倍数リズム・一元管理）
# ================================================================

STYLE = {
    'pad':        12,    # フォーム外周の余白
    'gap_x':      8,     # 行内の横間隔
    'gap_y':      8,     # 行間
    'ctrl_h':     22,    # TextBox/ComboBox の標準高さ
    'lbl_h':      18,    # ラベルの標準高さ
    'btn_w':      72,    # ボタンの最小幅
    'btn_h':      24,    # ボタンの標準高さ
    'input_w':    160,   # ストレッチ入力の最小幅
    'font':       12,    # 標準フォントサイズ
    'heading_pt': 13,    # 見出しのフォントサイズ（Bold）
    'frame_pad':  8,     # Frame 内側の余白
    'frame_top':  14,    # Frame のキャプション帯ぶんの上余白
    'chrome_w':   11,    # フォーム Width と InsideWidth の差（実測値）
    'chrome_h':   29,    # フォーム Height と InsideHeight の差（実測値）
    'page_pad':   8,     # MultiPage の各ページ内側の余白
    'tab_h':      18,    # MultiPage のタブ帯の高さ
    'mp_chrome':  6,     # MultiPage の枠ぶんの余白（左右合計・下）
}


def _text_w(s, font=None):
    """文字列の概算表示幅（pt）。全角=フォントサイズ、半角=0.55倍で見積もる"""
    font = font or STYLE['font']
    w = 0.0
    for ch in s or '':
        w += font if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else font * 0.55
    return int(w) + 8


# ================================================================
# 要素（宣言）。すべて dict を返すだけの純粋関数
# ================================================================

def lbl(caption, width=None, bold=False, name=None, tab_index=None):
    """ラベル。行の先頭に置くとラベル列として自動整列。

    name 省略時は build 時に lbl1, lbl2... と機械採番する。VBA から
    `lblStatus.Caption = ...` のように名指しするラベルは name を明示すること
    （inspect --to-layout は実名を出力するので、往復しても参照が切れない）。
    """
    if name is not None:
        check_control_name(name)
    return {'kind': 'lbl', 'caption': caption, 'width': width,
            'height': STYLE['lbl_h'], 'bold': bold, 'name': name,
            'tab_index': tab_index}


def heading(caption, name=None):
    """セクション見出し（Bold・大きめ）。1行を占有する"""
    if name is not None:
        check_control_name(name)
    e = {'kind': 'lbl', 'caption': caption, 'width': _text_w(caption, STYLE['heading_pt']),
         'height': STYLE['lbl_h'] + 2, 'bold': True, 'font': STYLE['heading_pt'],
         'name': name}
    return {'kind': 'row', 'elems': [e]}


def txt(name, width=None, height=None, multiline=False, value="", required=False,
        tab_index=None):
    """テキストボックス。width 省略で入力列いっぱいにストレッチ。

    required=True で直前ラベルに「＊」を付け、vba_stub の実行ボタンに
    空チェックの雛形が入る（チェック内容の判断は書く側）。
    """
    check_control_name(name)
    return {'kind': 'txt', 'name': name, 'width': width,
            'height': height or STYLE['ctrl_h'],
            'multiline': multiline, 'value': value, 'required': required,
            'tab_index': tab_index}


def refedit(name, width=None):
    """セル範囲の選択欄（TextBox + 「選択」ボタンの複合部品）。

    本物の RefEdit コントロールは COM からの挿入が Forms エンジンの信頼設定で
    ブロックされる（実測: 「サブジェクトは信頼されていません」）ため、
    TextBox + Application.InputBox(Type:=8) 方式で実装する。
    こちらはモードレス表示でも安定して動く。vba_stub がボタンの
    範囲選択ハンドラを自動生成する。
    """
    check_control_name(name)
    t = {'kind': 'txt', 'name': name, 'width': width,
         'height': STYLE['ctrl_h'], 'multiline': False, 'value': "",
         'required': False}
    b = {'kind': 'btn', 'name': name + "Pick", 'caption': "選択",
         'width': 44, 'height': STYLE['ctrl_h'],
         'default': False, 'cancel': False, 'bold': False, 'accel': None,
         'pick_for': name}
    return [t, b]


def img(name, width, height, picture=None):
    """画像。picture にファイルパスを渡すと vba_stub の Initialize で
    LoadPicture される（デザイン時埋め込みではなく実行時読込）。"""
    check_control_name(name)
    return {'kind': 'img', 'name': name, 'width': width, 'height': height,
            'picture': picture}


def spin_txt(name, value="0", width=48, min_=0, max_=100):
    """▲▼付き数値入力（TextBox + SpinButton の複合部品）。

    行内に txt と spin を並べて返し、vba_stub が連動イベント
    （Spin変更→txt反映、txt変更→Spin追随）を生成する。
    """
    check_control_name(name)
    t = {'kind': 'txt', 'name': name, 'width': width, 'height': STYLE['ctrl_h'],
         'multiline': False, 'value': str(value), 'required': False,
         'spin': {'min': min_, 'max': max_}}
    s = {'kind': 'spin', 'name': name + "Spin", 'width': 14,
         'height': STYLE['ctrl_h'], 'for': name, 'min': min_, 'max': max_}
    return [t, s]


def combo(name, width=None, items=None, rowsource=None, tab_index=None):
    """コンボボックス。items は vba_stub の Initialize で AddItem、
    rowsource="シート名!A1:A10" はシート範囲への直結（デザイン時プロパティ）"""
    check_control_name(name)
    return {'kind': 'combo', 'name': name, 'width': width,
            'height': STYLE['ctrl_h'], 'items': items or [],
            'rowsource': rowsource, 'tab_index': tab_index}


def lst(name, width=None, height=None, rows_visible=6, items=None, rowsource=None,
        tab_index=None):
    """リストボックス。height 省略時は rows_visible 行ぶん。rowsource はシート範囲直結

    一覧系フォームは ListBox が TabIndex 0（開いた瞬間のフォーカスが一覧＝PageDown
    が効く）のが正。配置順の自動採番でそれが崩れる場合は tab_index=0 を明示する。
    """
    check_control_name(name)
    return {'kind': 'lst', 'name': name, 'width': width,
            'height': height or (rows_visible * 12 + 6), 'items': items or [],
            'rowsource': rowsource, 'tab_index': tab_index}


def chk(name, caption, width=None, tab_index=None):
    """チェックボックス"""
    check_control_name(name)
    return {'kind': 'chk', 'name': name, 'caption': caption,
            'width': width or (_text_w(caption) + 16), 'height': STYLE['lbl_h'],
            'tab_index': tab_index}


def opt(name, caption, width=None, group=None, tab_index=None):
    """オプションボタン。group で排他グループを指定"""
    check_control_name(name)
    return {'kind': 'opt', 'name': name, 'caption': caption,
            'width': width or (_text_w(caption) + 16), 'height': STYLE['lbl_h'],
            'group': group, 'tab_index': tab_index}


def opt_group(*pairs, group=None):
    """排他ボタン群を横並びで返す。pairs は (名前, 表示) のタプル列。

    例: row(lbl("優先度"), *opt_group(("optHigh","高"), ("optLow","低")))
    """
    group = group or (pairs[0][0] + "_grp" if pairs else "grp")
    return [opt(n, c, group=group) for n, c in pairs]


def btn(name, caption, width=None, height=None, default=False, cancel_btn=False,
        bold=False, accel=None, tab_index=None):
    """ボタン。button_bar 内では全ボタンが同サイズに揃えられる。accel=アクセラレータ文字"""
    check_control_name(name)
    return {'kind': 'btn', 'name': name, 'caption': caption,
            'width': width or max(STYLE['btn_w'], _text_w(caption) + 16),
            'height': height or STYLE['btn_h'],
            'default': default, 'cancel': cancel_btn, 'bold': bold, 'accel': accel,
            'tab_index': tab_index}


def ok(name="btnOK", caption="OK", **kw):
    """実行ボタン（Enter=Default）"""
    kw.setdefault('default', True)
    return btn(name, caption, **kw)


def cancel(name="btnCancel", caption="キャンセル", **kw):
    """キャンセルボタン（Esc=Cancel）。button_bar の右端に置くのが規約"""
    kw.setdefault('cancel_btn', True)
    return btn(name, caption, **kw)


def row(*elems):
    """1行。先頭が lbl ならラベル列として整列する。opt_group の返すリストは展開する"""
    flat = []
    for e in elems:
        if isinstance(e, list):
            flat.extend(e)
        else:
            flat.append(e)
    # 空の行を通すと配置計算の max() が空列で ValueError になり、原因が分からない
    # エラーとして噴き出す。宣言の受理時点で理由つきで止める（縦の空きは spacer()）
    if not flat:
        raise ValueError("row() には要素を1つ以上入れてください"
                         "（縦の空きが欲しい場合は spacer() を使う）")
    return {'kind': 'row', 'elems': flat}


def button_bar(*btns):
    """ボタンバー（右下寄せ・同サイズ揃え）。並び順は「実行系 → キャンセルが右端」"""
    if not btns:
        raise ValueError("button_bar() にはボタンを1つ以上入れてください")
    return {'kind': 'bar', 'elems': list(btns)}


def spacer(height=None):
    """縦の空き（セクション区切り）。既定は gap_y の2倍"""
    return {'kind': 'spacer', 'height': height or STYLE['gap_y'] * 2}


def frame(caption, *rows_, name=None, tab_index=None):
    """枠付きグループ。中に row/spacer/bar を入れられる（frame の入れ子は不可）

    name 省略時は caption から決定的に採番する。hash() はプロセス毎に乱数化
    されるため使わない（再 build のたびに名前が変わり VBA からの参照が切れる）。
    同一 caption の frame を2つ置くと名前が衝突して build が失敗するので、
    その場合は name を明示すること。
    """
    for r in rows_:
        if r['kind'] in ('frame', 'multipage'):
            raise ValueError("frame の中に frame / multipage は入れられません")
    auto = f"fra{zlib.crc32(str(caption).encode('utf-8')) % 10000:04d}"
    return {'kind': 'frame', 'caption': caption, 'rows': list(rows_),
            'name': check_control_name(name or auto), 'tab_index': tab_index}


def page(caption, *rows_):
    """multipage の1タブ。中に row/spacer/bar/frame を入れられる"""
    for r in rows_:
        if r['kind'] == 'multipage':
            raise ValueError("page の中に multipage は入れられません（入れ子は1段まで）")
    return {'kind': 'page', 'caption': caption, 'rows': list(rows_)}


def multipage(name, *pages_, tab_index=None):
    """タブ付きコンテナ。page(...) を並べる。

    例: multipage("mpMain",
            page("基本", row(lbl("名前"), txt("txtName"))),
            page("詳細", row(lbl("メモ"), txt("txtMemo", multiline=True))))
    """
    if not pages_ or any(p['kind'] != 'page' for p in pages_):
        raise ValueError("multipage には page(...) を1つ以上入れてください")
    check_control_name(name)
    return {'kind': 'multipage', 'name': name, 'pages': list(pages_),
            'tab_index': tab_index}


# ================================================================
# レイアウト計算（純粋計算・COM なし）
# ================================================================

REQUIRED_MARK = ' ＊'   # required=True の行のラベルに付く印（幅計算と描画で共有）


def _label_col_width(rows):
    """先頭要素が lbl の行から、ラベル列の幅を決める（frame 内は独立に計算）

    required=True の入力を含む行のラベルには _layout_region が ＊ を付ける。
    印を付ける前のキャプションで幅を測ると、必須項目のラベルほど末尾が欠ける
    （＊ を含めた幅で測る）。
    """
    w = 0
    for r in rows:
        if r['kind'] == 'row' and r['elems'] and r['elems'][0]['kind'] == 'lbl' \
           and len(r['elems']) > 1:
            e = r['elems'][0]
            cap = e.get('caption') or ''
            if any(x.get('required') for x in r['elems'][1:]):
                cap += REQUIRED_MARK
            w = max(w, e['width'] or _text_w(cap))
    return w


def _row_fixed_width(r, label_w):
    """行の固定部分の幅（ストレッチ要素は最小幅で数える）"""
    gap = STYLE['gap_x']
    elems = r['elems']
    if r['kind'] == 'bar':
        # 実配置（_layout_region）は「最大ボタン幅 × 本数」で同サイズに並べる。
        # 各ボタン幅の合計で数えると実バー幅より小さくなり、バーが左にあふれる
        if not elems:
            return 0
        bw = max([e['width'] for e in elems] + [STYLE['btn_w']])
        return bw * len(elems) + gap * (len(elems) - 1)
    total = 0
    has_label = elems and elems[0]['kind'] == 'lbl' and len(elems) > 1
    start = 0
    if has_label:
        total += label_w + gap
        start = 1
    for i, e in enumerate(elems[start:]):
        if i > 0:
            total += gap
        total += e['width'] if e['width'] else STYLE['input_w']
    return total


def _natural_width(r, label_w):
    if r['kind'] in ('row', 'bar'):
        return _row_fixed_width(r, label_w)
    if r['kind'] == 'frame':
        inner_lw = _label_col_width(r['rows'])
        inner = max([_natural_width(x, inner_lw) for x in r['rows']
                     if x['kind'] in ('row', 'bar', 'frame')] or [STYLE['input_w']])
        return inner + 2 * STYLE['frame_pad']
    if r['kind'] == 'multipage':
        widest = STYLE['input_w']
        for pg in r['pages']:
            inner_lw = _label_col_width(pg['rows'])
            for x in pg['rows']:
                if x['kind'] in ('row', 'bar', 'frame'):
                    widest = max(widest, _natural_width(x, inner_lw))
        return widest + 2 * STYLE['page_pad'] + STYLE['mp_chrome']
    return 0


def _layout_region(rows, content_width, x0, y0):
    """rows を幅 content_width の領域 (x0, y0 起点) に配置する。

    戻り値: (placements, 消費した高さ)。frame の子は frame 相対座標で
    placement の elem['children'] に入る。
    """
    gap_x = STYLE['gap_x']
    gap_y = STYLE['gap_y']
    label_w = _label_col_width(rows)
    placements = []
    y = y0
    trailing_gap = 0   # 最後に足した行間（末尾が spacer なら 0）

    for r in rows:
        if r['kind'] == 'spacer':
            y += r['height']
            trailing_gap = 0
            continue

        if r['kind'] == 'frame':
            fp = STYLE['frame_pad']
            inner_w = content_width - 2 * fp
            children, inner_h = _layout_region(r['rows'], inner_w, fp, STYLE['frame_top'])
            # 高さ = キャプション帯 + 中身 + 下余白（帯を忘れると最終行がクリップされる）
            fh = STYLE['frame_top'] + inner_h + fp
            fe = dict(r)
            fe['children'] = children
            placements.append((fe, x0, y, content_width, fh))
            y += fh + gap_y
            trailing_gap = gap_y
            continue

        if r['kind'] == 'multipage':
            pp = STYLE['page_pad']
            inner_w = content_width - 2 * pp - STYLE['mp_chrome']
            pages_layout = []
            max_h = 0
            for pg in r['pages']:
                children, ih = _layout_region(pg['rows'], inner_w, pp, pp)
                pages_layout.append({'caption': pg['caption'], 'children': children})
                max_h = max(max_h, ih)
            # 高さ = タブ帯 + ページ内余白 + 最も高いページの中身 + 下余白 + 枠
            mh = STYLE['tab_h'] + pp + max_h + pp + STYLE['mp_chrome']
            me = dict(r)
            me['pages_layout'] = pages_layout
            placements.append((me, x0, y, content_width, mh))
            y += mh + gap_y
            trailing_gap = gap_y
            continue

        elems = r['elems']
        if r['kind'] == 'bar':
            bw = max([e['width'] for e in elems] + [STYLE['btn_w']])
            bh = max(e['height'] for e in elems)
            total = bw * len(elems) + gap_x * (len(elems) - 1)
            x = x0 + content_width - total
            for e in elems:
                placements.append((e, x, y, bw, bh))
                x += bw + gap_x
            y += bh + gap_y
            trailing_gap = gap_y
            continue

        # 通常の行
        row_h = max(e['height'] for e in elems)
        has_label = elems[0]['kind'] == 'lbl' and len(elems) > 1
        # 必須項目（required=True）を含む行は、ラベルに ＊ を付けて見せる
        if has_label and any(x.get('required') for x in elems[1:]):
            head = dict(elems[0])
            head['caption'] = (head.get('caption') or '') + REQUIRED_MARK
            elems = [head] + elems[1:]
        x = x0
        fixed = _row_fixed_width(r, label_w)
        stretch_elems = [e for e in elems if e['width'] is None
                         and e['kind'] in ('txt', 'combo', 'lst')]
        stretch_extra = max(0, content_width - fixed)
        per_stretch = (STYLE['input_w'] + stretch_extra // len(stretch_elems)) \
            if stretch_elems else 0

        for i, e in enumerate(elems):
            if i == 0 and has_label:
                w = label_w
            elif e['width'] is None and e['kind'] in ('txt', 'combo', 'lst'):
                w = per_stretch
            else:
                w = e['width'] or _text_w(e.get('caption', ''))
            h = e['height']
            top = y + (row_h - h) / 2 if h < row_h else y   # 行内で上下中央
            placements.append((e, x, round(top, 1), w, h))
            x += w + gap_x
        y += row_h + gap_y
        trailing_gap = gap_y

    # 末尾が spacer の場合に gap_y を引くと宣言した余白が 8pt 目減りするため、
    # 「最後に足した行間」だけを差し引く
    return placements, (y - trailing_gap - y0) if placements or y > y0 else 0


def compute_layout(rows, content_width=None):
    """宣言リストから配置計画 [(elem, left, top, width, height), ...] を計算する。

    戻り値: (placements, content_width, content_height)
    COM に触らない純粋計算なので単体テスト・プレビュー描画に使える。
    """
    pad = STYLE['pad']
    label_w = _label_col_width(rows)
    if content_width is None:
        content_width = max([_natural_width(r, label_w)
                             for r in rows
                             if r['kind'] in ('row', 'bar', 'frame', 'multipage')]
                            or [STYLE['input_w']])
    placements, used_h = _layout_region(rows, content_width, pad, pad)
    return placements, content_width, used_h + 2 * pad


# ================================================================
# ワイヤーフレームプレビュー（Excel 不要・PIL のみ）
# ================================================================

def preview_layout(rows, out_path=None, content_width=None, scale=2):
    """Excel を起動せずに配置図 PNG を描く（設計の高速な試行錯誤用）。

    実物の見た目ではなく「配置の検証」用。実物は build_form(png=True) で確認する。
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("Pillow が無いためプレビューは使えません（実機の build_form(png=True) を使ってください）")
        return None
    out_path = out_path or PREVIEW_FILE
    placements, cw, chh = compute_layout(rows, content_width)
    W = int((cw + 2 * STYLE['pad']) * scale)
    H = int(chh * scale)
    img = Image.new("RGB", (W, H), (240, 240, 240))
    drw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("meiryo.ttc", 6 * scale)
        font_s = ImageFont.truetype("meiryo.ttc", 5 * scale)
    except Exception:
        font = font_s = ImageFont.load_default()

    def draw_one(e, l, t, w, h, dx=0, dy=0):
        x1, y1 = (l + dx) * scale, (t + dy) * scale
        x2, y2 = x1 + w * scale, y1 + h * scale
        kind = e['kind']
        if kind == 'frame':
            drw.rectangle([x1, y1, x2, y2], outline=(120, 120, 120), width=1)
            drw.text((x1 + 4, y1 - 2), e['caption'], fill=(60, 60, 60), font=font_s)
            for ce, cl, ct2, cw2, ch2 in e['children']:
                draw_one(ce, cl, ct2, cw2, ch2, dx=l + dx, dy=t + dy)
            return
        if kind == 'multipage':
            tab_h = STYLE['tab_h'] * scale
            drw.rectangle([x1, y1 + tab_h, x2, y2], outline=(120, 120, 120), width=1)
            # タブ帯（先頭がアクティブ。中身の描画は1ページ目のみ＝プレビューの割り切り）
            tx = x1
            for i, pg in enumerate(e['pages_layout']):
                tw = (_text_w(pg['caption'], 9) + 10) * scale
                fill = (255, 255, 255) if i == 0 else (215, 215, 215)
                drw.rectangle([tx, y1, tx + tw, y1 + tab_h], fill=fill,
                              outline=(120, 120, 120))
                drw.text((tx + 5 * scale, y1 + 3 * scale), pg['caption'],
                         fill=(0, 0, 0), font=font_s)
                tx += tw
            for ce, cl, ct2, cw2, ch2 in e['pages_layout'][0]['children']:
                draw_one(ce, cl, ct2, cw2, ch2,
                         dx=l + dx + 2, dy=t + dy + STYLE['tab_h'])
            if len(e['pages_layout']) > 1:
                drw.text((x1 + 4, y2 - 12 * scale / 2 - 2),
                         f"(他タブの中身は build 後に確認: "
                         f"{', '.join(p['caption'] for p in e['pages_layout'][1:])})",
                         fill=(150, 150, 150), font=font_s)
            return
        fill = {'txt': (255, 255, 255), 'combo': (255, 255, 255),
                'lst': (255, 255, 255), 'btn': (225, 225, 225),
                'refedit': (255, 255, 250), 'img': (210, 210, 210),
                'spin': (225, 225, 225),
                'chk': None, 'opt': None, 'lbl': None}.get(kind)
        if fill:
            drw.rectangle([x1, y1, x2, y2], fill=fill, outline=(100, 100, 100))
        cap = e.get('caption') or e.get('name') or ''
        if kind == 'lbl':
            drw.text((x1, y1 + 2), cap, fill=(0, 0, 0), font=font)
        elif kind == 'btn':
            bbox = drw.textbbox((0, 0), cap, font=font)
            drw.text(((x1 + x2 - (bbox[2] - bbox[0])) / 2, y1 + 3), cap,
                     fill=(0, 0, 0), font=font)
            if e.get('default'):
                drw.rectangle([x1 - 2, y1 - 2, x2 + 2, y2 + 2],
                              outline=(0, 90, 200), width=1)
        elif kind in ('chk', 'opt'):
            box = [x1, y1 + 3, x1 + 9 * scale / 2, y1 + 3 + 9 * scale / 2]
            if kind == 'chk':
                drw.rectangle(box, outline=(80, 80, 80))
            else:
                drw.ellipse(box, outline=(80, 80, 80))
            drw.text((x1 + 7 * scale, y1 + 2), cap, fill=(0, 0, 0), font=font)
        else:
            drw.text((x1 + 2, y1 + 2), e.get('name', ''), fill=(150, 150, 150), font=font_s)

    for e, l, t, w, h in placements:
        draw_one(e, l, t, w, h)
    img.save(out_path)
    print(f"レイアウトプレビュー: {out_path}  (コンテンツ {cw:g}x{chh:g}pt / Excel不使用)")
    return out_path


# ================================================================
# VBA イベント雛形の生成（機械的な骨組みのみ・中身の判断はしない）
# ================================================================

def _iter_rows(rows):
    """宣言ツリーから row/bar の要素リストを列挙（frame / multipage の中も辿る）"""
    for r in rows:
        if r['kind'] in ('row', 'bar'):
            yield r['elems']
        elif r['kind'] == 'frame':
            yield from _iter_rows(r['rows'])
        elif r['kind'] == 'multipage':
            for pg in r['pages']:
                yield from _iter_rows(pg['rows'])


def generate_vba_stub(rows, out_path=None):
    """コントロール構成からイベントプロシージャの骨組みを機械生成して UTF-8 で保存する。

    - UserForm_Initialize: combo/lst の items を AddItem、spin_txt の初期同期、
      img の LoadPicture、最初の入力に SetFocus
    - required=True の入力: 実行（Default）ボタンの Click 冒頭に空チェックを生成
    - ok/通常ボタン: Click 雛形（中身は ' TODO）、cancel ボタン: Unload Me
    - spin_txt: SpinButton_Change → TextBox 反映のイベントを生成
    """
    all_rows = list(_iter_rows(rows))
    elems = [e for row_elems in all_rows for e in row_elems]
    lines = []

    # 必須項目: (コントロール名, 行頭ラベルの表示名)
    required = []
    for row_elems in all_rows:
        label = None
        if row_elems and row_elems[0]['kind'] == 'lbl':
            label = row_elems[0].get('caption')
        for e in row_elems:
            if e.get('required'):
                required.append((e['name'], label or e['name']))

    init = []
    for e in elems:
        if e['kind'] in ('combo', 'lst') and e.get('items'):
            for it in e['items']:
                # VBA の文字列リテラル中の " は "" でエスケープする
                # （生のまま埋めると項目に " が含まれた時点で生成コードが壊れる）
                esc = str(it).replace('"', '""')
                init.append(f'    {e["name"]}.AddItem "{esc}"')
    for e in elems:
        if e['kind'] == 'spin':
            init.append(f'    {e["name"]}.Value = Val({e["for"]}.Value)')
        if e['kind'] == 'img' and e.get('picture'):
            init.append(f'    {e["name"]}.Picture = LoadPicture("{e["picture"]}")')
    first_input = next((e for e in elems
                        if e['kind'] in ('txt', 'combo', 'refedit')), None)
    if first_input:
        init.append(f'    {first_input["name"]}.SetFocus')
    if init:
        lines.append("Private Sub UserForm_Initialize()")
        lines.extend(init)
        lines.append("End Sub")
        lines.append("")

    for e in elems:
        if e['kind'] == 'spin':
            lines.append(f"Private Sub {e['name']}_Change()")
            lines.append(f"    {e['for']}.Value = {e['name']}.Value")
            lines.append("End Sub")
            lines.append("")

    for e in elems:
        if e['kind'] != 'btn':
            continue
        lines.append(f"Private Sub {e['name']}_Click()")
        if e.get('pick_for'):
            # refedit 複合部品: シート上で範囲を選ばせて TextBox に反映
            lines.append("    Dim r As Range")
            lines.append("    On Error Resume Next")
            lines.append('    Set r = Application.InputBox("対象範囲をシート上で選択してください", Type:=8)')
            lines.append("    On Error GoTo 0")
            lines.append(f"    If Not r Is Nothing Then {e['pick_for']}.Value = r.Address(False, False)")
        elif e.get('cancel'):
            lines.append("    Unload Me")
        else:
            if e.get('default') and required:
                for name, label in required:
                    lines.append(f'    If Trim({name}.Value) = "" Then')
                    lines.append(f'        MsgBox "{label} を入力してください", vbExclamation')
                    lines.append(f'        {name}.SetFocus')
                    lines.append("        Exit Sub")
                    lines.append("    End If")
            lines.append(f"    ' TODO: {e['caption']} の処理")
        lines.append("End Sub")
        lines.append("")

    code = "\n".join(lines).rstrip() + "\n"
    out_path = out_path or STUB_FILE
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(code)
    print(f"イベント雛形を生成: {out_path}  ({len([l for l in lines if l.startswith('Private')])}プロシージャ)")
    return out_path


# ================================================================
# ビルド（COM。配置計画をフォームに反映）
# ================================================================

_PROGID = {
    'lbl':   "Forms.Label.1",
    'txt':   "Forms.TextBox.1",
    'combo': "Forms.ComboBox.1",
    'lst':   "Forms.ListBox.1",
    'chk':   "Forms.CheckBox.1",
    'opt':   "Forms.OptionButton.1",
    'btn':   "Forms.CommandButton.1",
    'frame': "Forms.Frame.1",
    'multipage': "Forms.MultiPage.1",
    'img':   "Forms.Image.1",
    'spin':  "Forms.SpinButton.1",
}


def _backup_existing_form(fb, form_name):
    """既存フォームを作り直す前に .frm/.frx を backups へ退避する（安全網）"""
    for comp in fb.vbproject.VBComponents:
        if comp.Name.lower() == form_name.lower() and comp.Type == 3:
            try:
                os.makedirs(BACKUP_DIR, exist_ok=True)
                stamp = time.strftime("%Y%m%d_%H%M%S")
                base = os.path.splitext(os.path.basename(fb.workbook.FullName))[0]
                path = os.path.join(BACKUP_DIR, f"{base}_{form_name}_{stamp}.frm")
                comp.Export(path)
                print(f"フォーム退避: backups/{os.path.basename(path)} (+.frx)")
            except Exception as e:
                print(f"⚠ フォーム退避に失敗（続行）: {e}")
            return


def _place_control(container, e, left, top, w, h, tab_state):
    """1コントロールを container（フォーム or Frame）に配置する"""
    kind = e['kind']
    # ラベルは name 省略可（lbl1, lbl2... と機械採番）。名前を明示したラベルは
    # VBA から名指しされている可能性があるので、その名前をそのまま使う
    if kind == 'lbl' and not e.get('name'):
        tab_state['lbl_seq'] += 1
        name = f"lbl{tab_state['lbl_seq']}"
    else:
        name = e.get('name') or f"ctl{tab_state['tab']}"
    check_control_name(name)
    ct = container.Controls.Add(_PROGID[kind], name)
    ct.Left = left; ct.Top = top; ct.Width = w; ct.Height = h
    try:
        ct.Font.Size = e.get('font') or STYLE['font']
    except Exception:
        pass
    if 'caption' in e and kind not in ('txt',):
        try:
            ct.Caption = e['caption']
        except Exception:
            pass
    if kind == 'txt':
        if e.get('value'):
            ct.Value = e['value']
        if e.get('multiline'):
            ct.MultiLine = True
            ct.ScrollBars = 2
    if kind == 'opt' and e.get('group'):
        try:
            ct.GroupName = e['group']
        except Exception:
            pass
    if kind == 'spin':
        try:
            ct.Min = e.get('min', 0)
            ct.Max = e.get('max', 100)
        except Exception:
            pass
    if kind in ('combo', 'lst') and e.get('rowsource'):
        try:
            ct.RowSource = e['rowsource']
        except Exception as ex:
            print(f"⚠ {name} の RowSource 設定に失敗（続行）: {ex}")
    if kind == 'btn':
        if e.get('default'):
            ct.Default = True
        if e.get('cancel'):
            ct.Cancel = True
        if e.get('bold'):
            ct.Font.Bold = True
        if e.get('accel'):
            try:
                ct.Accelerator = e['accel']
            except Exception:
                pass
    if kind == 'lbl' and e.get('bold'):
        ct.Font.Bold = True
    # tab_index を明示した要素はその値を使う（省略時のみ配置順で自動採番）。
    # ListBox 先頭フォーカスのような「配置順では表せない意図」を、
    # inspect --to-layout → build の往復で失わないための逃げ道
    ti = e.get('tab_index')
    try:
        ct.TabIndex = tab_state['tab'] if ti is None else ti
    except Exception:
        pass
    tab_state['tab'] += 1
    return ct


def build_form(form_name, caption, rows, width=None, vba_file=None,
               vba_stub=False, png=False, save=True, launcher=None):
    """宣言的レイアウトからフォームを構築する。

    width: コンテンツ幅の明示（省略時は自動）。vba_file: UTF-8 の .vba を注入。
    vba_stub=True でイベント雛形を機械生成して注入（vba_file 指定が優先）。
    png=True で構築後に実表示PNG（_last_form.png）を撮って目視検証につなげる。
    launcher="モジュール名" で `Sub <フォーム名>を開く()` を標準モジュールに自動追加
    （メニュー方式のブックならそのままメニューに載る。既にあればスキップ）。
    既存フォームは backups へ退避してから作り直す。
    環境変数 FORM_LAYOUT_PREVIEW=1 のときは構築せず配置図PNGだけ生成する
    （CLI の preview モードが使う。Excel 不要）。
    """
    if os.environ.get('FORM_LAYOUT_PREVIEW') == '1':
        print(f"[preview モード] {form_name}（{caption}）→ 配置図のみ生成（Excel不使用）")
        cw = width - 2 * STYLE['pad'] if width is not None else None
        return preview_layout(rows, content_width=cw)

    content_w = None
    if width is not None:
        content_w = width - 2 * STYLE['pad']
    placements, content_w, content_h = compute_layout(rows, content_w)

    form_w = content_w + 2 * STYLE['pad'] + STYLE['chrome_w']
    form_h = content_h + STYLE['chrome_h']

    if vba_stub and not vba_file:
        vba_file = generate_vba_stub(rows)

    with FormBuilder.connect() as fb:
        _backup_existing_form(fb, form_name)
        frm = fb.get_or_create(form_name, caption=caption,
                               width=form_w, height=form_h)
        f = fb.clear_controls(frm)
        tab_state = {'tab': 0, 'lbl_seq': 0}
        n = 0
        for e, left, top, w, h in placements:
            if e['kind'] == 'frame':
                fra = _place_control(f, e, left, top, w, h, tab_state)
                n += 1
                for ce, cl, ct2, cw2, ch2 in e['children']:
                    _place_control(fra, ce, cl, ct2, cw2, ch2, tab_state)
                    n += 1
            elif e['kind'] == 'multipage':
                mp = _place_control(f, e, left, top, w, h, tab_state)
                n += 1
                # 既定で2ページ作られるので、宣言の数に合わせて増減する
                want = len(e['pages_layout'])
                while mp.Pages.Count < want:
                    mp.Pages.Add()
                while mp.Pages.Count > want:
                    mp.Pages.Remove(mp.Pages.Count - 1)
                for pi, pg in enumerate(e['pages_layout']):
                    pobj = mp.Pages(pi)          # MSForms の Pages は 0 始まり
                    pobj.Caption = pg['caption']
                    for ce, cl, ct2, cw2, ch2 in pg['children']:
                        ctl = _place_control(pobj, ce, cl, ct2, cw2, ch2, tab_state)
                        n += 1
                        if ce['kind'] == 'frame':
                            # page 内の frame はトップレベルの frame と同じく
                            # 子を再帰配置する（忘れると枠だけ置かれ中身が消える）
                            for fe, fl, ft, fw, fh in ce['children']:
                                _place_control(ctl, fe, fl, ft, fw, fh, tab_state)
                                n += 1
                mp.Value = 0                     # 先頭タブを表示状態に
            else:
                _place_control(f, e, left, top, w, h, tab_state)
                n += 1

        print(f"レイアウト構築: {form_name}  コントロール {n}個  "
              f"フォーム {form_w:g}x{form_h:g}")
        if vba_file:
            fb.inject_vba(frm, vba_file)

        if launcher:
            # 起動マクロ: Sub <フォーム名>を開く() を指定モジュールへ（無ければ追加しない
            # で候補を出す。フォームは作れてもメニューから呼べない、の穴を塞ぐ）
            mod = None
            names = []
            for c in fb.vbproject.VBComponents:
                if c.Type == 1:
                    names.append(c.Name)
                    if c.Name.lower() == str(launcher).lower():
                        mod = c
            if mod is None:
                print(f"⚠ 起動マクロ先モジュール '{launcher}' が見つかりません。"
                      f"標準モジュール: {', '.join(names) or '(なし)'}")
            else:
                proc = f"{form_name}を開く"
                # 識別子ガード（注入点側の多層防御）: フォーム名由来の Sub 名を注入前に検査
                import vba_manager
                bad = vba_manager.check_vba_identifier(proc)
                if bad:
                    print(f"⚠ 起動マクロを追加しません（VBA の識別子規則違反）: {bad}")
                else:
                    cm = mod.CodeModule
                    try:
                        cm.ProcStartLine(proc, 0)
                        print(f"起動マクロは既存: [{mod.Name}] {proc}（変更なし）")
                    except Exception:
                        # 改行は CRLF に揃える（改行二重化ガード／他の注入点と統一。
                        # LF のまま InsertLines に渡していた唯一の穴だった）
                        body = "\r\n".join(
                            ["", f"Sub {proc}()", f"    {form_name}.Show", "End Sub"])
                        cm.InsertLines(cm.CountOfLines + 1, body)
                        print(f"起動マクロ追加: [{mod.Name}] Sub {proc}()")

        if save:
            fb.save()

        if png:
            # 構築したら必ず目で確認する、のループを1関数で完結させる
            try:
                from form_inspect import render_form_png
                out = os.path.join(SCRIPT_DIR, "_last_form.png")
                render_form_png(fb._xl, fb.workbook, form_name, out)
            except Exception as ex:
                print(f"⚠ PNG撮影に失敗（フォーム自体は構築済み）: {ex}")
    return form_w, form_h


# ================================================================
# CLI（宣言ファイルを直接 preview / build する）
# ================================================================

def main():
    """py form_layout.py <preview|build> <宣言スクリプト.py>

    宣言スクリプトは build_form(...) を呼ぶ普通の Python ファイル（UTF-8）。
    preview は Excel を起動せず配置図PNGだけ生成する（FORM_LAYOUT_PREVIEW=1 で
    build_form が preview_layout に切り替わる仕組みなので、スクリプト側の
    書き分けは不要＝同じファイルが preview でも build でも動く）。
    """
    import argparse
    import runpy
    ap = argparse.ArgumentParser(description="宣言的フォームレイアウトの実行")
    ap.add_argument("mode", choices=["preview", "build"],
                    help="preview=配置図PNGのみ（Excel不要） / build=実構築")
    ap.add_argument("spec", help="build_form(...) を呼ぶ宣言スクリプト（UTF-8 .py）")
    args = ap.parse_args()
    if not os.path.exists(args.spec):
        print(f"エラー: 宣言スクリプトが見つかりません: {args.spec}")
        sys.exit(1)
    if args.mode == "preview":
        os.environ['FORM_LAYOUT_PREVIEW'] = '1'
    else:
        os.environ.pop('FORM_LAYOUT_PREVIEW', None)
    runpy.run_path(args.spec, run_name="__main__")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    main()
