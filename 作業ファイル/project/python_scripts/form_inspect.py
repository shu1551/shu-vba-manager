# -*- coding: utf-8 -*-
"""
form_inspect.py — UserForm を「1回のCOM接続」でまとめて把握する。

狙い：vba_manager の export-module / get を別々に何度も叩くと、CLIは毎回
COM接続し直す（一番重い）。フォーム修正の下調べでこれを繰り返すと時間を食う。
このスクリプトは1プロセス・1接続で「フォーム情報 ＋ コントロール配置 ＋ VBAコード
(全体 or 指定プロシージャ)」をまとめて出力する。

接続は gencache 非経由の late-binding（dynamic.Dispatch）。gen_py キャッシュ破損の
影響を受けない。対象は常に「今アクティブに開いているブック」で、特定ブック・パスに
依存しない（他のExcelでもそのまま使える自立ユニット）。

使い方:
  py form_inspect.py <フォーム名>                       # フォーム情報+コントロール一覧+全コード
  py form_inspect.py <フォーム名> WriteDate btnX_Click  # 指定プロシージャだけ
  py form_inspect.py --list                              # 開いているブックの全フォーム名

オプション:
  --font        コントロール一覧に Font.Size 列を追加
  --json        フォーム情報＋コントロール一覧＋(プロシージャ指定時は)コードを
                機械可読 JSON で stdout へ出力（人間向け出力は stderr へ）
  --png [パス]  フォームを実表示して PNG 撮影（既定: スクリプトと同じ場所の _last_form.png）
                Pillow が無い環境では BMP で保存
  --names       --png にコントロールの枠と名前を描き込む（AIが名指しで直せる画像になる）
  --lint        重なり・はみ出し・ボタン/フォント不揃い・タブ順のずれを機械検査

例:
  py form_inspect.py F_Calendar
  py form_inspect.py F_Calendar WriteDate btnNyuryoku_Click
  py form_inspect.py F_Calendar --font --png
"""
import argparse
import json
import os
import sys
import pythoncom
import pywintypes
import win32com.client.dynamic

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PNG = os.path.join(SCRIPT_DIR, "_last_form.png")


def get_active_excel():
    """起動中の Excel に late-binding で接続（gencache 非経由・キャッシュ破損に強い）。"""
    pythoncom.CoInitialize()
    clsid = pywintypes.IID("Excel.Application")
    unk = pythoncom.GetActiveObject(clsid)
    return win32com.client.dynamic.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))


# COM の GetTypeInfo は環境によりインターフェース名（IMdcText 等）を返すため、
# MSForms の一般名に正規化する（lint / --to-layout の型判定はこの名前で行う）
_TYPE_NORMALIZE = {
    "ILabelControl":    "Label",
    "IMdcText":         "TextBox",
    "IMdcCombo":        "ComboBox",
    "IMdcList":         "ListBox",
    "IMdcCheckBox":     "CheckBox",
    "IMdcOptionButton": "OptionButton",
    "ICommandButton":   "CommandButton",
    "IOptionFrame":     "Frame",
    "IMultiPage":       "MultiPage",
    "IPage":            "Page",
    "IImage":           "Image",
    "ISpinbutton":      "SpinButton",
    "IScrollbar":       "ScrollBar",
    "IMdcToggleButton": "ToggleButton",
}


def _normalize_type(t):
    if not t:
        return t
    if t in _TYPE_NORMALIZE:
        return _TYPE_NORMALIZE[t]
    # 既に一般名（Label/TextBox…）ならそのまま。辞書に無い I〜名は、辞書のキーと
    # 部分一致すればその一般名に寄せ、どれとも一致しなければ受け取った名前を
    # そのまま返す（接頭 I は外さない。判定側は startswith で見るので、
    # 素性の分からない型を勝手に一般名らしく見せない方が安全）
    if t.startswith("I") and len(t) > 1 and t[1].isupper():
        for key, val in _TYPE_NORMALIZE.items():
            if key.lower() in t.lower():
                return val
    return t


def find_component(wb, form_name):
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == form_name.lower():
            return c
    return None


def _round2(v):
    return None if v is None else round(float(v), 2)


def _fmt_num(v):
    if v is None:
        return "?"
    return f"{float(v):g}"


def get_form_info(comp):
    """フォーム自身の情報 (Caption/Width/Height/InsideWidth/InsideHeight) を dict で返す。"""
    info = {}
    for prop in ("Caption", "Width", "Height"):
        try:
            info[prop.lower()] = comp.Properties(prop).Value
        except Exception:
            info[prop.lower()] = None
    try:
        d = comp.Designer
        info["inside_width"] = d.InsideWidth
        info["inside_height"] = d.InsideHeight
    except Exception:
        info["inside_width"] = None
        info["inside_height"] = None
    return info


def print_form_info(info):
    print(f'フォーム情報: Caption="{info["caption"]}"  '
          f'W={_fmt_num(info["width"])} H={_fmt_num(info["height"])}  '
          f'InsideW={_fmt_num(info["inside_width"])} '
          f'InsideH={_fmt_num(info["inside_height"])}')


def collect_controls(comp):
    """コントロール情報を dict のリストで返す（表示 / --json 共通のデータ）。"""
    try:
        ctrls = comp.Designer.Controls
    except Exception as e:
        raise RuntimeError(f"コントロール取得不可: {e}")
    items = []
    for ct in ctrls:
        d = {}
        try:
            d["name"] = ct.Name
            d["left"] = _round2(ct.Left)
            d["top"] = _round2(ct.Top)
            d["width"] = _round2(ct.Width)
            d["height"] = _round2(ct.Height)
            d["caption"] = None
            try:
                c = ct.Caption
                if c:
                    d["caption"] = c
            except Exception:
                pass
            d["font_size"] = None
            try:
                d["font_size"] = _round2(ct.Font.Size)
            except Exception:
                pass
            # 座標は親コンテナ相対（Frame 内は Frame 基準）。lint/overlay で使う
            d["parent"] = None
            try:
                d["parent"] = ct.Parent.Name
            except Exception:
                pass
            d["tab_index"] = None
            try:
                d["tab_index"] = int(ct.TabIndex)
            except Exception:
                pass
            d["type"] = None
            try:
                d["type"] = _normalize_type(
                    ct._oleobj_.GetTypeInfo().GetDocumentation(-1)[0])
            except Exception:
                pass
            # 種別ごとの追加属性（lint / --to-layout 用）
            t = d["type"] or ""
            if t.startswith("CommandButton"):
                for prop in ("Default", "Cancel"):
                    try:
                        d[prop.lower()] = bool(getattr(ct, prop))
                    except Exception:
                        d[prop.lower()] = None
                try:
                    d["accelerator"] = ct.Accelerator or None
                except Exception:
                    d["accelerator"] = None
            if t.startswith("TextBox"):
                try:
                    d["multiline"] = bool(ct.MultiLine)
                except Exception:
                    d["multiline"] = None
            if t.startswith("OptionButton"):
                try:
                    d["group"] = ct.GroupName or None
                except Exception:
                    d["group"] = None
            if t.startswith(("Label", "CheckBox", "OptionButton", "Frame", "CommandButton")):
                try:
                    d["bold"] = bool(ct.Font.Bold)
                except Exception:
                    d["bold"] = None
            if (d["type"] or "").startswith("MultiPage"):
                try:
                    d["pages"] = [{"name": p.Name, "caption": p.Caption}
                                  for p in ct.Pages]
                except Exception:
                    d["pages"] = []
        except Exception as e:
            d["error"] = str(e)
        items.append(d)
    return items


def print_controls(comp, want_font=False):
    """Designer 経由でコントロールの配置を出す（配置調整の下調べ用）。"""
    try:
        items = collect_controls(comp)
    except RuntimeError as e:
        print(f"  ({e})")
        return
    head = "Name  L,T,W,H"
    if want_font:
        head += "  Font"
    print(f"\n=== コントロール {len(items)}個  ({head}  [Caption]) ===")
    for d in items:
        if "error" in d:
            print(f"  (1件取得失敗: {d['error']})")
            continue
        geo = (f"L{int(d['left'])} T{int(d['top'])} "
               f"W{int(d['width'])} H{int(d['height'])}")
        font = f"  F{_fmt_num(d['font_size'])}" if want_font else ""
        cap = f"  [{d['caption']}]" if d["caption"] else ""
        print(f"  {d['name']:<24} {geo}{font}{cap}")


def get_proc_code(cm, proc):
    """指定プロシージャのコードを返す。見つからなければ None。"""
    try:
        start = cm.ProcStartLine(proc, 0)
        count = cm.ProcCountLines(proc, 0)
        return cm.Lines(start, count).rstrip("\r\n")
    except Exception:
        return None


def print_code(comp, procs):
    cm = comp.CodeModule
    if procs:
        print("\n=== 指定プロシージャ ===")
        for p in procs:
            try:
                start = cm.ProcStartLine(p, 0)
                count = cm.ProcCountLines(p, 0)
                print(f"\n--- {p}  (行 {start}〜, {count}行) ---")
                print(cm.Lines(start, count).rstrip("\r\n"))
            except Exception as e:
                print(f"\n--- {p}: 見つかりません ({e}) ---")
    else:
        total = cm.CountOfLines
        print(f"\n=== 全コード ({total}行) ===")
        if total > 0:
            print(cm.Lines(1, total).rstrip("\r\n"))


# ================================================================
# lint（デザインの機械検査。判断はせず事実を報告する）
# ================================================================

# イベント名のホワイトリスト（孤児ハンドラ判定用。これ以外の接尾辞は
# ただのプロシージャ名として扱い、誤検知を避ける）
_EVENT_NAMES = {
    'click', 'dblclick', 'change', 'enter', 'exit', 'keydown', 'keyup',
    'keypress', 'mousedown', 'mouseup', 'mousemove', 'afterupdate',
    'beforeupdate', 'beforedragover', 'beforedroporpaste', 'spinup',
    'spindown', 'scroll', 'zoom', 'layout', 'addcontrol', 'removecontrol',
}


def lint_form(form_name, info, controls, code=None):
    """重なり・はみ出し・不揃いを機械検出して所見リストを返す。

    直すかどうかの判断はしない（報告のみ）。座標は親コンテナ相対なので、
    重なり・はみ出しの判定は同じ親（フォーム直下 or 同じFrame内）同士に限る。
    """
    findings = []
    ok_items = [c for c in controls if "error" not in c and c["left"] is not None]

    def rect(c):
        return (c["left"], c["top"], c["left"] + c["width"], c["top"] + c["height"])

    # 1. 重なり（同じ親の中だけ比較）
    #    境界がぴったり接する（座標の丸め誤差含む）だけのケースを誤検知しないよう、
    #    重複幅・高さが共に OVERLAP_EPS を超える場合だけを実質的な重なりとする。
    OVERLAP_EPS = 2.0
    by_parent = {}
    for c in ok_items:
        by_parent.setdefault(c["parent"], []).append(c)
    for parent, items in by_parent.items():
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                a, b = rect(items[i]), rect(items[j])
                ox = min(a[2], b[2]) - max(a[0], b[0])
                oy = min(a[3], b[3]) - max(a[1], b[1])
                if ox > OVERLAP_EPS and oy > OVERLAP_EPS:
                    findings.append(f"重なり: {items[i]['name']} と {items[j]['name']}"
                                    + (f"（{parent} 内）" if parent and parent != form_name else ""))

    # 2. フォームからのはみ出し（フォーム直下のみ。Frame内はFrame相対のため対象外）
    iw, ih = info.get("inside_width"), info.get("inside_height")
    if iw and ih:
        for c in ok_items:
            if c["parent"] not in (None, form_name):
                continue
            if c["left"] < 0 or c["top"] < 0 \
               or c["left"] + c["width"] > iw + 0.5 or c["top"] + c["height"] > ih + 0.5:
                findings.append(f"はみ出し: {c['name']} (L{c['left']:g} T{c['top']:g} "
                                f"W{c['width']:g} H{c['height']:g} / 内寸 {iw:g}x{ih:g})")

    # 3. ボタンの高さ不揃い（同じ行のボタン同士だけ比較。
    #    行が違えば役割も違う＝入力欄に高さを合わせた「選択」ボタン等は対象外）
    #    幅の不揃いは「役割によって幅を変える」意図的な設計が普通にあり、
    #    何種類までなら妥当かの基準がないため検査しない（ボタン幅3種チェックは廃止）
    btns = [c for c in ok_items if (c.get("type") or "").startswith("CommandButton")]
    if len(btns) >= 2:
        bands = {}
        for c in btns:
            band = round((c["top"] + c["height"] / 2) / 14)
            bands.setdefault(band, []).append(c)
        for band, bs in bands.items():
            if len(bs) < 2:
                continue
            hs = sorted({c["height"] for c in bs})
            if len(hs) > 1:
                findings.append(f"同じ行のボタン高さが不揃い: {hs}  "
                                f"({', '.join(c['name'] for c in bs)})")

    # 4. フォントサイズの混在（見出し＝太字ラベルは意図的な大きさなので対象外）
    #    多数派から 1.5pt 未満しか離れていない値は無視する。長年の手作業で
    #    フォームを拡縮・複製するたびに生じる丸め誤差（11.8/10.9/14.1pt 等の
    #    半端な値）が大半で、実際の見た目では区別が付かず「混在」と呼ぶには
    #    値しない。目に見える差（1.5pt以上）だけを報告する。
    FONT_DIFF_MIN = 1.5
    sizes = {}
    for c in ok_items:
        if (c.get("type") or "").startswith("Label") and c.get("bold"):
            continue
        if c["font_size"] is not None:
            sizes.setdefault(c["font_size"], []).append(c["name"])
    if len(sizes) > 1:
        major = max(sizes, key=lambda k: len(sizes[k]))
        for sz, names in sorted(sizes.items()):
            if sz != major and abs(sz - major) >= FONT_DIFF_MIN:
                findings.append(f"フォント混在: {sz}pt = {', '.join(names)}（多数派は {major}pt）")

    # 5. TabIndex が視線順（上→下、左→右）と食い違う
    #    行内の上下中央合わせでラベルと入力の top はずれるため、比較は「上下中央」で行う
    #    ListBox は比較から外す。一覧系フォームは ListBox が TabIndex 0（開いた瞬間の
    #    フォーカスが一覧＝PageDown が効く）のが正で、画面上は下にあっても正しい。
    #    ここで「不一致」と報告すると、それを見た者が tab-order で視線順に「直し」て
    #    フォーカスを壊す（2026-07-11 実害）。
    tabbed = [c for c in ok_items if c["tab_index"] is not None
              and c["parent"] in (None, form_name)
              and not (c.get("type") or "").startswith("ListBox")]
    visual = sorted(tabbed, key=lambda c: (round((c["top"] + c["height"] / 2) / 14), c["left"]))
    actual = sorted(tabbed, key=lambda c: c["tab_index"])
    mism = [(v["name"]) for v, a in zip(visual, actual) if v["name"] != a["name"]]
    if mism:
        findings.append(f"TabIndex が視線順と不一致: {', '.join(mism[:6])}"
                        + ("…" if len(mism) > 6 else ""))

    # 5b. 右端の「揃いかけのズレ」（最大右端から 2〜16pt だけ短い入力は
    #    揃え忘れの可能性が高い。大きく短いのは意図的な短欄とみなして対象外）
    #    比較はフォーム全体ではなく「同じ列（Left が近い）」の入力同士に限る。
    #    Left がバラバラな入力を比較すると、無関係な別セクション同士を
    #    揃え忘れ扱いしてしまう（倍率回転フォームで実際に誤検知した）。
    inputs = [c for c in ok_items if c["parent"] in (None, form_name)
              and (c.get("type") or "").startswith(("TextBox", "ComboBox", "ListBox"))]
    by_left = {}
    for c in inputs:
        by_left.setdefault(round(c["left"]), []).append(c)
    for col in by_left.values():
        if len(col) < 2:
            continue
        max_r = max(round(c["left"] + c["width"]) for c in col)
        near_miss = [c for c in col
                     if 2 <= max_r - round(c["left"] + c["width"]) <= 16]
        for c in near_miss:
            findings.append(f"右端の揃い忘れの疑い: {c['name']}"
                            f"（右端 {round(c['left'] + c['width'])} / 最大 {max_r}）")

    # 5c. ラベルと隣の入力の上下中央ずれ
    labels = [c for c in ok_items if (c.get("type") or "").startswith("Label")]
    for c in labels:
        cy = c["top"] + c["height"] / 2
        near = [i for i in ok_items
                if i["parent"] == c["parent"] and i is not c
                and (i.get("type") or "").startswith(("TextBox", "ComboBox"))
                and i["top"] < cy < i["top"] + i["height"]]
        for i in near:
            iy = i["top"] + i["height"] / 2
            if abs(cy - iy) > 3:
                findings.append(f"中央ずれ: {c['name']} と {i['name']}（差 {abs(cy - iy):.0f}pt）")

    # 5d. Cancel / Accelerator（キーボード操作の作法）
    # Default(Enter) は全フォームでほぼ確実に「未設定」と出る割に、単一ボタンの
    # 終了フォームや複数択一フォームでは意味を持たないケースが大半のため、
    # 発火率が低く実際に修正につながった Cancel(Esc) だけを検査する。
    btns2 = [c for c in ok_items if (c.get("type") or "").startswith("CommandButton")]
    if btns2:
        if not any(c.get("cancel") for c in btns2):
            findings.append("Cancel(Esc) のボタンが未設定")
        accs = {}
        for c in btns2:
            a = c.get("accelerator")
            if a:
                accs.setdefault(a.upper(), []).append(c["name"])
        for a, names in accs.items():
            if len(names) > 1:
                findings.append(f"Accelerator '{a}' が重複: {', '.join(names)}")

    # 7. コード×コントロールの整合（code を渡された場合のみ）
    if code:
        import re as _re
        handlers = _re.findall(r'^\s*(?:Private\s+|Public\s+)?Sub\s+(\w+)\s*\(',
                               code, _re.MULTILINE | _re.IGNORECASE)
        names = {c["name"].lower() for c in ok_items}
        # ページ名（Page1等）はコントロール一覧に出ないが Parent として実在する
        names |= {c["parent"].lower() for c in ok_items if c.get("parent")}
        for h in handlers:
            if '_' not in h:
                continue
            ctrl, evt = h.rsplit('_', 1)
            if evt.lower() not in _EVENT_NAMES:
                continue
            if ctrl.lower() in ('userform',):
                continue
            if ctrl.lower() not in names:
                findings.append(f"孤児ハンドラ: Sub {h}（コントロール '{ctrl}' が存在しない。"
                                "リネーム/削除の置き土産の可能性）")
        handler_set = {h.lower() for h in handlers}
        for c in ok_items:
            if (c.get("type") or "").startswith("CommandButton"):
                if f"{c['name'].lower()}_click" not in handler_set:
                    findings.append(f"Click ハンドラ未実装: {c['name']}")

    # 6. （廃止）左端ラインの数チェック。ヘッダー行除外後も実際の修正には
    #    一度も繋がらず、性質の違う区画が並ぶだけの小さいフォームでは区画数が
    #    そのまま発火するだけだったため撤去した。

    return findings


def print_lint(form_name, info, controls, echo=print, code=None):
    findings = lint_form(form_name, info, controls, code=code)
    echo(f"\n=== lint ({len(findings)}件) ===")
    if not findings:
        echo("  問題なし（重なり・はみ出し・不揃い・タブ順のずれは検出されず）")
    for s in findings:
        echo(f"  [!] {s}")
    return findings


# ================================================================
# リバース変換（既存フォーム → form_layout の宣言コード）
# ================================================================

def _cluster_rows(items):
    """コントロールを上下中央の近さで行にクラスタリングする（機械的な幾何処理）"""
    items = sorted(items, key=lambda c: (c["top"] + c["height"] / 2, c["left"]))
    rows = []
    for c in items:
        cy = c["top"] + c["height"] / 2
        placed = False
        for r in rows:
            rc = sum(x["top"] + x["height"] / 2 for x in r) / len(r)
            if abs(cy - rc) <= 8:
                r.append(c)
                placed = True
                break
        if not placed:
            rows.append([c])
    for r in rows:
        r.sort(key=lambda c: c["left"])
    rows.sort(key=lambda r: min(x["top"] for x in r))
    return rows


def _tab_arg(c, tab_map):
    """tab_map に載っているコントロールだけ tab_index=N を明示する引数文字列"""
    if tab_map and c["name"] in tab_map:
        return f", tab_index={tab_map[c['name']]}"
    return ""


# form_layout の STYLE['font'] と同じ既定。これと違うサイズだけ font= を出す
_DEFAULT_FONT_PT = 12
# form_layout の STYLE['heading_pt'] と同じ既定（heading() の逆変換用）
_HEADING_FONT_PT = 13


def _font_arg(c):
    """既定と違うフォントサイズなら font= を出す。

    出さないと、往復（--to-layout → build_form）で全コントロールが 12pt に
    揃えられてしまう（9pt で作られた既存フォームは見た目が別物になる）。
    form_layout の _place_control は e.get('font') を読む実装になっているので、
    要素式に font= を書けばそのまま効く。
    """
    fs = c.get("font_size")
    if not fs:
        return ""
    try:
        if abs(float(fs) - _DEFAULT_FONT_PT) < 0.5:
            return ""
    except Exception:
        return ""
    return f", font={float(fs):g}"


def _emit_elem(c, max_right, tab_map=None):
    """1コントロールを form_layout の要素式に変換する。未対応種別は None"""
    t = c.get("type") or ""
    name = c["name"]
    cap = (c.get("caption") or "").replace('"', '\\"')
    stretch = abs((c["left"] + c["width"]) - max_right) <= 1 if max_right else False
    w_arg = "" if stretch else f", width={c['width']:g}"
    ti = _tab_arg(c, tab_map)
    fo = _font_arg(c)

    if t.startswith("Label"):
        b = ", bold=True" if c.get("bold") else ""
        # ラベルも実名を出す。名前を捨てると受け側が lbl1, lbl2... と機械採番し、
        # VBA が lblStatus.Caption = ... と名指ししていた参照が黙って切れる
        return f'lbl("{cap}", name="{name}"{b}{fo}{ti})'
    if t.startswith("TextBox"):
        h = f", height={c['height']:g}" if abs(c["height"] - 22) > 1 else ""
        ml = ", multiline=True" if c.get("multiline") else ""
        return f'txt("{name}"{w_arg}{h}{ml}{fo}{ti})'
    if t.startswith("ComboBox"):
        return f'combo("{name}"{w_arg}{fo}{ti})'
    if t.startswith("ListBox"):
        return f'lst("{name}"{w_arg}, height={c["height"]:g}{fo}{ti})'
    if t.startswith("CheckBox"):
        return f'chk("{name}", "{cap}"{fo}{ti})'
    if t.startswith("OptionButton"):
        g = f', group="{c["group"]}"' if c.get("group") else ""
        return f'opt("{name}", "{cap}"{g}{fo}{ti})'
    if t.startswith("CommandButton"):
        acc = f', accel="{c["accelerator"]}"' if c.get("accelerator") else ""
        b = ", bold=True" if c.get("bold") else ""
        # 幅・高さを明示する。出さないと btn() の自動採寸（最小 72×24pt）に
        # 化け、カレンダーの日ボタン（24×20 等）が往復で膨張して格子が崩壊する
        wa = f", width={c['width']:g}"
        h = f", height={c['height']:g}" if abs(c["height"] - 24) > 1 else ""
        if c.get("cancel"):
            return f'cancel("{name}", "{cap}"{wa}{h}{acc}{b}{fo}{ti})'
        if c.get("default"):
            return f'ok("{name}", "{cap}"{wa}{h}{acc}{b}{fo}{ti})'
        return f'btn("{name}", "{cap}"{wa}{h}{acc}{b}{fo}{ti})'
    # form_layout が作れる型は逆変換もできなければ「片肺」になる
    # （作れるのに戻せない＝往復でそのコントロールが消える）
    if t.startswith("Image"):
        return f'img("{name}", {c["width"]:g}, {c["height"]:g})'
    # SpinButton は単体の要素関数が form_layout に無い（spin_txt が TextBox と
    # SpinButton をひと組で作る複合部品）。単純な逆変換にはならないので出力せず、
    # 呼び出し側で「未対応」として目に見える形に残す（黙って消さない）
    return None


def _emit_cluster(r, max_right, indent, tab_map=None):
    """1行クラスタ→宣言コード行（spacer は呼び出し側でブロック間ギャップから出す）"""
    types = [(c.get("type") or "") for c in r]
    # 見出し: 太字ラベル単独の行
    if len(r) == 1 and types[0].startswith("Label") and r[0].get("bold"):
        c0 = r[0]
        cap = (c0.get("caption") or "").replace('"', '\\"')
        # heading() の既定は 13pt。それと違う実測フォントは明示しないと
        # 往復で 13pt に化ける（lbl 系の font= と同じ理由。tab_index も同様）
        fo = ""
        fs = c0.get("font_size")
        try:
            if fs and abs(float(fs) - _HEADING_FONT_PT) >= 0.5:
                fo = f", font={float(fs):g}"
        except Exception:
            fo = ""
        ti = _tab_arg(c0, tab_map)
        return f'{indent}heading("{cap}", name="{c0["name"]}"{fo}{ti}),'
    # ボタンバー: ボタンだけの行。ただし button_bar は「右端寄せ＋全ボタンを
    # 最大幅に統一」する部品なので、無条件に当てるとカレンダーの日ボタン格子の
    # ような並びが右寄せ＋幅72ptに膨らんで崩壊する。少数（〜3個）のときだけ
    # ボタンバーと見なし、それ以外は幅を明示した row として出す
    if all(t.startswith("CommandButton") for t in types) and len(r) <= 3:
        parts = [_emit_elem(c, 0, tab_map) for c in r]
        return f"{indent}button_bar({', '.join(p for p in parts if p)}),"
    parts = []
    skipped = []
    for c in r:
        p = _emit_elem(c, max_right, tab_map)
        if p:
            parts.append(p)
        else:
            skipped.append(f"{c['name']}({c.get('type')})")
    if not parts:
        # 未対応型だけの行。row() は空だと ValueError で落ちるので、行ごと
        # コメントにして「消えた」ことを目に見える形で残す
        return (f"{indent}# ★未対応のため出力できませんでした（手動で追加してください）: "
                f"{', '.join(skipped)}")
    line = f"{indent}row({', '.join(parts)}),"
    if skipped:
        line += f"  # ★未対応のため手動で: {', '.join(skipped)}"
    return line


def _blocks_of(items, max_right, indent, tab_map=None):
    """行クラスタを (top, bottom, [コード行]) のブロック列にする"""
    blocks = []
    for r in _cluster_rows(items):
        top = min(c["top"] for c in r)
        bottom = max(c["top"] + c["height"] for c in r)
        blocks.append((top, bottom, [_emit_cluster(r, max_right, indent, tab_map)]))
    return blocks


def _frame_block(fr, ok_items, indent, tab_map=None):
    """Frame コントロールを frame(...) 宣言ブロックに展開する。

    トップレベルの frame と page 内の frame の両方から使う共通経路。
    （page 内の frame を展開し忘れると、中の子コントロールが宣言の
    どこにも現れず黙って消える）
    """
    children = [c for c in ok_items if c["parent"] == fr["name"]]
    cap = (fr.get("caption") or "").replace('"', '\\"')
    fl = [f'{indent}frame("{cap}",']
    c_inputs = [c for c in children
                if (c.get("type") or "").startswith(("TextBox", "ComboBox", "ListBox"))]
    c_right = max((c["left"] + c["width"] for c in c_inputs), default=0)
    for _, _, lines_ in _blocks_of(children, c_right, indent + "    ", tab_map):
        fl.extend(lines_)
    fl.append(f'{indent}    name="{fr["name"]}"{_tab_arg(fr, tab_map)}),')
    return (fr["top"], fr["top"] + fr["height"], fl)


def _decl_order(plain_items, container_items):
    """宣言（＝build_form の配置）順のコントロール名を返す。

    build_form の TabIndex 自動採番はこの順に 0,1,2... と振る。to_layout_code の
    ブロック並び（top 順・行内は left 順、frame/multipage はそれ自体で1ブロック）
    と一致させてある。
    """
    entries = []
    for r in _cluster_rows(plain_items):
        top = min(c["top"] for c in r)
        for i, c in enumerate(r):      # _cluster_rows は行内を left 昇順に並べる
            entries.append((top, i, c["name"]))
    for c in container_items:
        entries.append((c["top"], 0, c["name"]))
    entries.sort(key=lambda e: (e[0], e[1]))
    return [n for _, _, n in entries]


def _tab_overrides(order, by_name):
    """自動採番（宣言順）と実際の TabIndex が食い違う分だけ、明示指定の dict を返す。

    TabIndex はどこにも出力されていなかったため、往復すると受け側が配置順で
    振り直していた。一覧系フォームの「ListBox が TabIndex 0」のような、
    配置順では表せない意図がここで消える。食い違うときだけ明示する
    （一致していれば自動採番に任せて生成コードを汚さない）。
    """
    have = [n for n in order if by_name.get(n, {}).get("tab_index") is not None]
    if len(have) < 2:
        return {}
    if sorted(have, key=lambda n: by_name[n]["tab_index"]) == have:
        return {}
    return {n: by_name[n]["tab_index"] for n in have}


def to_layout_code(form_name, info, controls):
    """既存フォームを form_layout の宣言コード（たたき台）に逆変換する。

    幾何のクラスタリングによる機械変換。combo/lst の items 等の実行時情報は
    復元できないため、生成コードは出発点として使い、人が仕上げる前提。
    """
    ok_items = [c for c in controls if "error" not in c and c["left"] is not None]
    top_items = [c for c in ok_items if c["parent"] in (None, form_name)
                 and not (c.get("type") or "").startswith(("Frame", "MultiPage"))]
    frames = [c for c in ok_items if (c.get("type") or "").startswith("Frame")
              and c["parent"] in (None, form_name)]
    mpages = [c for c in ok_items if (c.get("type") or "").startswith("MultiPage")
              and c["parent"] in (None, form_name)]

    inputs = [c for c in top_items
              if (c.get("type") or "").startswith(("TextBox", "ComboBox", "ListBox"))]
    max_right = max((c["left"] + c["width"] for c in inputs), default=0)

    # TabIndex はコンテナ（フォーム / Frame / Page）ごとの番号なので、
    # 「宣言順の自動採番と食い違うか」もコンテナ単位で判定する
    by_name = {c["name"]: c for c in ok_items}
    tab_map = dict(_tab_overrides(_decl_order(top_items, frames + mpages), by_name))

    def _add_frame_tabs(fr):
        kids = [c for c in ok_items if c["parent"] == fr["name"]]
        tab_map.update(_tab_overrides(_decl_order(kids, []), by_name))

    for fr in frames:
        _add_frame_tabs(fr)
    for mp in mpages:
        for pg in mp.get("pages", []):
            kids = [c for c in ok_items if c["parent"] == pg["name"]]
            pg_frames = [c for c in kids if (c.get("type") or "").startswith("Frame")]
            plain = [c for c in kids if not (c.get("type") or "").startswith("Frame")]
            tab_map.update(_tab_overrides(_decl_order(plain, pg_frames), by_name))
            for fr in pg_frames:
                _add_frame_tabs(fr)

    # 通常行とフレームを (top, bottom, 行リスト) のブロックに統一して top 順に並べ、
    # ブロック間の縦ギャップから spacer を推定する（フレームの高さも正しく反映される）
    blocks = _blocks_of(top_items, max_right, "    ", tab_map) if top_items else []
    for fr in frames:
        blocks.append(_frame_block(fr, ok_items, "    ", tab_map))
    for mp in mpages:
        ml = [f'    multipage("{mp["name"]}",']
        for pg in mp.get("pages", []):
            children = [c for c in ok_items if c["parent"] == pg["name"]]
            # page 直下の frame は宣言として展開する（form_layout は page 内
            # frame を正式サポートしており、展開しないと子が黙って消える）
            pg_frames = [c for c in children
                         if (c.get("type") or "").startswith("Frame")]
            plain = [c for c in children
                     if not (c.get("type") or "").startswith("Frame")]
            cap = (pg.get("caption") or pg["name"]).replace('"', '\\"')
            ml.append(f'        page("{cap}",')
            c_inputs = [c for c in plain
                        if (c.get("type") or "").startswith(("TextBox", "ComboBox", "ListBox"))]
            c_right = max((c["left"] + c["width"] for c in c_inputs), default=0)
            pblocks = _blocks_of(plain, c_right, "            ", tab_map) if plain else []
            for fr in pg_frames:
                pblocks.append(_frame_block(fr, ok_items, "            ", tab_map))
            pblocks.sort(key=lambda b: b[0])
            for _, _, lines_ in pblocks:
                ml.extend(lines_)
            ml.append("        ),")
        if mp["name"] in tab_map:
            ml.append(f'        tab_index={tab_map[mp["name"]]},')
        ml.append("    ),")
        blocks.append((mp["top"], mp["top"] + mp["height"], ml))
    blocks.sort(key=lambda b: b[0])

    out = ["# form_inspect --to-layout による自動生成（たたき台。items等の実行時情報は要手動）",
           "from form_layout import (build_form, row, lbl, txt, combo, lst, chk, opt,",
           "                         btn, ok, cancel, button_bar, spacer, heading, frame,",
           "                         multipage, page, img)",
           "",
           "rows = ["]
    prev_bottom = None
    for top, bottom, lines_ in blocks:
        if prev_bottom is not None:
            gap = top - prev_bottom
            if gap > 14:
                out.append("    spacer()," if gap <= 24 else f"    spacer({gap:g}),")
        prev_bottom = bottom
        out.extend(lines_)
    out.append("]")
    cap = (info.get("caption") or form_name).replace('"', '\\"')
    out.append("")
    out.append(f'build_form("{form_name}", "{cap}", rows, vba_stub=False, png=True)')
    return "\n".join(out)


# ================================================================
# PNG 撮影（実表示 + BitBlt）
# ================================================================

def _save_bmp(path, w, h, bits):
    """32bit トップダウン DIB として BMP を書く（追加ライブラリ不要）。"""
    import struct
    header_size = 14 + 40
    with open(path, "wb") as fp:
        fp.write(struct.pack("<2sIHHI", b"BM", header_size + len(bits), 0, 0, header_size))
        fp.write(struct.pack("<IiiHHIIiiII", 40, w, -h, 1, 32, 0, len(bits), 0, 0, 0, 0))
        fp.write(bits)


def _save_bits(out_path, w, h, bits, echo=print):
    """BGRX の生ビット列を PNG(Pillow) か BMP(標準のみ) で保存し、実際のパスを返す。"""
    if out_path.lower().endswith(".png"):
        try:
            from PIL import Image
        except ImportError:
            out_path = os.path.splitext(out_path)[0] + ".bmp"
            echo("Pillow が無いため BMP で保存します。")
        else:
            Image.frombuffer("RGB", (w, h), bits, "raw", "BGRX", 0, 1).save(out_path)
            return out_path
    _save_bmp(out_path, w, h, bits)
    return out_path


def _draw_overlay(png_path, dx, dy, scale, controls, echo=print):
    """PNG にコントロールの枠と名前を描き込む（--names。Pillow 必須）。

    dx/dy: ウィンドウ左上→クライアント領域左上のオフセット(px)。
    scale: pt→px 換算係数。座標が親相対の Frame 内コントロールは対象外。
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        echo("Pillow が無いため --names の描き込みはスキップしました。")
        return False
    img = Image.open(png_path).convert("RGB")
    drw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("meiryo.ttc", 11)
    except Exception:
        font = ImageFont.load_default()
    for c in controls:
        if "error" in c or c["left"] is None:
            continue
        x1 = dx + c["left"] * scale
        y1 = dy + c["top"] * scale
        x2 = x1 + c["width"] * scale
        y2 = y1 + c["height"] * scale
        drw.rectangle([x1, y1, x2, y2], outline=(255, 0, 0), width=1)
        label = c["name"]
        tx, ty = x1 + 1, max(0, y1 - 12)
        bbox = drw.textbbox((tx, ty), label, font=font)
        drw.rectangle(bbox, fill=(255, 255, 160))
        drw.text((tx, ty), label, fill=(180, 0, 0), font=font)
    img.save(png_path)
    echo(f"名前オーバーレイを描き込みました: {png_path}")
    return True


def render_form_png(xl, wb, form_name, out_path, echo=print, overlay=None,
                    all_pages=None):
    """フォームを実表示（Show vbModeless）して BitBlt でキャプチャし PNG/BMP 保存する。

    overlay に {'inside_w', 'inside_h', 'controls'} を渡すと、フォーム直下の
    コントロールの枠と名前を PNG に描き込む（--names。私＝AIが画像を見て
    「btnOK を右に8」と名指しで直せるようにする）。
    all_pages に [(MultiPage名, [タブ名,...]), ...] を渡すと、タブを切り替えながら
    ページごとに <出力名>_tab<番号>.png を追加保存する（--png-all）。

    技法（C:\\tmp\\render_form2.py の汎化）:
      一時標準モジュールに Show/Unload の Sub を注入して xl.Run で実行し、
      "ThunderDFrame" ウィンドウを TOPMOST 化して BitBlt で撮る
      （PrintWindow は黒帯が出るので使わない）。
      finally で Unload と一時モジュール削除。保存したパスを返す（失敗時 None）。

      後始末の原則:
      - 一時モジュールの出し入れでブックが「未保存」に落ちないよう、撮影前の
        wb.Saved を控えて後で戻す（「目」であるはずの --png がブックを汚さない）。
      - Unload に失敗するとフォームはモードレス表示のまま残る。そこで一時
        モジュールまで消すと閉じる手段が無くなるので、失敗時はモジュールを
        残して閉じ方を案内する（黙って握りつぶさない）。
    """
    import time
    import ctypes
    import win32gui
    import win32ui
    import win32con

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

    vbp = wb.VBProject
    # 一時モジュールの Add/Remove でブックは「未保存」に落ちる。撮影前の状態を
    # 控えておき、後始末まで済んだら戻す（--png は見るだけの操作＝ブックを汚さない）
    was_saved = None
    try:
        was_saved = bool(wb.Saved)
    except Exception as e:
        echo(f"  wb.Saved の取得に失敗（保存状態の復元はスキップ）: {e}")
    # 撮影対象の特定用に Caption を控える。クラス名だけの FindWindow だと、
    # ユーザーが別フォームをモードレス表示中に「別の窓」を撮ってしまう
    form_caption = None
    try:
        for c in vbp.VBComponents:
            if c.Name.lower() == str(form_name).lower():
                form_caption = str(c.Properties("Caption").Value)
                break
    except Exception:
        pass
    mod = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
    code = (
        "Sub tmpFormInspectShow()\r\n"
        f"    {form_name}.Show vbModeless\r\n"
        "End Sub\r\n"
        "Sub tmpFormInspectUnload()\r\n"
        f"    Unload {form_name}\r\n"
        "End Sub\r\n"
        "Sub tmpFormInspectSetPage(nm As String, ByVal i As Integer)\r\n"
        f"    {form_name}.Controls(nm).Value = i\r\n"
        "End Sub\r\n"
    )
    mod.CodeModule.AddFromString(code)
    pre = "'" + wb.Name + "'!"
    shown = False
    try:
        xl.Run(pre + "tmpFormInspectShow")
        shown = True
        time.sleep(0.5)

        def _find_form_hwnd():
            hits = []
            def _cb(h, _):
                try:
                    if (win32gui.GetClassName(h) == "ThunderDFrame"
                            and win32gui.IsWindowVisible(h)):
                        hits.append(h)
                except Exception:
                    pass
            win32gui.EnumWindows(_cb, None)
            if form_caption:
                for h in hits:
                    if win32gui.GetWindowText(h) == form_caption:
                        return h
            return hits[0] if hits else 0

        hwnd = _find_form_hwnd()
        if not hwnd:
            echo("フォームウィンドウ (ThunderDFrame) が見つかりません。")
            return None
        try:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            echo(f"  最前面化の警告 (続行): {e}")
        ctypes.windll.user32.UpdateWindow(hwnd)
        time.sleep(0.4)

        l, t, r, b = win32gui.GetWindowRect(hwnd)
        w, h = r - l, b - t
        # --names 用: クライアント領域のオフセットと pt→px 係数を hwnd が生きている間に取る
        ov_geo = None
        if overlay:
            try:
                _, _, cw_px, ch_px = win32gui.GetClientRect(hwnd)
                cx, cy = win32gui.ClientToScreen(hwnd, (0, 0))
                iw = overlay.get('inside_w') or 0
                if iw and cw_px:
                    ov_geo = (cx - l, cy - t, cw_px / iw)
            except Exception as e:
                echo(f"  オーバーレイ座標の取得に失敗 (枠なしで保存): {e}")
        def snap(path):
            """現在の見た目を BitBlt で1枚キャプチャして保存する"""
            hdc_screen = win32gui.GetDC(0)
            mfc = win32ui.CreateDCFromHandle(hdc_screen)
            mem = mfc.CreateCompatibleDC()
            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(mfc, w, h)
            mem.SelectObject(bmp)
            mem.BitBlt((0, 0), (w, h), mfc, (l, t), win32con.SRCCOPY)
            bits = bmp.GetBitmapBits(True)
            try:
                return _save_bits(path, w, h, bits, echo)
            finally:
                win32gui.DeleteObject(bmp.GetHandle())
                mem.DeleteDC()
                mfc.DeleteDC()
                win32gui.ReleaseDC(0, hdc_screen)

        saved = snap(out_path)
        if saved:
            echo(f"保存: {saved}  ({w}x{h})")
            if overlay and ov_geo and saved.lower().endswith(".png"):
                dx, dy, scale = ov_geo
                ctrls = [c for c in overlay.get('controls', [])
                         if c.get("parent") in (None, form_name)]
                _draw_overlay(saved, dx, dy, scale, ctrls, echo)
            # --png-all: MultiPage のタブを切り替えながら全ページを撮る
            if all_pages:
                base, ext = os.path.splitext(saved)
                for mp_name, captions in all_pages:
                    for pi, cap in enumerate(captions):
                        try:
                            xl.Run(pre + "tmpFormInspectSetPage", mp_name, pi)
                        except Exception as e:
                            echo(f"  タブ切替に失敗 ({mp_name}→{pi}): {e}")
                            continue
                        time.sleep(0.3)
                        p = snap(f"{base}_{mp_name}_tab{pi + 1}{ext}")
                        if p:
                            echo(f"保存: {p}  （タブ: {cap}）")
                    try:
                        xl.Run(pre + "tmpFormInspectSetPage", mp_name, 0)
                    except Exception:
                        pass
        return saved
    finally:
        unload_failed = False
        if shown:
            try:
                xl.Run(pre + "tmpFormInspectUnload")
            except Exception as e:
                unload_failed = True
                echo(f"⚠ フォーム {form_name} の Unload に失敗しました: {e}")
                echo("  フォームがモードレス表示のまま残っている可能性があります。"
                     "閉じる手段として一時モジュールを残します"
                     f"（Excel で {pre}tmpFormInspectUnload を実行、"
                     f"またはフォームの × で閉じてから {mod.Name} を削除してください）。")
        if not unload_failed:
            try:
                vbp.VBComponents.Remove(mod)
            except Exception as e:
                unload_failed = True   # 掃除しきれていない＝保存状態は戻さない
                echo(f"⚠ 一時モジュール {mod.Name} の削除に失敗しました: {e}")
        # 後始末が完了したときだけ保存状態を戻す。残置物があるのに Saved=True に
        # すると「変更なし」の顔で本当の変更を隠してしまう
        if was_saved and not unload_failed:
            try:
                wb.Saved = True
            except Exception as e:
                echo(f"⚠ wb.Saved の復元に失敗（ブックが未保存状態のままです）: {e}")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    ap = argparse.ArgumentParser(
        description="UserForm を1回のCOM接続でまとめて把握する")
    ap.add_argument("form", nargs="?", help="フォーム名")
    ap.add_argument("procs", nargs="*", help="プロシージャ名 (省略時は全コード)")
    ap.add_argument("--list", action="store_true", dest="list_forms",
                    help="開いているブックの全フォーム名")
    ap.add_argument("--font", action="store_true",
                    help="コントロール一覧に Font.Size 列を追加")
    ap.add_argument("--json", action="store_true", dest="as_json",
                    help="機械可読 JSON で stdout に出力")
    ap.add_argument("--png", nargs="?", const=DEFAULT_PNG, default=None,
                    metavar="出力パス", help="フォームを実表示して PNG 撮影")
    ap.add_argument("--names", action="store_true",
                    help="--png にコントロールの枠と名前を描き込む（要 Pillow）")
    ap.add_argument("--lint", action="store_true",
                    help="重なり・はみ出し・不揃い・タブ順を機械検査して報告")
    ap.add_argument("--to-layout", action="store_true", dest="to_layout",
                    help="既存フォームを form_layout の宣言コード（たたき台）に逆変換して出力")
    ap.add_argument("--png-all", action="store_true", dest="png_all",
                    help="MultiPage の全タブを切り替えながら1枚ずつ撮る（--png を含意）")
    args = ap.parse_args()
    if args.png_all and not args.png:
        args.png = DEFAULT_PNG

    if not args.list_forms and not args.form:
        print("使い方: py form_inspect.py <フォーム名> [proc ...] "
              "[--font] [--json] [--png [パス]]   /   --list")
        return

    # --json 時は人間向けメッセージを stderr へ逃がす（stdout は JSON 専用）
    say = (lambda *a: print(*a, file=sys.stderr)) if args.as_json else print

    xl = get_active_excel()
    wb = xl.ActiveWorkbook
    if wb is None:
        say("アクティブなブックがありません。対象ブックを開いてください。")
        return

    if args.list_forms:
        print(f"対象ブック: {wb.Name}  フォーム一覧:")
        for c in wb.VBProject.VBComponents:
            # 3 = vbext_ct_MSForm
            if c.Type == 3:
                print("  -", c.Name)
        return

    form_name = args.form
    procs = args.procs
    if not args.as_json:
        # --to-layout は stdout がそのまま実行可能コードになるよう、案内は stderr へ
        print(f"対象ブック: {wb.Name}   フォーム: {form_name}",
              file=sys.stderr if args.to_layout else sys.stdout)

    comp = find_component(wb, form_name)
    if comp is None:
        say(f"フォーム '{form_name}' が見つかりません。VBComponents:")
        for c in wb.VBProject.VBComponents:
            say("  -", c.Name)
        return

    info = get_form_info(comp)

    if args.as_json:
        try:
            controls = collect_controls(comp)
        except RuntimeError as e:
            controls = []
            say(f"({e})")
        png_path = None
        if args.png:
            ov = ({'inside_w': info["inside_width"], 'inside_h': info["inside_height"],
                   'controls': controls} if args.names else None)
            ap_pages = ([(c["name"], [p["caption"] for p in c.get("pages", [])])
                         for c in controls
                         if (c.get("type") or "").startswith("MultiPage")]
                        if args.png_all else None)
            png_path = render_form_png(xl, wb, comp.Name, args.png, echo=say,
                                       overlay=ov, all_pages=ap_pages or None)
        data = {
            "book": wb.Name,
            "form": comp.Name,
            "caption": info["caption"],
            "width": _round2(info["width"]),
            "height": _round2(info["height"]),
            "inside_width": _round2(info["inside_width"]),
            "inside_height": _round2(info["inside_height"]),
            "controls": controls,
        }
        if procs:
            cm = comp.CodeModule
            data["procs"] = {p: get_proc_code(cm, p) for p in procs}
        if args.png:
            data["png"] = png_path
        if args.lint:
            cm = comp.CodeModule
            code_txt = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines > 0 else None
            data["lint"] = lint_form(comp.Name, info, controls, code=code_txt)
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return

    if args.to_layout:
        try:
            controls = collect_controls(comp)
        except RuntimeError as e:
            print(f"(逆変換不可: {e})")
            return
        print(to_layout_code(comp.Name, info, controls))
        return

    print_form_info(info)
    print_controls(comp, want_font=args.font)
    print_code(comp, procs)
    if args.lint:
        try:
            cm = comp.CodeModule
            code_txt = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines > 0 else None
            print_lint(comp.Name, info, collect_controls(comp), code=code_txt)
        except RuntimeError as e:
            print(f"(lint 不可: {e})")
    if args.png:
        ov = None
        ap_pages = None
        try:
            ctrls_png = collect_controls(comp)
        except RuntimeError:
            ctrls_png = []
        if args.names and ctrls_png:
            ov = {'inside_w': info["inside_width"], 'inside_h': info["inside_height"],
                  'controls': ctrls_png}
        if args.png_all:
            ap_pages = [(c["name"], [p["caption"] for p in c.get("pages", [])])
                        for c in ctrls_png
                        if (c.get("type") or "").startswith("MultiPage")] or None
        render_form_png(xl, wb, comp.Name, args.png, overlay=ov, all_pages=ap_pages)


if __name__ == "__main__":
    main()
