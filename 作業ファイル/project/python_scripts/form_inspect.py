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
    # 既に一般名（Label/TextBox…）ならそのまま。未知のI〜名は接頭Iを外して返す
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
    by_parent = {}
    for c in ok_items:
        by_parent.setdefault(c["parent"], []).append(c)
    for parent, items in by_parent.items():
        for i in range(len(items)):
            for j in range(i + 1, len(items)):
                a, b = rect(items[i]), rect(items[j])
                if a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]:
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

    # 3. ボタンのサイズ不揃い（同じ行のボタン同士だけ比較。
    #    行が違えば役割も違う＝入力欄に高さを合わせた「選択」ボタン等は対象外）
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
            ws = sorted({c["width"] for c in bs})
            if len(hs) > 1:
                findings.append(f"同じ行のボタン高さが不揃い: {hs}  "
                                f"({', '.join(c['name'] for c in bs)})")
            if len(ws) > 2:
                findings.append(f"同じ行のボタン幅が3種以上: {ws}  "
                                f"({', '.join(c['name'] for c in bs)})")

    # 4. フォントサイズの混在（見出し＝太字ラベルは意図的な大きさなので対象外）
    sizes = {}
    for c in ok_items:
        if (c.get("type") or "").startswith("Label") and c.get("bold"):
            continue
        if c["font_size"] is not None:
            sizes.setdefault(c["font_size"], []).append(c["name"])
    if len(sizes) > 1:
        major = max(sizes, key=lambda k: len(sizes[k]))
        for sz, names in sorted(sizes.items()):
            if sz != major:
                findings.append(f"フォント混在: {sz}pt = {', '.join(names)}（多数派は {major}pt）")

    # 5. TabIndex が視線順（上→下、左→右）と食い違う
    #    行内の上下中央合わせでラベルと入力の top はずれるため、比較は「上下中央」で行う
    tabbed = [c for c in ok_items if c["tab_index"] is not None
              and c["parent"] in (None, form_name)]
    visual = sorted(tabbed, key=lambda c: (round((c["top"] + c["height"] / 2) / 14), c["left"]))
    actual = sorted(tabbed, key=lambda c: c["tab_index"])
    mism = [(v["name"]) for v, a in zip(visual, actual) if v["name"] != a["name"]]
    if mism:
        findings.append(f"TabIndex が視線順と不一致: {', '.join(mism[:6])}"
                        + ("…" if len(mism) > 6 else ""))

    # 5b. 右端の「揃いかけのズレ」（最大右端から 2〜16pt だけ短い入力は
    #    揃え忘れの可能性が高い。大きく短いのは意図的な短欄とみなして対象外）
    inputs = [c for c in ok_items if c["parent"] in (None, form_name)
              and (c.get("type") or "").startswith(("TextBox", "ComboBox", "ListBox"))]
    if len(inputs) >= 2:
        max_r = max(round(c["left"] + c["width"]) for c in inputs)
        near_miss = [c for c in inputs
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

    # 5d. Default / Cancel / Accelerator（キーボード操作の作法）
    btns2 = [c for c in ok_items if (c.get("type") or "").startswith("CommandButton")]
    if btns2:
        if not any(c.get("default") for c in btns2):
            findings.append("Default(Enter) のボタンが未設定")
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

    # 6. 左端ラインの数（縦の整列ライン。2〜3本が目安）
    #    右寄せのボタンバーと行内のチェック/オプションは設計上ラインが増えるので対象外
    line_kinds = ('Label', 'TextBox', 'ComboBox', 'ListBox', 'Frame')
    lefts = sorted({round(c["left"]) for c in ok_items
                    if c["parent"] in (None, form_name)
                    and any((c.get("type") or "").startswith(k) for k in line_kinds)})
    if len(lefts) >= 4:
        findings.append(f"左端の縦ライン（ラベル/入力系）が{len(lefts)}本: {lefts}（2〜3本が目安）")

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


def _emit_elem(c, max_right):
    """1コントロールを form_layout の要素式に変換する。未対応種別は None"""
    t = c.get("type") or ""
    name = c["name"]
    cap = (c.get("caption") or "").replace('"', '\\"')
    stretch = abs((c["left"] + c["width"]) - max_right) <= 1 if max_right else False
    w_arg = "" if stretch else f", width={c['width']:g}"

    if t.startswith("Label"):
        b = ", bold=True" if c.get("bold") else ""
        return f'lbl("{cap}"{b})'
    if t.startswith("TextBox"):
        h = f", height={c['height']:g}" if abs(c["height"] - 22) > 1 else ""
        ml = ", multiline=True" if c.get("multiline") else ""
        return f'txt("{name}"{w_arg}{h}{ml})'
    if t.startswith("ComboBox"):
        return f'combo("{name}"{w_arg})'
    if t.startswith("ListBox"):
        return f'lst("{name}"{w_arg}, height={c["height"]:g})'
    if t.startswith("CheckBox"):
        return f'chk("{name}", "{cap}")'
    if t.startswith("OptionButton"):
        g = f', group="{c["group"]}"' if c.get("group") else ""
        return f'opt("{name}", "{cap}"{g})'
    if t.startswith("CommandButton"):
        acc = f', accel="{c["accelerator"]}"' if c.get("accelerator") else ""
        if c.get("cancel"):
            return f'cancel("{name}", "{cap}"{acc})'
        if c.get("default"):
            return f'ok("{name}", "{cap}"{acc})'
        return f'btn("{name}", "{cap}"{acc})'
    return None


def _emit_cluster(r, max_right, indent):
    """1行クラスタ→宣言コード行（spacer は呼び出し側でブロック間ギャップから出す）"""
    types = [(c.get("type") or "") for c in r]
    # 見出し: 太字ラベル単独の行
    if len(r) == 1 and types[0].startswith("Label") and r[0].get("bold"):
        cap = (r[0].get("caption") or "").replace('"', '\\"')
        return f'{indent}heading("{cap}"),'
    # ボタンバー: ボタンだけの行（通常は最終行・右寄せ）
    if all(t.startswith("CommandButton") for t in types):
        parts = [_emit_elem(c, 0) for c in r]
        return f"{indent}button_bar({', '.join(p for p in parts if p)}),"
    parts = []
    skipped = []
    for c in r:
        p = _emit_elem(c, max_right)
        if p:
            parts.append(p)
        else:
            skipped.append(f"{c['name']}({c.get('type')})")
    line = f"{indent}row({', '.join(parts)}),"
    if skipped:
        line += f"  # 未対応のため手動で: {', '.join(skipped)}"
    return line


def _blocks_of(items, max_right, indent):
    """行クラスタを (top, bottom, [コード行]) のブロック列にする"""
    blocks = []
    for r in _cluster_rows(items):
        top = min(c["top"] for c in r)
        bottom = max(c["top"] + c["height"] for c in r)
        blocks.append((top, bottom, [_emit_cluster(r, max_right, indent)]))
    return blocks


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

    # 通常行とフレームを (top, bottom, 行リスト) のブロックに統一して top 順に並べ、
    # ブロック間の縦ギャップから spacer を推定する（フレームの高さも正しく反映される）
    blocks = _blocks_of(top_items, max_right, "    ") if top_items else []
    for fr in frames:
        children = [c for c in ok_items if c["parent"] == fr["name"]]
        cap = (fr.get("caption") or "").replace('"', '\\"')
        fl = [f'    frame("{cap}",']
        c_inputs = [c for c in children
                    if (c.get("type") or "").startswith(("TextBox", "ComboBox", "ListBox"))]
        c_right = max((c["left"] + c["width"] for c in c_inputs), default=0)
        for _, _, lines_ in _blocks_of(children, c_right, "        "):
            fl.extend(lines_)
        fl.append("    ),")
        blocks.append((fr["top"], fr["top"] + fr["height"], fl))
    for mp in mpages:
        ml = [f'    multipage("{mp["name"]}",']
        for pg in mp.get("pages", []):
            children = [c for c in ok_items if c["parent"] == pg["name"]]
            cap = (pg.get("caption") or pg["name"]).replace('"', '\\"')
            ml.append(f'        page("{cap}",')
            c_inputs = [c for c in children
                        if (c.get("type") or "").startswith(("TextBox", "ComboBox", "ListBox"))]
            c_right = max((c["left"] + c["width"] for c in c_inputs), default=0)
            for _, _, lines_ in _blocks_of(children, c_right, "            "):
                ml.extend(lines_)
            ml.append("        ),")
        ml.append("    ),")
        blocks.append((mp["top"], mp["top"] + mp["height"], ml))
    blocks.sort(key=lambda b: b[0])

    out = ["# form_inspect --to-layout による自動生成（たたき台。items等の実行時情報は要手動）",
           "from form_layout import (build_form, row, lbl, txt, combo, lst, chk, opt,",
           "                         btn, ok, cancel, button_bar, spacer, heading, frame,",
           "                         multipage, page)",
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
        hwnd = win32gui.FindWindow("ThunderDFrame", None)
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
        if shown:
            try:
                xl.Run(pre + "tmpFormInspectUnload")
            except Exception:
                pass
        try:
            vbp.VBComponents.Remove(mod)
        except Exception:
            pass


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
