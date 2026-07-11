# -*- coding: utf-8 -*-
"""
form_tool.py — UserForm の機械的な幾何操作を「1回のCOM接続」で行う CLI。

form_inspect.py（目）と対になる「手」。座標・サイズ・整列・リネーム・削除といった
機械的な操作だけを担当し、何をどう変えるかの判断は呼び出し側が行う。
接続は form_inspect.py と同方式（GetActiveObject → dynamic.Dispatch、gen_py 非依存）。
対象は常に「今アクティブに開いているブック」。特定ブック・パスに依存しない。

使い方:
  py form_tool.py scale <フォーム名> <倍率> [--save]
      フォーム自身と全コントロールの L/T/W/H と Font.Size に倍率を適用
      （Font.Size は round）

  py form_tool.py set <フォーム名> <コントロール名[,名2,...]>
      [--left N] [--top N] [--width N] [--height N]
      [--caption 文字] [--font-size N] [--default] [--cancel] [--save]
      指定コントロール（カンマ区切りで複数可）に同じ変更を一括適用
      --default / --cancel は CommandButton の Default(Enter) / Cancel(Esc)
      プロパティを True にする（フォーム内で既に別のボタンが True の場合は
      フォームエンジンの仕様により自動的にそちらが False に切り替わる）

  py form_tool.py align <フォーム名> <コントロール名,...> (--top N | --left N) [--save]
      指定コントロールの上辺 / 左辺を揃える

  py form_tool.py rename-control <フォーム名> <旧名> <新名> [--save]
      コントロール名を変更し、CodeModule のイベントプロシージャ宣言行
      （^(Private |Public )?Sub 旧名_...）だけを機械的に追随（置換件数を報告）

  py form_tool.py delete-control <フォーム名> <コントロール名> [--save]
      削除前に種類・位置・Caption を表示。イベントプロシージャが残る場合は
      警告のみ（コードは削除しない＝要否の判断はユーザー）

変更系はどれも既定では保存しない。--save 指定時のみ wb.Save() する。
"""
import argparse
import os
import re
import sys
import time

from form_inspect import get_active_excel

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', 'backups'))


def backup_form(wb, comp):
    """破壊的操作（rename/delete/copy元の改変等）の前に .frm/.frx を退避する安全網"""
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        base = os.path.splitext(os.path.basename(wb.FullName))[0]
        path = os.path.join(BACKUP_DIR, f"{base}_{comp.Name}_{stamp}.frm")
        comp.Export(path)
        print(f"フォーム退避: backups/{os.path.basename(path)} (+.frx)")
    except Exception as e:
        print(f"⚠ フォーム退避に失敗（続行）: {e}")


# ================================================================
# 接続・解決ヘルパー
# ================================================================

def connect_form(form_name):
    """アクティブブックに接続し (xl, wb, フォームの VBComponent) を返す。"""
    try:
        xl = get_active_excel()
    except Exception:
        sys.exit("起動中の Excel が見つかりません。"
                 "対象ブックを Excel で開いてから実行してください。")
    wb = xl.ActiveWorkbook
    if wb is None:
        sys.exit("アクティブなブックがありません。対象ブックを開いてください。")
    try:
        comps = wb.VBProject.VBComponents
    except Exception as e:
        sys.exit(
            "VBProject にアクセスできません。Excel の設定で\n"
            "  ファイル > オプション > トラストセンター > トラストセンターの設定 >\n"
            "  マクロの設定 > 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」\n"
            f"を有効にしてから再実行してください。({e})")
    comp = None
    forms = []
    for c in comps:
        if c.Type == 3:  # vbext_ct_MSForm
            forms.append(c.Name)
        if c.Name.lower() == form_name.lower():
            comp = c
    if comp is None:
        sys.exit(f"フォーム '{form_name}' が見つかりません。存在するフォーム:\n"
                 + "\n".join("  - " + n for n in forms))
    if comp.Type != 3:
        sys.exit(f"'{comp.Name}' はフォームではありません (Type={comp.Type})。"
                 "存在するフォーム:\n" + "\n".join("  - " + n for n in forms))
    return xl, wb, comp


def pick_controls(comp, names_csv):
    """カンマ区切りのコントロール名を Designer.Controls から解決して返す。"""
    pairs = [(ct.Name, ct) for ct in comp.Designer.Controls]
    by_lower = {n.lower(): ct for n, ct in pairs}
    wanted = [s.strip() for s in names_csv.split(",") if s.strip()]
    if not wanted:
        sys.exit("コントロール名を指定してください。")
    missing = [n for n in wanted if n.lower() not in by_lower]
    if missing:
        sys.exit("コントロールが見つかりません: " + ", ".join(missing)
                 + "\n存在するコントロール:\n"
                 + "\n".join("  - " + n for n, _ in pairs))
    return [by_lower[n.lower()] for n in wanted]


def control_type(ct):
    """コントロールの種類名 (CommandButton 等) を返す。取得できなければ '?'。"""
    try:
        return ct._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
    except Exception:
        return "?"


def _fmt(v):
    return f"{float(v):g}"


def finish(wb, save):
    """変更系コマンドの締め。--save 指定時のみ保存する。"""
    if save:
        wb.Save()
        print(f"保存しました: {wb.Name}")
    else:
        print("（保存はしていません。確定するには --save を付けるか Excel 側で保存）")


# ================================================================
# 各コマンド
# ================================================================

def cmd_scale(args):
    if args.factor <= 0:
        sys.exit("倍率は正の数で指定してください。")
    xl, wb, comp = connect_form(args.form)
    k = args.factor
    # フォームの Width/Height は枠・タイトルバー（クローム）込み。クロームまで
    # 倍率を掛けると、縮小時に内寸が「コントロール座標×k」より狭くなって
    # 右端・下端が見切れる。内寸にだけ倍率を掛け、クロームは実寸のまま保つ
    try:
        chrome_w = float(comp.Properties("Width").Value) - float(comp.Designer.InsideWidth)
        chrome_h = float(comp.Properties("Height").Value) - float(comp.Designer.InsideHeight)
    except Exception:
        chrome_w, chrome_h = 11.0, 29.0   # 取得できない場合の実測既定値
    for prop, chrome in (("Width", chrome_w), ("Height", chrome_h)):
        p = comp.Properties(prop)
        old = float(p.Value)
        new = chrome + (old - chrome) * k
        p.Value = new
        print(f"フォーム {prop}: {_fmt(old)} -> {_fmt(new)}")
    n_geo = 0
    n_font = 0
    for ct in comp.Designer.Controls:
        ct.Left = float(ct.Left) * k
        ct.Top = float(ct.Top) * k
        ct.Width = float(ct.Width) * k
        ct.Height = float(ct.Height) * k
        n_geo += 1
        try:
            fs = float(ct.Font.Size)
            ct.Font.Size = max(1, round(fs * k))
            n_font += 1
        except Exception:
            pass
    print(f"倍率 {_fmt(k)} を適用: コントロール {n_geo}個 (L/T/W/H)、"
          f"うち Font.Size 変更 {n_font}個")
    finish(wb, args.save)


def cmd_set(args):
    geo = [("Left", args.left), ("Top", args.top),
           ("Width", args.width), ("Height", args.height)]
    if (all(v is None for _, v in geo)
            and args.caption is None and args.font_size is None
            and not args.default and not args.cancel):
        sys.exit("変更する項目 (--left/--top/--width/--height/--caption/--font-size/"
                 "--default/--cancel) を1つ以上指定してください。")
    xl, wb, comp = connect_form(args.form)
    for ct in pick_controls(comp, args.controls):
        parts = []
        for prop, v in geo:
            if v is not None:
                setattr(ct, prop, v)
                parts.append(f"{prop}={_fmt(v)}")
        if args.caption is not None:
            try:
                ct.Caption = args.caption
                parts.append(f'Caption="{args.caption}"')
            except Exception as e:
                parts.append(f"Caption 設定不可 ({e})")
        if args.font_size is not None:
            try:
                ct.Font.Size = args.font_size
                parts.append(f"Font.Size={_fmt(args.font_size)}")
            except Exception as e:
                parts.append(f"Font.Size 設定不可 ({e})")
        if args.default:
            try:
                ct.Default = True
                parts.append("Default=True")
            except Exception as e:
                parts.append(f"Default 設定不可 ({e})")
        if args.cancel:
            try:
                ct.Cancel = True
                parts.append("Cancel=True")
            except Exception as e:
                parts.append(f"Cancel 設定不可 ({e})")
        print(f"  {ct.Name}: " + "  ".join(parts))
    finish(wb, args.save)


def cmd_align(args):
    if (args.top is None) == (args.left is None):
        sys.exit("--top か --left のどちらか一方を指定してください。")
    xl, wb, comp = connect_form(args.form)
    prop, v = ("Top", args.top) if args.top is not None else ("Left", args.left)
    ctrls = pick_controls(comp, args.controls)
    for ct in ctrls:
        setattr(ct, prop, v)
        print(f"  {ct.Name}: {prop}={_fmt(v)}")
    print(f"{len(ctrls)}個の {prop} を {_fmt(v)} に揃えました。")
    finish(wb, args.save)


def cmd_rename(args):
    xl, wb, comp = connect_form(args.form)
    backup_form(wb, comp)
    names = []
    target = None
    for ct in comp.Designer.Controls:
        names.append(ct.Name)
        if ct.Name.lower() == args.old.lower():
            target = ct
    if target is None:
        sys.exit(f"コントロール '{args.old}' が見つかりません。存在するコントロール:\n"
                 + "\n".join("  - " + n for n in names))
    if any(n.lower() == args.new.lower() for n in names):
        sys.exit(f"新しい名前 '{args.new}' は既に存在します。")
    old_actual = target.Name
    target.Name = args.new
    print(f"コントロール名変更: {old_actual} -> {args.new}")

    # イベントプロシージャの宣言行だけを機械的に追随
    # （\b旧名_\w+\b の全文置換は呼び出し・文字列も巻き込むため宣言行に限定）
    cm = comp.CodeModule
    pat = re.compile(r"^(Private\s+|Public\s+)?Sub\s+" + re.escape(old_actual) + r"_",
                     re.IGNORECASE)
    count = 0
    for i in range(1, cm.CountOfLines + 1):
        line = cm.Lines(i, 1)
        m = pat.match(line)
        if m:
            new_line = (m.group(1) or "") + "Sub " + args.new + "_" + line[m.end():]
            cm.ReplaceLine(i, new_line)
            count += 1
    print(f"イベントプロシージャ宣言行の置換: {count}件"
          "（宣言行のみ。呼び出し側などにあれば手動確認）")
    finish(wb, args.save)


def cmd_move(args):
    """相対移動: move <form> <controls,> --dx N --dy N"""
    if args.dx is None and args.dy is None:
        sys.exit("--dx か --dy を指定してください（相対移動量）。")
    xl, wb, comp = connect_form(args.form)
    dx = args.dx or 0
    dy = args.dy or 0
    for ct in pick_controls(comp, args.controls):
        ct.Left = float(ct.Left) + dx
        ct.Top = float(ct.Top) + dy
        print(f"  {ct.Name}: L{_fmt(ct.Left)} T{_fmt(ct.Top)}  (Δ{_fmt(dx)},{_fmt(dy)})")
    finish(wb, args.save)


def cmd_size_match(args):
    """同サイズ化: size-match <form> <controls,> [--ref 基準名] [--width-only|--height-only]"""
    xl, wb, comp = connect_form(args.form)
    ctrls = pick_controls(comp, args.controls)
    if args.ref:
        ref = pick_controls(comp, args.ref)[0]
    else:
        ref = ctrls[0]
    rw, rh = float(ref.Width), float(ref.Height)
    for ct in ctrls:
        if ct.Name == ref.Name:
            continue
        if not args.height_only:
            ct.Width = rw
        if not args.width_only:
            ct.Height = rh
        print(f"  {ct.Name}: W{_fmt(ct.Width)} H{_fmt(ct.Height)}")
    print(f"基準: {ref.Name} (W{_fmt(rw)} H{_fmt(rh)})")
    finish(wb, args.save)


def cmd_distribute(args):
    """等間隔配置: distribute <form> <controls,> (--vertical|--horizontal) [--gap N]

    位置順に並べ、先頭の位置を起点に指定 gap（省略時は現在の平均間隔）で等間隔化する。
    """
    if args.vertical == args.horizontal:
        sys.exit("--vertical か --horizontal のどちらか一方を指定してください。")
    xl, wb, comp = connect_form(args.form)
    ctrls = pick_controls(comp, args.controls)
    if len(ctrls) < 2:
        sys.exit("2個以上のコントロールを指定してください。")
    if args.vertical:
        ctrls.sort(key=lambda c: float(c.Top))
        if args.gap is not None:
            gap = args.gap
        else:
            spans = [float(ctrls[i + 1].Top) - (float(ctrls[i].Top) + float(ctrls[i].Height))
                     for i in range(len(ctrls) - 1)]
            gap = sum(spans) / len(spans)
        y = float(ctrls[0].Top)
        for ct in ctrls:
            ct.Top = y
            print(f"  {ct.Name}: T{_fmt(y)}")
            y += float(ct.Height) + gap
    else:
        ctrls.sort(key=lambda c: float(c.Left))
        if args.gap is not None:
            gap = args.gap
        else:
            spans = [float(ctrls[i + 1].Left) - (float(ctrls[i].Left) + float(ctrls[i].Width))
                     for i in range(len(ctrls) - 1)]
            gap = sum(spans) / len(spans)
        x = float(ctrls[0].Left)
        for ct in ctrls:
            ct.Left = x
            print(f"  {ct.Name}: L{_fmt(x)}")
            x += float(ct.Width) + gap
    print(f"間隔 {_fmt(gap)} で等間隔化しました。")
    finish(wb, args.save)


def cmd_tab_order(args):
    """TabIndex の整列: tab-order <form> [--controls 名,...]

    --controls 指定時はその順に 0,1,2...。省略時は視線順（上→下、左→右。
    行内の上下中央合わせを考慮して上下中央で行を判定）に自動整列。
    TabIndex はコンテナ（フォーム / Frame / Page）ごとの番号なので、
    自動整列はコンテナ単位で行う。
    """
    xl, wb, comp = connect_form(args.form)
    if args.controls:
        ctrls = pick_controls(comp, args.controls)
        for i, ct in enumerate(ctrls):
            try:
                ct.TabIndex = i
            except Exception:
                pass
        order = " → ".join(ct.Name for ct in ctrls[:10]) + ("…" if len(ctrls) > 10 else "")
        print(f"TabIndex を指定順に整列: {order}")
    else:
        groups = {}
        for ct in comp.Designer.Controls:
            try:
                pname = ct.Parent.Name
            except Exception:
                pname = ""
            groups.setdefault(pname, []).append(ct)
        total = 0
        for pname, ctrls in groups.items():
            ctrls.sort(key=lambda c: (round((float(c.Top) + float(c.Height) / 2) / 14),
                                      float(c.Left)))
            for i, ct in enumerate(ctrls):
                try:
                    ct.TabIndex = i
                except Exception:
                    pass
            total += len(ctrls)
        print(f"TabIndex をコンテナ単位の視線順に整列: {len(groups)}コンテナ / {total}コントロール")
    finish(wb, args.save)


def cmd_copy_form(args):
    """フォーム複製: copy-form <form> <newname>

    Export した .frm/.frx のヘッダー（VB_Name / Begin行 / OleObjectBlob 参照）を
    CP932 バイナリのまま書き換えて Import する。レイアウトもコードも丸ごと複製。
    """
    # 旧regex [^\d\W]\w* は先頭 _ を通してしまう（\w−数字＝英字＋_）ため共通ガードに統一
    import vba_manager
    reason = vba_manager.check_vba_identifier(args.new)
    if reason:
        sys.exit(f"'{args.new}' はフォーム名に使えません: {reason}")
    xl, wb, comp = connect_form(args.form)
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == args.new.lower():
            sys.exit(f"'{args.new}' は既に存在します。")

    exp = os.path.join(SCRIPT_DIR, f"_tmpcopy_{comp.Name}.frm")
    exp_frx = os.path.splitext(exp)[0] + ".frx"
    new_frm = os.path.join(SCRIPT_DIR, args.new + ".frm")
    new_frx = os.path.join(SCRIPT_DIR, args.new + ".frx")
    comp.Export(exp)
    try:
        with open(exp, "rb") as f:
            data = f.read()
        old_name = comp.Name.encode("cp932")
        exp_base = os.path.splitext(os.path.basename(exp))[0].encode("cp932")
        new_b = args.new.encode("cp932")
        # ヘッダー行（VB_Name / Begin / OleObjectBlob）だけを書き換え、
        # コード中の同名文字列は触らない（先頭20行に限定）。
        # 行種別ごとに1回だけ置換する（重ねがけすると "F_Range2" の中の
        # "F_Range" が再置換されて "F_Range22" になる）
        lines = data.split(b"\r\n")
        for i, ln in enumerate(lines[:20]):
            if b"OleObjectBlob" in ln:
                lines[i] = ln.replace(exp_base + b".frx", new_b + b".frx")
            elif b"VB_Name" in ln:
                lines[i] = ln.replace(old_name, new_b)
            elif ln.strip().startswith(b"Begin "):
                # Begin {GUID} フォーム名 — 末尾のフォーム名トークンだけ書き換える。
                # 全体 replace だと旧名が16進文字だけの短い名前（"C" 等）のとき
                # GUID 側まで置換されてヘッダーが壊れる
                head, sep, tail = ln.rpartition(b" ")
                if sep and tail == old_name:
                    lines[i] = head + b" " + new_b
                else:
                    lines[i] = ln.replace(old_name, new_b)
        with open(new_frm, "wb") as f:
            f.write(b"\r\n".join(lines))
        if os.path.exists(exp_frx):
            with open(exp_frx, "rb") as f:
                frx = f.read()
            with open(new_frx, "wb") as f:
                f.write(frx)
        wb.VBProject.VBComponents.Import(new_frm)
        print(f"フォーム複製: {comp.Name} → {args.new}（レイアウト+コード）")
        print("  ※ コード内のフォーム自己参照（Me以外で旧名を書いている箇所）があれば手動確認")
    finally:
        for p in (exp, exp_frx, new_frm, new_frx):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
    finish(wb, args.save)


def cmd_delete(args):
    if "," in args.control:
        sys.exit("delete-control は1個ずつ指定してください（カンマ区切り不可）。")
    xl, wb, comp = connect_form(args.form)
    backup_form(wb, comp)
    ct = pick_controls(comp, args.control)[0]
    name = ct.Name
    cap = ""
    try:
        c = ct.Caption
        if c:
            cap = c
    except Exception:
        pass
    print(f"削除対象: {name}  種類={control_type(ct)}  "
          f"L{_fmt(ct.Left)} T{_fmt(ct.Top)} W{_fmt(ct.Width)} H{_fmt(ct.Height)}"
          + (f'  Caption="{cap}"' if cap else ""))

    # 残るイベントプロシージャの検出（削除はしない＝判断はユーザー）
    cm = comp.CodeModule
    pat = re.compile(r"^(?:Private\s+|Public\s+)?Sub\s+(" + re.escape(name) + r"_\w+)",
                     re.IGNORECASE)
    leftovers = []
    for i in range(1, cm.CountOfLines + 1):
        m = pat.match(cm.Lines(i, 1))
        if m:
            leftovers.append(m.group(1))

    comp.Designer.Controls.Remove(name)
    print(f"{name} を削除しました。")
    if leftovers:
        print("警告: 以下のイベントプロシージャがコードに残っています"
              "（削除していません。要否を判断のうえ手動で）:")
        for p in leftovers:
            print("  - " + p)
    finish(wb, args.save)


# ================================================================
# エントリポイント
# ================================================================

def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(
        description="UserForm の機械的な幾何操作 (1回のCOM接続で完結)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("scale", help="フォーム+全コントロールに倍率適用")
    p.add_argument("form", help="フォーム名")
    p.add_argument("factor", type=float, help="倍率 (例: 1.5)")
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_scale)

    p = sub.add_parser("set", help="コントロールの位置/サイズ/Caption/Font.Size を設定")
    p.add_argument("form", help="フォーム名")
    p.add_argument("controls", help="コントロール名 (カンマ区切りで複数可)")
    p.add_argument("--left", type=float)
    p.add_argument("--top", type=float)
    p.add_argument("--width", type=float)
    p.add_argument("--height", type=float)
    p.add_argument("--caption")
    p.add_argument("--font-size", type=float, dest="font_size")
    p.add_argument("--default", action="store_true",
                    help="CommandButton の Default(Enter) を True にする")
    p.add_argument("--cancel", action="store_true",
                    help="CommandButton の Cancel(Esc) を True にする")
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_set)

    p = sub.add_parser("align", help="指定辺を揃える (--top N か --left N)")
    p.add_argument("form", help="フォーム名")
    p.add_argument("controls", help="コントロール名 (カンマ区切りで複数可)")
    p.add_argument("--top", type=float)
    p.add_argument("--left", type=float)
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_align)

    p = sub.add_parser("move", help="相対移動 (--dx/--dy)")
    p.add_argument("form", help="フォーム名")
    p.add_argument("controls", help="コントロール名 (カンマ区切りで複数可)")
    p.add_argument("--dx", type=float, help="横の移動量")
    p.add_argument("--dy", type=float, help="縦の移動量")
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_move)

    p = sub.add_parser("size-match", help="同サイズ化（基準は --ref か先頭）")
    p.add_argument("form", help="フォーム名")
    p.add_argument("controls", help="コントロール名 (カンマ区切り)")
    p.add_argument("--ref", help="基準コントロール名（省略時は先頭）")
    p.add_argument("--width-only", dest="width_only", action="store_true")
    p.add_argument("--height-only", dest="height_only", action="store_true")
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_size_match)

    p = sub.add_parser("distribute", help="等間隔配置 (--vertical|--horizontal)")
    p.add_argument("form", help="フォーム名")
    p.add_argument("controls", help="コントロール名 (カンマ区切り・3個以上推奨)")
    p.add_argument("--vertical", action="store_true", help="縦に等間隔")
    p.add_argument("--horizontal", action="store_true", help="横に等間隔")
    p.add_argument("--gap", type=float, help="間隔（省略時は現在の平均間隔）")
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_distribute)

    p = sub.add_parser("tab-order", help="TabIndex を視線順（or 指定順）に整列")
    p.add_argument("form", help="フォーム名")
    p.add_argument("--controls", help="この順に 0,1,2...（省略時は視線順の自動整列）")
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_tab_order)

    p = sub.add_parser("rename-control",
                       help="コントロール名変更 + イベントプロシージャ宣言行の追随")
    p.add_argument("form", help="フォーム名")
    p.add_argument("old", help="旧コントロール名")
    p.add_argument("new", help="新コントロール名")
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_rename)

    p = sub.add_parser("copy-form", help="フォームを別名で丸ごと複製（レイアウト+コード）")
    p.add_argument("form", help="複製元フォーム名")
    p.add_argument("new", help="新しいフォーム名")
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_copy_form)

    p = sub.add_parser("delete-control",
                       help="コントロール削除 (イベントコードは残す)")
    p.add_argument("form", help="フォーム名")
    p.add_argument("control", help="コントロール名 (1個)")
    p.add_argument("--save", action="store_true", help="実行後にブックを保存")
    p.set_defaults(func=cmd_delete)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
