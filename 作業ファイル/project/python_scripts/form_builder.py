r"""
form_builder.py - UserForm作成のための共通ユーティリティ

【使い方】
    from form_builder import FormBuilder, add_btn, add_lbl, add_txt, add_lst, add_combo
    from form_builder import Grid, vstack, hstack

【FormBuilder の基本フロー】
    with FormBuilder.connect(wb_path="path/to/file.xlsm") as fb:
        frm = fb.get_or_create("MyForm", caption="タイトル", width=300, height=200)
        f = fb.clear_controls(frm)
        add_btn(f, "BtnOK", "OK", 100, 160, 60, 20)
        fb.inject_vba(frm, "my_form_code.vba")
        fb.save()

【接続方法】
    FormBuilder.connect(wb_path=r"C:\...\file.xlsm")   # パス指定
    FormBuilder.connect(wb_keyword="ポスター")           # キーワード検索
    FormBuilder.connect()                                # アクティブブック
"""

import os
import re
import pythoncom
import win32com.client

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


# ================================================================
# コントロール追加ヘルパー
# ================================================================

def add_btn(frm, name, caption, left, top, width, height, *, font_size=12, font_bold=False):
    """CommandButton を追加する"""
    b = frm.Controls.Add("Forms.CommandButton.1", name)
    b.Caption = caption
    b.Left = left; b.Top = top; b.Width = width; b.Height = height
    b.Font.Size = font_size
    if font_bold:
        b.Font.Bold = True
    return b


def add_lbl(frm, name, caption, left, top, width, height, *,
            align=1, fore=None, back=None, font_size=12, font_bold=False):
    """Label を追加する (align: 1=左, 2=中央, 3=右)"""
    lbl = frm.Controls.Add("Forms.Label.1", name)
    lbl.Caption = caption
    lbl.Left = left; lbl.Top = top; lbl.Width = width; lbl.Height = height
    lbl.TextAlign = align
    if fore is not None:
        lbl.ForeColor = fore
    if back is not None:
        lbl.BackColor = back
    lbl.Font.Size = font_size
    if font_bold:
        lbl.Font.Bold = True
    return lbl


def add_txt(frm, name, left, top, width, height, *,
            value="", multiline=False, scrollbars=0, font_size=12):
    """TextBox を追加する"""
    t = frm.Controls.Add("Forms.TextBox.1", name)
    t.Left = left; t.Top = top; t.Width = width; t.Height = height
    if value:
        t.Value = value
    if multiline:
        t.MultiLine = True
        t.ScrollBars = scrollbars  # 0=なし, 1=水平, 2=垂直, 3=両方
    t.Font.Size = font_size
    return t


def add_lst(frm, name, left, top, width, height, *, font_size=12):
    """ListBox を追加する"""
    lb = frm.Controls.Add("Forms.ListBox.1", name)
    lb.Left = left; lb.Top = top; lb.Width = width; lb.Height = height
    lb.Font.Size = font_size
    return lb


def add_combo(frm, name, left, top, width, height, *, font_size=12):
    """ComboBox を追加する"""
    cb = frm.Controls.Add("Forms.ComboBox.1", name)
    cb.Left = left; cb.Top = top; cb.Width = width; cb.Height = height
    cb.Font.Size = font_size
    return cb


def add_chk(frm, name, caption, left, top, width, height, *, font_size=12):
    """CheckBox を追加する"""
    c = frm.Controls.Add("Forms.CheckBox.1", name)
    c.Caption = caption
    c.Left = left; c.Top = top; c.Width = width; c.Height = height
    c.Font.Size = font_size
    return c


def add_opt(frm, name, caption, left, top, width, height, *, font_size=12):
    """OptionButton を追加する"""
    o = frm.Controls.Add("Forms.OptionButton.1", name)
    o.Caption = caption
    o.Left = left; o.Top = top; o.Width = width; o.Height = height
    o.Font.Size = font_size
    return o


def add_frame(frm, name, caption, left, top, width, height, *, font_size=12):
    """Frame を追加する"""
    f2 = frm.Controls.Add("Forms.Frame.1", name)
    f2.Caption = caption
    f2.Left = left; f2.Top = top; f2.Width = width; f2.Height = height
    f2.Font.Size = font_size   # 他コントロールと同じく既定12pt（Caption に効く）
    return f2


def add_img(frm, name, left, top, width, height):
    """Image を追加する"""
    img = frm.Controls.Add("Forms.Image.1", name)
    img.Left = left; img.Top = top; img.Width = width; img.Height = height
    return img


def add_spin(frm, name, left, top, width, height):
    """SpinButton を追加する"""
    sp = frm.Controls.Add("Forms.SpinButton.1", name)
    sp.Left = left; sp.Top = top; sp.Width = width; sp.Height = height
    return sp


def add_scroll(frm, name, left, top, width, height, *, orientation=0):
    """ScrollBar を追加する (orientation: 0=水平, 1=垂直)"""
    sb = frm.Controls.Add("Forms.ScrollBar.1", name)
    sb.Left = left; sb.Top = top; sb.Width = width; sb.Height = height
    sb.Orientation = orientation
    return sb


# ================================================================
# 配置ヘルパー（機械的な座標計算のみ）
# ================================================================

class Grid:
    """グリッド配置の座標計算。

    g = Grid(left0, top0, col_w, row_h, gap_x=6, gap_y=6)
    g.pos(row, col) -> (left, top)

    例: カレンダーの 7列×6行 ボタン群
        g = Grid(10, 40, 24, 20, gap_x=2, gap_y=2)
        for i in range(42):
            left, top = g.pos(i // 7, i % 7)
            add_btn(f, f"btnDay{i}", "", left, top, 24, 20)
    """

    def __init__(self, left0, top0, col_w, row_h, gap_x=6, gap_y=6):
        self.left0 = left0
        self.top0 = top0
        self.col_w = col_w
        self.row_h = row_h
        self.gap_x = gap_x
        self.gap_y = gap_y

    def pos(self, row, col):
        """(row, col) セルの (left, top) を返す"""
        return (self.left0 + col * (self.col_w + self.gap_x),
                self.top0 + row * (self.row_h + self.gap_y))


def vstack(prev, gap=6):
    """直前コントロールの下端 + gap の top を返す"""
    return prev.Top + prev.Height + gap


def hstack(prev, gap=6):
    """直前コントロールの右端 + gap の left を返す"""
    return prev.Left + prev.Width + gap


# ================================================================
# FormBuilder クラス
# ================================================================

class FormBuilder:
    r"""
    UserForm の作成・編集を管理するクラス。

    with 文で使用すると CoInitialize/CoUninitialize を自動管理する。

    例:
        with FormBuilder.connect(wb_path=r"C:\...\秀.xlsm") as fb:
            frm = fb.get_or_create("MyForm", caption="My Form", width=300, height=200)
            f = fb.clear_controls(frm)
            add_btn(f, "BtnOK", "OK", 100, 160, 60, 20)
            fb.inject_vba(frm, "my_code.vba")
            fb.save()
    """

    def __init__(self, xl, wb):
        self._xl = xl
        self._wb = wb

    # ---- ファクトリ ----

    @classmethod
    def connect(cls, *, wb_path=None, wb_keyword=None):
        """
        Excelに接続し FormBuilder を返す。

        Args:
            wb_path (str):    ブックのフルパス (省略可)
            wb_keyword (str): ブック名に含まれるキーワード (省略可)
                              両方省略するとアクティブブックを使用。
        """
        pythoncom.CoInitialize()
        try:
            try:
                # gen_py キャッシュ破損の影響を受けない生の GetActiveObject
                # （vba_manager.py の _get_active_excel と同方式）
                unk = pythoncom.GetActiveObject("Excel.Application")
                disp = unk.QueryInterface(pythoncom.IID_IDispatch)
                xl = win32com.client.dynamic.Dispatch(disp)
            except Exception:
                if not wb_path:
                    # 無警告で新規 Excel を起動すると、アドイン未ロードのインスタンスや
                    # 空 Excel の残骸（ゾンビ）を生む。パス指定がない限り起動しない。
                    raise RuntimeError(
                        "起動中の Excel が見つかりません。"
                        "対象ブックを Excel で開いてから実行してください。")
                # DispatchEx＝必ず別インスタンス（Dispatch は既存に接続してしまい
                # ユーザーの Excel を巻き込む事故のもと。2026-07-03 に vba_manager 側で実害）
                xl = win32com.client.DispatchEx("Excel.Application")
                xl.Visible = True
                print("注意: Excelが未起動だったため自動化用に新規起動しました。")
                print("      このExcelにはアドイン・PERSONAL.XLSB が読み込まれていません。")
                print("      作業後はこのExcelを閉じてください。")

            if wb_path:
                abs_path = os.path.abspath(wb_path).lower()
                for w in xl.Workbooks:
                    if w.FullName.lower() == abs_path:
                        return cls(xl, w)
                wb = xl.Workbooks.Open(os.path.abspath(wb_path))
                return cls(xl, wb)

            if wb_keyword:
                for w in xl.Workbooks:
                    if wb_keyword in w.Name:
                        return cls(xl, w)
                raise RuntimeError(f"ブック '{wb_keyword}' が見つかりません。Excelで開いてください。")

            # アクティブブック
            wb = xl.ActiveWorkbook
            if wb is None:
                raise RuntimeError("アクティブなブックがありません。")
            return cls(xl, wb)
        except Exception:
            # connect 失敗時は __exit__ に到達しないため、ここで対にして解放する
            pythoncom.CoUninitialize()
            raise

    # ---- コンテキストマネージャ ----

    def __enter__(self):
        return self

    def __exit__(self, *_):
        pythoncom.CoUninitialize()

    # ---- プロパティ ----

    @property
    def workbook(self):
        return self._wb

    @property
    def vbproject(self):
        return self._wb.VBProject

    # ---- フォーム操作 ----

    def _close_vbe_designer_windows(self):
        """VBEのデザインウィンドウを全て閉じる。
        UserFormのリネーム前に呼ぶことで、旧フォーム名への参照ロックを解除する。
        """
        try:
            vbe = self.vbproject.VBE
            # Windows.Count は動的に変化するので逆順で閉じる
            for i in range(vbe.Windows.Count, 0, -1):
                win = vbe.Windows.Item(i)
                # Type 0 = vbext_wt_Designer（フォームのデザインウィンドウ）
                if win.Type == 0:
                    win.Close()
        except Exception as e:
            print(f"VBEウィンドウクローズ中に例外 (無視): {e}")

    def get_or_create(self, form_name, *, caption=None, width=None, height=None):
        """
        指定名のフォームを探して返す。なければ新規作成する。
        caption / width / height を指定するとプロパティも設定する。
        """
        frm_comp = None
        for comp in self.vbproject.VBComponents:
            if comp.Name == form_name:
                frm_comp = comp
                print(f"既存 {form_name} を再利用")
                break

        if frm_comp is None:
            # リネーム前にVBEデザインウィンドウを閉じて参照ロックを解除
            self._close_vbe_designer_windows()
            frm_comp = self.vbproject.VBComponents.Add(3)  # vbext_ct_MSForm
            try:
                frm_comp.Name = form_name
                print(f"新規 {form_name} を作成")
            except Exception:
                # 1回目で失敗したら少し待ってリトライ
                import time
                time.sleep(0.5)
                self._close_vbe_designer_windows()
                try:
                    frm_comp.Name = form_name
                    print(f"新規 {form_name} を作成 (リトライ成功)")
                except Exception as ex:
                    # 既定名（UserForm1等）のまま続行すると、以後の get_or_create が
                    # 毎回「見つからない→新規作成」で同名フォームを量産するため停止する
                    default_name = frm_comp.Name
                    try:
                        self.vbproject.VBComponents.Remove(frm_comp)
                    except Exception:
                        pass
                    raise RuntimeError(
                        f"フォームのリネームに失敗しました: {default_name} → {form_name} ({ex})\n"
                        "  VBE のデザイナーウィンドウが開いているとリネームできないことがあります。"
                        "Excel/VBE を確認して再実行してください。")

        # 直前に Show/Unload されたフォーム等では Properties が一時的に COM エラーを
        # 返すことがある。Designer に触れて実体化→少し待って再試行、で安定する。
        def _set_props():
            if caption is not None:
                frm_comp.Properties("Caption").Value = caption
            if width is not None:
                frm_comp.Properties("Width").Value = width
            if height is not None:
                frm_comp.Properties("Height").Value = height
        try:
            _set_props()
        except Exception:
            import time
            try:
                _ = frm_comp.Designer.Controls.Count   # Designer を実体化
            except Exception:
                pass
            time.sleep(0.5)
            self._close_vbe_designer_windows()
            _set_props()                               # 2回目も失敗なら例外を上げる

        return frm_comp

    def delete_form(self, form_name):
        """指定名のフォームを削除する (存在しない場合は何もしない)"""
        for comp in self.vbproject.VBComponents:
            if comp.Name == form_name:
                # 削除前にデザインウィンドウを閉じて参照を解放
                self._close_vbe_designer_windows()
                self.vbproject.VBComponents.Remove(comp)
                print(f"{form_name} を削除しました")
                return
        print(f"{form_name} は存在しません")

    def clear_controls(self, frm_comp):
        """フォームの既存コントロールをすべて削除し、Designer オブジェクトを返す"""
        f = frm_comp.Designer
        names = [c.Name for c in f.Controls]
        for n in names:
            try:
                f.Controls.Remove(n)
            except Exception:
                pass
        print(f"コントロールをクリア ({len(names)} 個削除)")
        return f

    def inject_vba(self, frm_comp, vba_file=None):
        """
        .vba ファイル (UTF-8) からコードを読み込み CodeModule に注入する。
        既存コードは全削除してから上書き。
        vba_file 省略時はスクリプトと同じ場所の _last_form_code.vba を使う。
        """
        if vba_file is None:
            vba_file = os.path.join(SCRIPT_DIR, "_last_form_code.vba")
        abs_vba = os.path.abspath(vba_file)
        with open(abs_vba, "r", encoding="utf-8") as fp:
            code = fp.read()
        # 識別子ガード（注入点側の多層防御）: 先頭 _ 等の無効なプロシージャ名は
        # AddFromString が黙って受け入れてコンパイルで死ぬ。既存コードを消す前に止める。
        import vba_manager
        bad = vba_manager._find_invalid_procedure_names(re.sub(r'\r\n|\r', '\n', code))
        if bad:
            for ln, _nm, reason in bad:
                print(f"エラー: 行{ln}: プロシージャ名が VBA の識別子規則に反しています: {reason}")
            raise ValueError("VBA 識別子規則違反のため注入を中止しました（フォームの既存コードは残っています）")
        cm = frm_comp.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        # 改行を正規CRLFに揃える（改行二重化ガード／AddFromString直前の防御統一）
        code = re.sub(r'\r+\n?', '\n', code).replace('\n', '\r\n')
        cm.AddFromString(code)
        print(f"VBAコード注入完了 ({os.path.basename(vba_file)})")

    def save(self):
        """ブックを上書き保存する"""
        self._wb.Save()
        print(f"保存完了: {self._wb.Name}")

    def list_forms(self):
        """ブック内のフォーム名一覧を返す"""
        return [c.Name for c in self.vbproject.VBComponents if c.Type == 3]
