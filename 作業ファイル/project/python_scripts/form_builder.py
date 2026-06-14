r"""
form_builder.py - UserForm作成のための共通ユーティリティ

【使い方】
    from form_builder import FormBuilder, add_btn, add_lbl, add_txt, add_lst, add_combo

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


def add_frame(frm, name, caption, left, top, width, height):
    """Frame を追加する"""
    f2 = frm.Controls.Add("Forms.Frame.1", name)
    f2.Caption = caption
    f2.Left = left; f2.Top = top; f2.Width = width; f2.Height = height
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
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            xl = win32com.client.Dispatch("Excel.Application")
            xl.Visible = True

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
                except Exception:
                    print(f"警告: リネーム失敗 ({frm_comp.Name} として続行)")

        if caption is not None:
            frm_comp.Properties("Caption").Value = caption
        if width is not None:
            frm_comp.Properties("Width").Value = width
        if height is not None:
            frm_comp.Properties("Height").Value = height

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

    def inject_vba(self, frm_comp, vba_file):
        """
        .vba ファイル (UTF-8) からコードを読み込み CodeModule に注入する。
        既存コードは全削除してから上書き。
        """
        abs_vba = os.path.abspath(vba_file)
        with open(abs_vba, "r", encoding="utf-8") as fp:
            code = fp.read()
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
