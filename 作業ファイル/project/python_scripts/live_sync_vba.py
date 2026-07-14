"""秀.xlsm の shu001 / shu003 を、開いているブックの上で直接手直しするスクリプト。

  1. 定型の説明コメントブロック（'==== / 処理名: / 概要: / ====）を削除
  2. shu003 に「アドインの登録解除」マクロが無ければ追記

CodeModule のテキストを直接書き戻す（Remove+Import を使わない）ので、
Attribute 行やショートカット登録は壊れない。

安全上の約束（2026-07-14 に修正）:
  - DeleteLines の前に必ず旧コードを控え、AddFromString が失敗したら書き戻す
    （空のまま保存してモジュールを失う事故を防ぐ）
  - 1つでも失敗したら保存しない・成功と報告しない・非ゼロ終了する
  - Excel が起動していなければ新インスタンスを作らずに終わる
    （作っても対象ブックは開いておらず、ゾンビ Excel が残るだけ）
"""

import sys
import win32com.client
import pythoncom
import re

TARGET_NAME = "秀.xlsm"

# 削除パターン（説明文）
DESC_PATTERN = re.compile(
    r"'\s*={10,}\s*\n"
    r"'\s*処理名:.*\n"
    r"'\s*概要\s*:.*\n"
    r"'\s*={10,}\s*\n?",
    re.MULTILINE
)

# 追加マクロ（アドインの登録解除）
NEW_MACRO = """
Sub アドインの登録解除()
    Dim addinPath As String
    ' アドインファイルのパスを設定
    addinPath = Environ("AppData") & "\\Microsoft\\AddIns\\秀.xlam"

    On Error Resume Next
    ' アドインを無効化
    AddIns("秀").Installed = False

    ' ファイル実体が存在する場合は削除（Kill命令）
    If Dir(addinPath) <> "" Then
        Kill addinPath
        If Err.Number = 0 Then
            MsgBox "アドイン「秀」の登録解除とファイルの削除が完了しました。", vbInformation
        Else
            MsgBox "登録は解除されましたが、ファイルの削除に失敗しました（使用中の可能性があります）。", vbExclamation
        End If
    Else
        MsgBox "アドイン「秀」の解除処理を行いました（ファイルは見つかりませんでした）。", vbInformation
    End If
    On Error GoTo 0
End Sub
"""


def _normalize_crlf(text):
    """改行を CRLF に揃える（\\r\\r\\n 化＝行の二重膨張を防ぐ）"""
    return re.sub(r'\r+\n?', '\n', text).replace('\n', '\r\n')


def update_module(vbp, mod_name):
    """1モジュールを書き換える。成功したら True。

    DeleteLines → AddFromString の間にモジュールは一瞬「空」になる。
    ここで例外が飛ぶと空のまま残るため、必ず旧コードを控えて書き戻す。
    """
    comp = vbp.VBComponents(mod_name)
    cm = comp.CodeModule

    old_code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""

    new_code = DESC_PATTERN.sub("", old_code)

    if mod_name == "shu003" and "Sub アドインの登録解除()" not in new_code:
        new_code = new_code.rstrip() + "\n" + NEW_MACRO

    if new_code.strip() == old_code.strip():
        print(f"  = {mod_name}: 変更はありません（説明文なし・マクロ追記済み）")
        return True

    new_code = _normalize_crlf(new_code)

    cm.DeleteLines(1, cm.CountOfLines)
    try:
        cm.AddFromString(new_code)
    except Exception:
        # 空のモジュールを残さない。旧コードを書き戻してから失敗を伝える
        try:
            cm.DeleteLines(1, cm.CountOfLines)
        except Exception:
            pass
        cm.AddFromString(_normalize_crlf(old_code))
        print(f"  ! {mod_name}: 書き込みに失敗したため、元のコードに戻しました")
        raise

    print(f"  ✓ {mod_name}: 説明文を削除（必要ならマクロも追記）")
    return True


def update_excel_live():
    pythoncom.CoInitialize()
    try:
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            # 新インスタンスは作らない。作っても対象ブックは開いておらず、
            # 「見つからない」で終わってゾンビ Excel が残るだけ
            print(f"エラー: Excel が起動していません。{TARGET_NAME} を開いてから実行してください。")
            return 1

        wb = None
        for w in xl.Workbooks:
            if w.Name == TARGET_NAME:
                wb = w
                break

        if wb is None:
            print(f"エラー: {TARGET_NAME} が開かれていません。")
            return 1

        print(f"対象ブック: {wb.Name}")
        vbp = wb.VBProject

        failed = []
        for mod_name in ["shu001", "shu003"]:
            try:
                update_module(vbp, mod_name)
            except Exception as e:
                print(f"  ! {mod_name} の更新に失敗しました: {e}")
                failed.append(mod_name)

        if failed:
            # 1つでもコケたら保存しない（中途半端な状態をファイルに焼き付けない）
            print(f"エラー: {', '.join(failed)} の更新に失敗したため、保存しませんでした。")
            print("  開いているブックは元の内容に戻してあります。"
                  "保存せずに閉じれば確実に元通りです。")
            return 1

        wb.Save()
        print("保存しました（shu001 / shu003 とも更新済み）。")
        return 0

    except Exception as e:
        print(f"エラー: {e}")
        return 1
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(update_excel_live())
