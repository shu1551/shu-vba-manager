import win32com.client
import pythoncom
import re

def update_excel_live():
    pythoncom.CoInitialize()
    xl = None
    try:
        # 起動中のExcelを取得
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
            print("Successfully connected to active Excel instance.")
        except:
            print("Active Excel instance not found. Opening file...")
            xl = win32com.client.Dispatch("Excel.Application")
            xl.Visible = True
        
        target_name = "秀.xlsm"
        wb = None
        for w in xl.Workbooks:
            if w.Name == target_name:
                wb = w
                break
        
        if not wb:
            print(f"Error: {target_name} is not open.")
            return

        vbp = wb.VBProject
        
        # 削除パターン（説明文）
        pattern = re.compile(
            r"'\s*={10,}\s*\n"
            r"'\s*処理名:.*\n"
            r"'\s*概要\s*:.*\n"
            r"'\s*={10,}\s*\n?",
            re.MULTILINE
        )

        # 追加マクロ（アドインの登録解除）
        new_macro = """
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

        for mod_name in ["shu001", "shu003"]:
            try:
                comp = vbp.VBComponents(mod_name)
                cm = comp.CodeModule
                
                # 現在のコードを全取得
                code = cm.Lines(1, cm.CountOfLines)
                
                # 1. 説明文を削除
                new_code = pattern.sub("", code)
                
                # 2. shu003 の場合はマクロ追記（まだなければ）
                if mod_name == "shu003" and "Sub アドインの登録解除()" not in new_code:
                    # 「アドインの更新登録」の後ろ、または末尾に追記
                    if "Sub アドインの更新登録()" in new_code:
                        # 本来は正規表現でSub終了位置を探すべきだが、今回は末尾追記で安全を期す
                        new_code = new_code.rstrip() + "\n" + new_macro
                    else:
                        new_code = new_code.rstrip() + "\n" + new_macro
                
                # 3. メモリ上のテキストをVBAにそのまま書き戻す（これが文字化けしない最強の方法）
                cm.DeleteLines(1, cm.CountOfLines)
                cm.AddFromString(new_code)
                print(f"Updated {mod_name} (Removed descriptions, added macro if needed)")
                
            except Exception as e:
                print(f"Error updating {mod_name}: {e}")

        # 保存
        wb.Save()
        print("Successfully saved changes to active workbook.")
        
    except Exception as e:
        print(f"Critical error: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    update_excel_live()
