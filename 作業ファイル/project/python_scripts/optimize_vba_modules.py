"""
shu001.bas / shu003.bas の最適化スクリプト
変更点:
  shu001: Application.Wait削除(2), 空白クリアバグ修正(2), 空白行削除高速化, 行高列幅バグ修正, ProcOfLine最適化(2)
  shu003: ProcOfLine最適化(3), 重複Load削除(2)
"""

import os
import sys
import win32com.client
import pythoncom
import vbam_core  # 実名検証つき Import（名前衝突ガード）を本体と共用する
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def read_bas_cp932(fname):
    with open(os.path.join(BASE_DIR, fname), 'rb') as f:
        return f.read().decode('cp932')


def write_bas_cp932(fname, text):
    with open(os.path.join(BASE_DIR, fname), 'wb') as f:
        f.write(text.encode('cp932'))
    print(f"  -> 保存: {fname}")


def check_and_replace(text, old, new, label):
    count = text.count(old)
    if count == 1:
        text = text.replace(old, new)
        print(f"  ✓ {label}")
    elif count == 0:
        print(f"  ! WARN: {label} -- パターン未検出")
    else:
        print(f"  ! WARN: {label} -- {count}箇所 (想定外の複数マッチ)")
    return text


# ============================================================
# shu001.bas の最適化
# ============================================================
def optimize_shu001(text):
    print("【shu001.bas】")

    # 1. Application.Wait 削除 (2箇所まとめて)
    wait_line = 'Application.Wait [Now()] + 225 / 86400000\r\n'
    cnt = text.count(wait_line)
    if cnt == 2:
        text = text.replace(wait_line, '')
        print(f"  ✓ Application.Wait 削除: 2箇所")
    else:
        print(f"  ! WARN: Application.Wait: {cnt}箇所 (想定2)")

    # 2. 選択範囲の空白じゃない空白をクリア -- On Error と Offset バグ修正
    text = check_and_replace(text,
        'Sub 選択範囲の空白じゃない空白をクリア()\r\n'
        'Application.ScreenUpdating = False\r\n'
        'On Error Resume Next\r\n'
        'For Each c In Selection\r\n'
        'If c = "" Then\r\n'
        '   c.Clear\r\n'
        'End If\r\n'
        'ActiveCell.Offset(1).Select\r\n'
        'Next c\r\n'
        'Application.ScreenUpdating = True\r\n'
        'End Sub\r\n',
        'Sub 選択範囲の空白じゃない空白をクリア()\r\n'
        'Application.ScreenUpdating = False\r\n'
        'For Each c In Selection\r\n'
        '    If c.Value = "" Then c.ClearContents\r\n'
        'Next c\r\n'
        'Application.ScreenUpdating = True\r\n'
        'End Sub\r\n',
        '選択範囲の空白じゃない空白をクリア: On Error削除 / Offsetバグ修正 / ClearContentsに変更'
    )

    # 3. 選択範囲の空白に見えるスペースのみをクリア -- Offset バグ修正
    text = check_and_replace(text,
        'Sub 選択範囲の空白に見えるスペースのみをクリア()\r\n'
        'Application.ScreenUpdating = False\r\n'
        'For Each c In Selection\r\n'
        'If c = " " Or c = "\u3000" Then\r\n'
        '   c.Clear\r\n'
        'End If\r\n'
        'ActiveCell.Offset(1).Select\r\n'
        'Next c\r\n'
        'Application.ScreenUpdating = True\r\n'
        'End Sub\r\n',
        'Sub 選択範囲の空白に見えるスペースのみをクリア()\r\n'
        'Application.ScreenUpdating = False\r\n'
        'For Each c In Selection\r\n'
        '    If c.Value = " " Or c.Value = "\u3000" Then c.ClearContents\r\n'
        'Next c\r\n'
        'Application.ScreenUpdating = True\r\n'
        'End Sub\r\n',
        '選択範囲の空白に見えるスペースのみをクリア: Offsetバグ修正 / ClearContentsに変更'
    )

    # 4. 空白行の削除 -- Union バッチ削除に変更
    text = check_and_replace(text,
        'Sub 空白行の削除()\r\n'
        '  Dim lngLstRow As Long\r\n'
        '  Dim lngLop As Long\r\n'
        "  'If vbNo = MsgBox(\"\u51e6\u7406\u3092\u884c\u3044\u307e\u3059\u3001\u3044\u3044\u3067\u3059\u304b\uff1f\", vbYesNo, \"\u51e6\u7406\u306e\u78ba\u8a8d\") Then Exit Sub\r\n"
        '  lngLstRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).row\r\n'
        '  Application.ScreenUpdating = False\r\n'
        '  For lngLop = lngLstRow To 1 Step -1\r\n'
        '    If Application.WorksheetFunction.CountA(Rows(lngLop)) = 0 Then Rows(lngLop).Delete\r\n'
        '  Next lngLop\r\n'
        '  Application.ScreenUpdating = True\r\n'
        'End Sub\r\n',
        'Sub 空白行の削除()\r\n'
        '  Dim lngLstRow As Long\r\n'
        '  Dim lngLop As Long\r\n'
        '  Dim delRng As Range\r\n'
        '  lngLstRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row\r\n'
        '  Application.ScreenUpdating = False\r\n'
        '  For lngLop = lngLstRow To 1 Step -1\r\n'
        '    If Application.WorksheetFunction.CountA(Rows(lngLop)) = 0 Then\r\n'
        '      If delRng Is Nothing Then Set delRng = Rows(lngLop) Else Set delRng = Union(delRng, Rows(lngLop))\r\n'
        '    End If\r\n'
        '  Next lngLop\r\n'
        '  If Not delRng Is Nothing Then delRng.Delete\r\n'
        '  Application.ScreenUpdating = True\r\n'
        'End Sub\r\n',
        '空白行の削除: Unionバッチ削除で高速化'
    )

    # 5a. 行高と列幅を1枚目のシートに揃える -- RowHeight バグ修正
    text = check_and_replace(text,
        '      .Rows(r).RowHeight = _\r\n'
        '        ActiveSheet.Rows(r).RowHeight\r\n',
        '      .Rows(r).RowHeight = Worksheets(1).Rows(r).RowHeight\r\n',
        '行高と列幅: RowHeight の参照先を Worksheets(1) に修正'
    )

    # 5b. 行高と列幅を1枚目のシートに揃える -- ColumnWidth バグ修正
    text = check_and_replace(text,
        '      .Columns(c).ColumnWidth = _\r\n'
        '        ActiveSheet.Columns(c).ColumnWidth\r\n',
        '      .Columns(c).ColumnWidth = Worksheets(1).Columns(c).ColumnWidth\r\n',
        '行高と列幅: ColumnWidth の参照先を Worksheets(1) に修正'
    )

    # 6. ワークシートを追加してマクロ一覧パーソナル -- ProcOfLine 最適化
    text = check_and_replace(text,
        '      With ブック.VBProject.VBComponents(i).CodeModule\r\n'
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          If proc <> .ProcOfLine(j, 0) Then\r\n'
        '            proc = .ProcOfLine(j, 0)\r\n'
        '            行 = 行 + 1: Cells(行, "A") = "" & proc\r\n'
        '          End If\r\n'
        '        Next j\r\n'
        '      End With\r\n'
        'KOKO:\r\n'
        '    Next i\r\n'
        '  Next\r\n'
        '  Cells.EntireColumn.AutoFit\r\n'
        ' End Sub\r\n',
        '      With ブック.VBProject.VBComponents(i).CodeModule\r\n'
        '        Dim curProc As String\r\n'
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          curProc = .ProcOfLine(j, 0)\r\n'
        '          If proc <> curProc Then\r\n'
        '            proc = curProc\r\n'
        '            If proc <> "" Then 行 = 行 + 1: Cells(行, "A") = proc\r\n'
        '          End If\r\n'
        '        Next j\r\n'
        '      End With\r\n'
        'KOKO:\r\n'
        '    Next i\r\n'
        '  Next\r\n'
        '  Cells.EntireColumn.AutoFit\r\n'
        ' End Sub\r\n',
        'マクロ一覧パーソナル: ProcOfLine 1回呼出しに最適化'
    )

    # 7. ワークシート追加してマクロ一覧アクティブ -- ProcOfLine 最適化
    text = check_and_replace(text,
        '      With ブック.VBProject.VBComponents(i).CodeModule\r\n'
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          If proc <> .ProcOfLine(j, 0) Then\r\n'
        '            proc = .ProcOfLine(j, 0)\r\n'
        '            行 = 行 + 1: Cells(行, "A") = "" & proc\r\n'
        '          End If\r\n'
        '        Next j\r\n'
        '      End With\r\n'
        'KOKO:\r\n'
        '    Next i\r\n'
        'koko2:\r\n'
        '  Next\r\n'
        '  Cells.EntireColumn.AutoFit\r\n'
        'End Sub\r\n',
        '      With ブック.VBProject.VBComponents(i).CodeModule\r\n'
        '        Dim curProc2 As String\r\n'
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          curProc2 = .ProcOfLine(j, 0)\r\n'
        '          If proc <> curProc2 Then\r\n'
        '            proc = curProc2\r\n'
        '            If proc <> "" Then 行 = 行 + 1: Cells(行, "A") = proc\r\n'
        '          End If\r\n'
        '        Next j\r\n'
        '      End With\r\n'
        'KOKO:\r\n'
        '    Next i\r\n'
        'koko2:\r\n'
        '  Next\r\n'
        '  Cells.EntireColumn.AutoFit\r\n'
        'End Sub\r\n',
        'マクロ一覧アクティブ: ProcOfLine 1回呼出しに最適化'
    )

    return text


# ============================================================
# shu003.bas の最適化
# ============================================================
def optimize_shu003(text):
    print("【shu003.bas】")

    # 1. パーソナルのマクロ一覧 -- ProcOfLine 最適化
    text = check_and_replace(text,
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          If proc <> .ProcOfLine(j, 0) Then\r\n'
        '            proc = .ProcOfLine(j, 0)\r\n'
        '            パーソナルマクロフォーム.ListBox1.AddItem proc\r\n'
        '          End If\r\n'
        '        Next j\r\n',
        '        Dim curP1 As String\r\n'
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          curP1 = .ProcOfLine(j, 0)\r\n'
        '          If proc <> curP1 Then\r\n'
        '            proc = curP1\r\n'
        '            If proc <> "" Then パーソナルマクロフォーム.ListBox1.AddItem proc\r\n'
        '          End If\r\n'
        '        Next j\r\n',
        'パーソナルのマクロ一覧: ProcOfLine 1回呼出しに最適化'
    )

    # 2. アドインのマクロ一覧 -- ProcOfLine 最適化
    text = check_and_replace(text,
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          If proc <> .ProcOfLine(j, 0) Then\r\n'
        '            proc = .ProcOfLine(j, 0)\r\n'
        '            マクロフォーム.ListBox1.AddItem proc\r\n'
        '          End If\r\n'
        '        Next j\r\n',
        '        Dim curP2 As String\r\n'
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          curP2 = .ProcOfLine(j, 0)\r\n'
        '          If proc <> curP2 Then\r\n'
        '            proc = curP2\r\n'
        '            If proc <> "" Then マクロフォーム.ListBox1.AddItem proc\r\n'
        '          End If\r\n'
        '        Next j\r\n',
        'アドインのマクロ一覧: ProcOfLine 1回呼出しに最適化'
    )

    # 2b. アドインのマクロ一覧 -- 重複 Load 削除
    text = check_and_replace(text,
        'On Error Resume Next\r\n'
        'Load マクロフォーム\r\n'
        'マクロフォーム.ListBox1.Selected(0) = True\r\n'
        'マクロフォーム.StartUpPosition = 1\r\n'
        'Load マクロフォーム\r\n'
        'マクロフォーム.StartUpPosition = 0\r\n',
        'On Error Resume Next\r\n'
        'Load マクロフォーム\r\n'
        'マクロフォーム.ListBox1.Selected(0) = True\r\n'
        'マクロフォーム.StartUpPosition = 0\r\n',
        'アドインのマクロ一覧: 重複 Load 削除'
    )

    # 3. アクティブのマクロ一覧 -- ProcOfLine 最適化
    text = check_and_replace(text,
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          If proc <> .ProcOfLine(j, 0) Then\r\n'
        '            proc = .ProcOfLine(j, 0)\r\n'
        '            アクティブマクロフォーム.ListBox1.AddItem proc\r\n'
        '          End If\r\n'
        '        Next j\r\n',
        '        Dim curP3 As String\r\n'
        '        proc = ""\r\n'
        '        For j = 1 To .CountOfLines\r\n'
        '          curP3 = .ProcOfLine(j, 0)\r\n'
        '          If proc <> curP3 Then\r\n'
        '            proc = curP3\r\n'
        '            If proc <> "" Then アクティブマクロフォーム.ListBox1.AddItem proc\r\n'
        '          End If\r\n'
        '        Next j\r\n',
        'アクティブのマクロ一覧: ProcOfLine 1回呼出しに最適化'
    )

    # 3b. アクティブのマクロ一覧 -- 重複 Load 削除
    text = check_and_replace(text,
        'On Error Resume Next\r\n'
        'Load アクティブマクロフォーム\r\n'
        'アクティブマクロフォーム.ListBox1.Selected(0) = True\r\n'
        'アクティブマクロフォーム.StartUpPosition = 1\r\n'
        'Load アクティブマクロフォーム\r\n'
        'アクティブマクロフォーム.StartUpPosition = 0\r\n',
        'On Error Resume Next\r\n'
        'Load アクティブマクロフォーム\r\n'
        'アクティブマクロフォーム.ListBox1.Selected(0) = True\r\n'
        'アクティブマクロフォーム.StartUpPosition = 0\r\n',
        'アクティブのマクロ一覧: 重複 Load 削除'
    )

    return text


def apply_module(wb, module_name, bas_fname):
    """既存モジュールを Remove + Import で差し替え（Attributeを正しく処理）

    素の Remove→Import は、VBE の Remove が遅延完了する隙に同名を取り込んで
    「shu0031」のような連番別名になり、しかもエラーにならず成功扱いになる
    （2026-07-11 の shu005 消滅事故と同型）。vba_manager 本体と同じ
    _import_module_verified（実名検証つき Import）を通す。
    """
    bas_path = os.path.join(BASE_DIR, bas_fname)
    if not os.path.exists(bas_path):
        print(f"  ! エラー: {bas_fname} がありません")
        return False

    comp = vbam_core._find_component(wb, module_name)
    if comp is None:
        print(f"  ! エラー: モジュール '{module_name}' が見つかりません")
        return False

    # 差し替え前の内容を退避（Import 失敗時の復旧素材）
    module_backup = vbam_core.make_module_backup(wb, module_name)

    removed = False
    imported = False
    try:
        wb.VBProject.VBComponents.Remove(comp)
        removed = True
        pythoncom.PumpWaitingMessages()
        vbam_core._import_module_verified(wb, bas_path, module_name)
        imported = True
        print(f"  ✓ モジュール '{module_name}' を Remove+Import で更新完了（実名検証済み）")
        return True
    except vbam_core.ModuleNameCollisionError as ex:
        vbam_core._print_collision_guidance(ex, module_name, module_backup)
        return False
    except Exception as ex:
        print(f"  ! エラー: '{module_name}' の差し替えに失敗しました: {ex}")
        if removed and not imported:
            print(f"  ⚠ モジュール '{module_name}' は Remove 済みです。バックアップから復旧を試みます...")
            try:
                if module_backup and os.path.exists(module_backup):
                    vbam_core._import_module_verified(wb, module_backup, module_name)
                    print(f"  復旧成功: {module_backup} を再インポートしました（差し替え前の内容に戻っています）")
                else:
                    raise RuntimeError("モジュールバックアップがありません")
            except Exception as ex2:
                print(f"  復旧失敗: {ex2}")
                print("  ⚠ このままブックを保存するとモジュールがファイルからも消えます。")
                print("  対処: ブックを『保存せずに閉じて』開き直してください。")
        return False


def main():
    # --- shu001 最適化 ---
    shu001_orig = read_bas_cp932('shu001.bas')
    shu001_opt = optimize_shu001(shu001_orig)
    write_bas_cp932('shu001_optimized.bas', shu001_opt)

    # --- shu003 最適化 ---
    shu003_orig = read_bas_cp932('shu003.bas')
    shu003_opt = optimize_shu003(shu003_orig)
    write_bas_cp932('shu003_optimized.bas', shu003_opt)

    # --- Excel に適用 ---
    print("\n【Excel に適用】")
    pythoncom.CoInitialize()
    # 起動中の Excel のアクティブブックに作業（特定ブックに限定しない＝汎用）
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        print("  ! Excel が起動していません。対象ブックを開いてから実行してください。")
        return 1

    try:
        wb = xl.ActiveWorkbook
        if wb is None:
            print("  ! アクティブなブックがありません。対象ブックを開いてから実行してください。")
            return 1
        print(f"  対象ブック: {wb.Name}")

        ok1 = apply_module(wb, 'shu001', 'shu001_optimized.bas')
        ok3 = apply_module(wb, 'shu003', 'shu003_optimized.bas')

        if ok1 and ok3:
            vbam_core._save_with_retry(wb)
            print(f"  ✓ {wb.Name} 保存完了")
        else:
            print("  ! エラーがあるため保存しませんでした")
            return 1

    except Exception as e:
        print(f"  ! エラー: {e}")
        import traceback
        traceback.print_exc()
        return 1

    print("\n完了")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
