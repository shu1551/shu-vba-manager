"""実機の Excel / VBE を使う E2E テスト（オプトイン）。

  走らせ方:  py -m pytest test_e2e_com.py --run-e2e -q
  既定では   py -m pytest -q          ← 純ロジックだけ（1〜2秒）でスキップされる

■ なぜ正式ファイルなのか（2026-07-14）
E2E がセッションごとの使い捨てスクリプトにしか無く、毎回書き直されていた。
その結果、同じ穴を何度も踏み直し「テストがまた落ちる」を繰り返していた。
ここに置いてある限り、書き直しは起きない。

■ 安全上の約束（過去の事故から）
  - シュウさんの Excel には絶対に触らない。既に Excel が起動していたら skip する
    （理由を名指しして止める。曖昧に落ちない）
  - ブックは毎回ユニークな一時パスに作る（固定パスは前回の残骸と衝突する）
  - _last_proc.vba は共有ファイル。退避して、終わったら必ず戻す
  - Excel は DispatchEx で起こし、必ず畳む（ゾンビを残さない）
"""
import os
import subprocess
import sys
import time

import pytest

pytestmark = pytest.mark.e2e

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VBM = os.path.join(SCRIPT_DIR, "vba_manager.py")
LAST_PROC = os.path.join(SCRIPT_DIR, "_last_proc.vba")

MOD_NAME = "TestMod"

MOD_CODE = (
    'Sub 一行完結(): Debug.Print 1: End Sub\r\n'
    'Sub 合計を出す()\r\n'
    '    Range("C1").Value = Range("A1").Value + Range("B1").Value\r\n'
    'End Sub\r\n'
)

# 閉じたブックを覗いただけで Workbook_Open が走ったら、その痕跡がセルに残る
THISWB_CODE = (
    'Private Sub Workbook_Open()\r\n'
    '    Worksheets("本体").Range("Z1").Value = "OPENED"\r\n'
    'End Sub\r\n'
)


# ================================================================
# 土台
# ================================================================

def _excel_is_running():
    """EXCEL.EXE が動いているか（シュウさんの作業中 Excel を巻き込まないための門番）"""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
            capture_output=True, encoding="cp932", errors="replace", timeout=30)
        return "EXCEL.EXE" in (out.stdout or "")
    except Exception:
        return False


def _wait_no_excel(limit=90.0):
    """EXCEL.EXE が完全に消えるまで待つ。消えたら True。

    Excel の死は非同期で、Quit を撃ってから実測 2〜3 秒（負荷次第でもっと）かかる。
    死にきる前に次の操作を始めると、vba_manager は GetActiveObject で
    その「まだ死んでいない Excel」に相乗りしてしまい、挙動が変わる。
    これがテストが実行ごとにコケたり通ったりする正体だった（2026-07-14 特定）。
    待ち時間を延ばして誤魔化すのではなく、毎回「無人」を確認してから次へ進む。
    """
    t0 = time.time()
    while time.time() - t0 < limit:
        if not _excel_is_running():
            return True
        time.sleep(0.25)
    return False


@pytest.fixture(scope="session", autouse=True)
def guard_users_excel():
    """作業中の Excel があるなら、理由を言って止まる（曖昧に落ちない）

    見えている Excel を閉じてもこれが出る場合は、非表示・ブック0 の残骸 Excel が
    居座っている（過去に「アドインが効かない」を起こした型）。
    list-open / タスクマネージャで確認して始末してから再実行する。
    """
    if _excel_is_running():
        pytest.skip(
            "Excel が起動しています。開いているブックを巻き込まないため E2E は行いません。\n"
            "  Excel を閉じてから --run-e2e してください。\n"
            "  閉じたはずなのに出る場合は、非表示の残骸 Excel が居ます"
            "（タスクマネージャの EXCEL.EXE を確認）。")
    yield


@pytest.fixture(scope="session", autouse=True)
def preserve_last_proc():
    """_last_proc.vba は MCP サーバー・GUI と共有の受け渡しファイル。退避して必ず戻す"""
    saved = None
    if os.path.exists(LAST_PROC):
        with open(LAST_PROC, "rb") as f:
            saved = f.read()
    yield
    if saved is not None:
        with open(LAST_PROC, "wb") as f:
            f.write(saved)
    elif os.path.exists(LAST_PROC):
        os.remove(LAST_PROC)


@pytest.fixture(scope="session", autouse=True)
def com_apartment():
    """COM の初期化はセッションで1回だけ。CoUninitialize は「しない」。

    2026-07-14 に実測して分かった2つの罠:
      1. CoInitialize / CoUninitialize はスレッドごとの参照カウント。helper のたびに
         対で呼ぶと vbam_core.cleanup_excel() が内部で呼ぶ CoUninitialize と釣り合わず、
         生きた COM 参照を抱えたまま COM を落として**プロセスごと即死**する。
      2. 終了時に自分で CoUninitialize しても、pywin32 側に残った参照との兼ね合いで
         インタプリタ終了時に落ち、全テストが通っているのに**終了コードが 1 になる**。
    COM アパートメントの後始末はプロセス終了時に OS に任せるのが安全。
    """
    import pythoncom
    pythoncom.CoInitialize()
    yield
    # あえて CoUninitialize しない（上記2）。ゾンビ Excel は no_zombie_excel が見張る。


@pytest.fixture(autouse=True)
def no_zombie_excel():
    """各テストの入口で無人を再確認し、出口で Excel を残していないことを確かめる。

    無人確認がセッション開始の1回だけだと、E2E 実行中にシュウさんが Excel を
    開いた場合、以後のテストがその Excel に GetActiveObject で相乗りし、
    さらにこの見張りが本人の Excel を「ゾンビ」と誤認して残り全テスト×最大90秒の
    空回り＋誤失格になる。入口で検知したら巻き込まずに skip する。
    """
    if _excel_is_running():
        pytest.skip("Excel が起動しています（E2E 実行中に開かれた可能性）。"
                    "作業中の Excel を巻き込まないため、このテストは行いません。")
    yield
    if not _wait_no_excel():
        pytest.fail("テスト後に EXCEL.EXE が残り続けています（ゾンビ Excel か、"
                    "テスト中に開かれた作業用 Excel）。後始末を確認してください。")


def _new_excel():
    """新インスタンス（DispatchEx）。COM の初期化はセッション側で済んでいる"""
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    return xl


# 後始末の作法（2026-07-14 に実測して分かったこと）:
#   xl.Quit() を撃っても、Python 側が COM 参照を1つでも握っている限り EXCEL.EXE は
#   死なない（プロセス終了時にようやく消える＝テスト中はゾンビに見える）。
#   ヘルパ関数に xl を渡して中で gc しても、呼び出し元の変数がまだ生きているので効かない。
#   参照を持っている当人が None にしてから gc.collect() すること。実測で 2.4 秒で消える。
#   そのため各所で quit → None → gc を「その場で」書く（関数に切り出さない）。


@pytest.fixture
def book(tmp_path):
    """使い捨ての .xlsm を1つ作って渡す（毎回ユニークなパス＝残骸と衝突しない）"""
    import gc
    path = str(tmp_path / "vbam_e2e.xlsm")
    xl = _new_excel()
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "本体"
        ws.Range("A1").Value = 2
        ws.Range("B1").Value = 3
        comp = wb.VBProject.VBComponents.Add(1)   # 1 = 標準モジュール
        comp.Name = MOD_NAME
        comp.CodeModule.AddFromString(MOD_CODE)
        wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString(THISWB_CODE)
        wb.SaveAs(path, FileFormat=52)            # 52 = xlOpenXMLWorkbookMacroEnabled
        wb.Close(SaveChanges=False)
        comp = ws = wb = None
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
    # 完全に消えるまで待ってから本番へ。生き残った Excel が居ると、
    # vba_manager がそいつに相乗りして挙動が変わる（テストがブレる原因）
    assert _wait_no_excel(), "ブック作成に使った Excel が消えません（後始末が壊れています）"
    yield path


def peek(path):
    """ブックを読み取り専用・イベント無効で覗く（この検査自体が Workbook_Open を起こさない）"""
    import gc
    xl = _new_excel()
    xl.EnableEvents = False
    wb = None
    try:
        wb = xl.Workbooks.Open(path, 0, True)
        modules = sorted(c.Name for c in wb.VBProject.VBComponents)
        cm = wb.VBProject.VBComponents(MOD_NAME).CodeModule
        code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
        ws = wb.Worksheets("本体")
        result = {
            "modules": modules,
            "code": code,
            "Z1": ws.Range("Z1").Value,
            "C1": ws.Range("C1").Value,
        }
        cm = ws = None
        wb.Close(SaveChanges=False)
        wb = None
        return result
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = None
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()   # 覗きに使った Excel が消えるまで待つ（次の操作を汚さない）


def run_cli(*args, timeout=180):
    """vba_manager.py を別プロセスで叩く（実際の使われ方に一番近い）"""
    p = subprocess.run([sys.executable, VBM, *args], capture_output=True,
                       encoding="utf-8", errors="replace", cwd=SCRIPT_DIR, timeout=timeout)
    return p.returncode, (p.stdout or "") + (p.stderr or "")


# ================================================================
# 本命: 「Import 成功後に保存が失敗しても、連番モジュールを作らない」
# 2026-07-14 に発見・修正した実害の回帰。実機の VBE で確かめる。
# ================================================================

# この検証だけは「別プロセス」で走らせる。
# vbam_core を pytest と同じプロセスで動かすと、終了時の COM 後始末とかち合って
# インタプリタが落ち（全テストが通っているのに終了コード 1 になる）、
# さらに Excel の COM 参照が居残ってゾンビ扱いになる（2026-07-14 実測）。
# プロセスを分ければ COM はプロセス終了で確実に片付く。実際の使われ方にも近い。
_PROBE = '''# -*- coding: utf-8 -*-
import argparse, json, sys
sys.path.insert(0, r"{scripts}")
import vbam_core, vbam_vba

# vba_manager.py の入口と同じく stdout/stderr を UTF-8 にする。
# これを忘れると cp932 のままになり、案内文の「⚠」（cp932 に無い）で
# UnicodeEncodeError になる。ライブラリとして import するときの作法。
vbam_core.setup_encoding()

book, new_bas, mod = sys.argv[1], sys.argv[2], sys.argv[3]

# Import が終わった直後の保存だけを確実に失敗させる（ディスク要因の保存失敗を模す）
def boom(wb, *a, **kw):
    raise RuntimeError("模擬: 保存に失敗しました")
vbam_vba._save_with_retry = boom

try:
    args = argparse.Namespace(posargs=[book, mod, new_bas], yes=True, force=False)
    ok = vbam_vba.cmd_replace_module(args)

    # 開いているブックの中身をそのまま見る（保存されていないのでファイルには出ない）
    xl, wb = vbam_core.get_workbook(book)
    names = sorted(c.Name for c in wb.VBProject.VBComponents)
    cm = wb.VBProject.VBComponents(mod).CodeModule
    code = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines else ""
    cm = wb = xl = None

    print("__RESULT__" + json.dumps(
        {{"ok": bool(ok), "modules": names, "has_new_code": "999" in code}},
        ensure_ascii=False))
finally:
    # 起こした Excel を PID ごと確実に始末する（非表示・ブック0 の Excel を残さない）。
    # 未保存のまま破棄＝ディスクには一切触らせない。
    # cleanup_excel は内部で CoUninitialize するが、この使い捨てプロセスは
    # ここで終わるので問題ない（pytest と同じプロセスで呼ぶと落ちる）
    try:
        vbam_core._wb_cache.clear()
    except Exception:
        pass
    vbam_core.cleanup_excel()
'''


def test_save_failure_after_import_does_not_create_numbered_module(book, tmp_path):
    """保存だけ失敗したとき、バックアップを再 Import して連番モジュールを産まないこと。

    旧コードは except が removed フラグしか見ておらず、Import 成功後の Save 失敗を
    「モジュールが消えた」と誤認してバックアップを重ね Import していた。正しい名前の
    モジュールが既に在るので必ず衝突し、ツール自身が「旧コード入りの TestMod1」を
    産んで 35 秒待たせた末に失敗を告げていた（実機の VBE で連番化を再現済み）。

    ここでは Import が終わった直後の保存だけを確実に失敗させて、その分岐を撃つ。
    """
    new_bas = str(tmp_path / "new_mod.bas")
    new_code = (
        f'Attribute VB_Name = "{MOD_NAME}"\r\n'
        'Sub 一行完結(): Debug.Print 1: End Sub\r\n'
        'Sub 合計を出す()\r\n'
        '    Range("C1").Value = 999\r\n'   # 置換後だと分かる印
        'End Sub\r\n'
    )
    with open(new_bas, "wb") as f:
        f.write(new_code.encode("cp932"))

    probe = str(tmp_path / "probe_save_fail.py")
    with open(probe, "w", encoding="utf-8") as f:
        f.write(_PROBE.format(scripts=SCRIPT_DIR))

    p = subprocess.run([sys.executable, probe, book, new_bas, MOD_NAME],
                       capture_output=True, encoding="utf-8", errors="replace",
                       cwd=SCRIPT_DIR, timeout=300)
    out = (p.stdout or "") + (p.stderr or "")
    line = next((l for l in out.splitlines() if l.startswith("__RESULT__")), None)
    assert line, f"検証プロセスが結果を返さなかった:\n{out}"
    result = __import__("json").loads(line[len("__RESULT__"):])

    assert result["ok"] is False, f"保存に失敗したのに成功と報告している:\n{out}"

    names = result["modules"]
    numbered = [n for n in names
                if n.lower() != MOD_NAME.lower() and n.lower().startswith(MOD_NAME.lower())]
    assert not numbered, (
        f"連番モジュールが生まれている: {numbered}\n"
        "  Import 成功後の Save 失敗で、ツールがバックアップを再 Import している"
        f"（imported フラグのガードが外れた）。\n{out}")

    # 置換自体は済んでいる＝「消えた」のではなく「保存できなかった」だけ
    assert MOD_NAME in names, f"モジュールが消えている:\n{out}"
    assert result["has_new_code"], f"取り込みは成功しているはずなのに新コードが入っていない:\n{out}"
    # 誤誘導せず「保存だけ失敗した」と伝えていること
    assert "保存" in out


# ================================================================
# 回帰: これまでの修正を壊していないこと
# ================================================================

def test_list_does_not_trigger_workbook_open(book):
    """読み取り系（list）が、閉じたブックの Workbook_Open を起こさないこと"""
    rc, out = run_cli("list", book)
    assert rc == 0, out
    assert peek(book)["Z1"] is None, "list で Workbook_Open が走っている（健診モードが壊れた）"


def test_run_macro_executes_target_macro(book):
    """run-macro が対象ブックのマクロを実際に動かすこと"""
    batch = os.path.join(os.path.dirname(book), "_batch.txt")
    with open(batch, "w", encoding="utf-8") as f:
        f.write(f'run-macro "{book}" 合計を出す\n')
        f.write(f'save "{book}"\n')
    rc, out = run_cli("batch", batch)
    assert rc == 0, out
    assert peek(book)["C1"] == 5, "A1(2)+B1(3)=5 になっていない"


def test_replace_procedure_rejects_two_procedures(book):
    """コードファイルに2本入っていたら弾くこと（黙って両方入れない）"""
    with open(LAST_PROC, "w", encoding="utf-8") as f:
        f.write('Sub 一行完結(): Debug.Print 99: End Sub\n'
                'Sub 合計を出す()\n    Debug.Print 2\nEnd Sub\n')
    rc, out = run_cli("replace-procedure", book, "-y")
    assert rc == 1, f"2本入りを通してしまった: {out}"
    assert "2 本のプロシージャ" in out, out


def test_replace_procedure_keeps_other_procedures(book):
    """1行完結 Sub を巻き込んで消さないこと（過去の実害の回帰）"""
    with open(LAST_PROC, "w", encoding="utf-8") as f:
        f.write('Sub 合計を出す()\n    Range("C1").Value = 7\nEnd Sub\n')
    rc, out = run_cli("replace-procedure", book, "-y")
    assert rc == 0, out
    code = peek(book)["code"]
    assert "Sub 一行完結()" in code, "隣の1行完結 Sub が消えている"
    assert "Sub 合計を出す()" in code
    assert "7" in code


def test_read_only_commands_work(book):
    """grep / list-modules / list-shortcuts が動くこと（健診モード化の回帰）"""
    rc1, _ = run_cli("grep", book, "Debug.Print")
    rc2, out2 = run_cli("list-modules", book)
    rc3, _ = run_cli("list-shortcuts", book)
    assert (rc1, rc2, rc3) == (0, 0, 0)
    assert MOD_NAME in out2


def test_snapshot_and_diff(book, tmp_path):
    """snapshot / snapshot-diff が動き、変更が無ければ差分なしと言うこと"""
    s1 = str(tmp_path / "s1.json")
    s2 = str(tmp_path / "s2.json")
    rc1, _ = run_cli("snapshot", book, "--out", s1)
    rc2, _ = run_cli("snapshot", book, "--out", s2)
    rc3, out3 = run_cli("snapshot-diff", s1, s2)
    assert (rc1, rc2, rc3) == (0, 0, 0)
    assert "差分なし" in out3, out3


def test_delete_module_removes_it(book):
    """delete-module が実際に消し、保存すること"""
    import gc
    rc, out = run_cli("delete-module", book, MOD_NAME, "-y")
    assert rc == 0, out
    xl = _new_excel()
    wb = None
    try:
        wb = xl.Workbooks.Open(book, 0, True)
        names = [c.Name for c in wb.VBProject.VBComponents]
        wb.Close(SaveChanges=False)
        wb = None
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = None
        try:
            xl.Quit()
        except Exception:
            pass
        xl = None
        gc.collect()
        _wait_no_excel()
    assert MOD_NAME not in names, "delete-module が「削除完了」と言ったのに残っている"
