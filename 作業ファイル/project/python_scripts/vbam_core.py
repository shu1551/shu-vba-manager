# -*- coding: utf-8 -*-
"""vbam_core.py — vba_manager 分割パート: COM接続・共通ユーティリティ・ガード（衝突ガード/消滅待ち/共有ヘルパー）

vba_manager.py から機械分割（2026-07-12）。単体で実行せず、vba_manager.py 経由で使う。
"""
import sys
import os
import re
import shutil
import zlib
import argparse
import time
import datetime
import unicodedata
import pythoncom
import pywintypes
import win32com.client
import win32com.client.dynamic

import sys
import os
import re
import shutil
import zlib
import argparse
import time
import datetime
import unicodedata
import pythoncom
import pywintypes
import win32com.client
import win32com.client.dynamic

# ---- パス定数 ----
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR  = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', 'backups'))
LAST_PROC_FILE = os.path.join(SCRIPT_DIR, '_last_proc.vba')   # get の出力先
XL_EXTS     = ('.xlsm', '.xlam', '.xlsx', '.xls', '.xlsb')


# ================================================================
# ユーティリティ
# ================================================================

# このツールが自動起動した Excel インスタンスの記録。
# batch/shell で「未起動の別ファイル」を複数触ると DispatchEx が複数走るため、
# スカラー1個だと最後の1台しか始末できずゾンビが残る（2026-07-09 点検で発見）。
# 起動した全台を積んで、cleanup_excel で全部閉じる。
_created_instances = []          # [{"xl": <COM>, "pid": int|None}, ...]
# 後方互換の別名（コード内の他参照・デバッグ表示用に「最後の1台」を指す）
_created_xl = None
_created_xl_pid = None
# 直前の _get_workbook_uncached が「このツール自身でブックを開いた」かどうか。
# 健診モード(readonly)の後始末は自分が開いたブックにしか許されない
# （ユーザーが読み取り専用で開いていたブックを閉じる事故の防止・2026-07-14）
_last_open_by_tool = False


def setup_encoding():
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')   # --json 時の情報行が stderr に行くため同様に固定
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    pythoncom.CoInitialize()

    # --json オプションが引数に含まれている場合、通常の print を標準エラー出力にリダイレクトする
    if "--json" in sys.argv:
        import builtins
        _orig_print = builtins.print
        def custom_print(*args, **kwargs):
            if 'file' not in kwargs:
                kwargs['file'] = sys.stderr
            _orig_print(*args, **kwargs)
        builtins.print = custom_print


def looks_like_xl_file(s):
    """文字列がExcelファイルパスっぽいか判定。

    Excel拡張子で終わるものだけを対象ブックと見なす。
    以前は「パス区切りを含む」「ファイルとして存在する」でも True にしていたが、
    それだと `replace-procedure fix.vba` の fix.vba が対象ブック扱いになり、
    無関係ファイルのオープンや古い _last_proc.vba への静かなフォールバックを招く。
    """
    if not s:
        return False
    return s.lower().endswith(XL_EXTS)


def smart_path_resolve(filename):
    """ファイルパスを柔軟に解決"""
    if not filename:
        return None
    if os.path.exists(filename):
        return os.path.abspath(filename)
    for d in [os.getcwd(), SCRIPT_DIR,
              os.path.join(SCRIPT_DIR, '..', '..'),
              os.path.join(SCRIPT_DIR, '..', '..', '..')]:
        c = os.path.join(os.path.abspath(d), filename)
        if os.path.exists(c):
            return c
    return None


def parse_target_and_rest(posargs):
    """
    posargs の先頭が Excel ファイルなら target_file に、残りを rest に返す。
    そうでなければ target_file=None で全部 rest に返す。
    """
    if posargs and looks_like_xl_file(posargs[0]):
        return posargs[0], list(posargs[1:])
    return None, list(posargs)


# Excel が「今は応答できない」ときに返す HRESULT（参照が死んだわけではない）。
# マクロ実行中・モーダルダイアログ表示中・セル編集中は普通にこれが返る。
_COM_BUSY_HRESULTS = (
    -2147418111,  # 0x80010001 RPC_E_CALL_REJECTED
    -2147417846,  # 0x8001010A RPC_E_SERVERCALL_RETRYLATER
    -2147417847,  # 0x80010109 RPC_E_SERVERCALL_REJECTED
    -2147417851,  # 0x80010105 RPC_E_SERVERFAULT
    -2146777072,  # 0x800AC472 VBA_E_IGNORE（Excel がセル編集中・モーダル表示中に返す
                  #             一番ありふれた「今は無理」。これを外すと、ビジーな
                  #             Excel を「死んでいる」と誤判定して強制終了し、
                  #             未保存の変更を捨てる＝温存の約束を破る）
)


def _com_is_busy(ex):
    """COM 例外が「Excel がビジー（＝生きているが今は応答できない）」かどうか。

    判定できない例外（COM 以外・HRESULT が読めない）は False を返すが、
    呼び出し側はそれを「死んでいる」と決めつけずに扱うこと。
    """
    for val in (getattr(ex, 'hresult', None),) + tuple(getattr(ex, 'args', ()) or ()):
        if isinstance(val, int) and val in _COM_BUSY_HRESULTS:
            return True
    return False


def _pid_is_excel(pid):
    """その PID が今も EXCEL.EXE かを確認する（強制終了の前の身元確認）。

    Windows は終了したプロセスの PID を再利用する。参照が死んだ＝プロセスは
    もう無い、という状況で PID を撃つと、その番号を引き継いだ無関係のプロセスを
    殺しうる。確認できないときは False（撃たない）を返す。
    """
    if pid is None:
        return False
    try:
        import win32api
        import win32con
        import win32process
        h = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
        try:
            name = win32process.GetModuleFileNameEx(h, 0)
        finally:
            win32api.CloseHandle(h)
        return os.path.basename(name).lower() == 'excel.exe'
    except Exception:
        return False


def release_created_instances(only_saved=True):
    """自動起動した Excel インスタンスのうち、後始末してよいものを閉じて終了する。

    常駐プロセス（MCPサーバー）では終了時の cleanup_excel が何時間も来ないため、
    パス指定で自動起動した非表示 Excel がセッション中ずっと残留する。COM 起動の
    Excel はアドイン・PERSONAL.XLSB を読まないため、残留中にユーザーが開いた
    ファイルがそこへ合流すると「アドインが効かない・マクロが効かない」状態になる
    （2026-07-12 特定。6/13 ゾンビ事故と同症状。当日実測で非表示3体が残留していた）。
    MCP のアイドル解放（5秒）からこれを呼んで畳む。

    only_saved=True では未保存の変更を持つインスタンスは温存する
    （アイドルのたびに無言で変更破棄をしないため。それらは終了時の
    cleanup_excel が警告つきで始末する）。戻り値: (閉じた台数, 温存した台数)
    """
    global _created_instances, _created_xl, _created_xl_pid
    import gc
    import os
    import signal
    keep = []
    closed = 0
    for inst in _created_instances:
        xl = inst.get("xl")
        pid = inst.get("pid")
        alive = True
        busy = False
        dirty = False
        try:
            wbs = xl.Workbooks
            for i in range(wbs.Count, 0, -1):
                try:
                    if not wbs.Item(i).Saved:
                        dirty = True
                        break
                except Exception:
                    pass
        except Exception as ex:
            if _com_is_busy(ex):
                # マクロ実行中・モーダル表示中。生きているし、未保存かどうかも
                # 今は判定できない。ここで畳むと only_saved の約束（未保存は温存）を
                # 破って変更を捨てるので、触らず次のアイドルに回す
                busy = True
            else:
                alive = False  # 参照が本当に死んでいる（手動で閉じられた等）
        if busy:
            keep.append(inst)
            continue
        if alive and only_saved and dirty:
            keep.append(inst)
            continue
        if alive:
            try:
                wbs = xl.Workbooks
                for i in range(wbs.Count, 0, -1):
                    try:
                        wbs.Item(i).Close(SaveChanges=False)
                    except Exception:
                        pass
                xl.Quit()
            except Exception:
                pass
        # Quit がゾンビ化しても残さない（cleanup_excel と同じ最終手段）。ただし
        # PID の身元を確認してから撃つ（再利用された PID の巻き添えを防ぐ）
        if _pid_is_excel(pid):
            try:
                os.kill(pid, signal.SIGTERM)
            except Exception:
                pass
        closed += 1
    _created_instances = keep
    if not keep:
        _created_xl = None
        _created_xl_pid = None
    gc.collect()
    return closed, len(keep)


def cleanup_excel():
    """新規起動されたExcelインスタンスがあれば（複数でも）全て終了し、COMを初期化解除する"""
    global _created_instances, _created_xl, _created_xl_pid
    import gc
    import os
    import signal
    # 自動起動したインスタンスが無ければ何もしない回なので、DEBUG も出さない
    # （全コマンドの末尾に毎回2行のノイズが出ていた）
    if _created_instances:
        print(f"[DEBUG] cleanup_excel called. instances: {len(_created_instances)}")

    # Python側のCOM参照を解放するためにGCを強制実行
    gc.collect()

    for inst in _created_instances:
        xl = inst.get("xl")
        pid = inst.get("pid")
        if xl is not None:
            try:
                print("[DEBUG] Closing open workbooks...")
                try:
                    # 切断エラーを回避しつつ、逆順にブックを閉じる
                    wbs = xl.Workbooks
                    for i in range(wbs.Count, 0, -1):
                        try:
                            wb = wbs.Item(i)
                            name = wb.Name
                            try:
                                if not wb.Saved:
                                    # 自動起動経路では閉じる＝未保存変更の破棄。無言だと
                                    # 「書いたつもりが消えていた」になるため明示する
                                    print(f"⚠ 未保存の変更を破棄して閉じます: {name}")
                                    print("  （保存したい場合は同じ batch 内で save を実行するか、"
                                          "Excel を先に起動してから作業してください）")
                            except Exception:
                                pass
                            wb.Close(SaveChanges=False)
                            print(f"[DEBUG] Workbook closed: {name}")
                        except Exception as ex:
                            print(f"[DEBUG] Failed to close wb {i}: {ex}")
                except Exception as ex:
                    print(f"[DEBUG] Failed to access Workbooks: {ex}")

                print("[DEBUG] Calling xl.Quit()...")
                xl.Quit()
                print("[DEBUG] xl.Quit() completed.")
            except Exception as ex:
                print(f"[DEBUG] Error during Excel cleanup: {ex}")

        # 新規起動したPIDが存在する場合は強制クリーンアップ。
        # 撃つ前に「その PID が今も EXCEL.EXE か」を確認する（既に終了していた場合、
        # Windows は PID を再利用するため無関係のプロセスを殺しうる）
        if pid is not None:
            if not _pid_is_excel(pid):
                print(f"[DEBUG] PID {pid} は既に EXCEL.EXE ではありません（終了済み）。強制終了はしません。")
            else:
                try:
                    print(f"[DEBUG] Force-killing Excel process (PID: {pid})...")
                    os.kill(pid, signal.SIGTERM)
                    print("[DEBUG] Excel process force-killed successfully.")
                except Exception as ex:
                    print(f"[DEBUG] Excel process force-kill failed or already exited: {ex}")

    _created_instances = []
    _created_xl = None
    _created_xl_pid = None

    # 最後の解放
    gc.collect()
    try:
        pythoncom.CoUninitialize()
    except Exception as ex:
        print(f"[DEBUG] CoUninitialize failed: {ex}")


def load_excel_addins_and_personal(xl):
    """新規起動されたExcelにアドインとPERSONAL.XLSBをロードする"""
    import os
    import time

    # 警告を非表示にする
    try:
        xl.DisplayAlerts = False
    except Exception:
        pass

    loaded_any = False

    # 1. アドインのロード
    try:
        for addin in xl.AddIns:
            if addin.Installed:
                try:
                    # すでに開いていなければ開く
                    opened = False
                    for wb in xl.Workbooks:
                        if wb.Name.lower() == os.path.basename(addin.FullName).lower():
                            opened = True
                            break
                    if not opened:
                        xl.Workbooks.Open(addin.FullName)
                        print(f"[DEBUG] Loaded addin: {addin.Name}")
                        loaded_any = True
                except Exception as ex:
                    # Excelがビジー状態の時のリトライ (0x800ac472)
                    if "800ac472" in str(ex):
                        time.sleep(0.5)
                        try:
                            xl.Workbooks.Open(addin.FullName)
                            print(f"[DEBUG] Loaded addin (retry): {addin.Name}")
                            loaded_any = True
                        except Exception as ex2:
                            print(f"[DEBUG] Failed to load addin {addin.Name} after retry: {ex2}")
                    else:
                        print(f"[DEBUG] Failed to load addin {addin.Name}: {ex}")
    except Exception as ex:
        print(f"[DEBUG] Failed to access AddIns: {ex}")

    # 2. PERSONAL.XLSB のロード
    try:
        startup_path = xl.StartupPath
        if startup_path:
            for fname in ["PERSONAL.XLSB", "personal.xlsb", "PERSONAL.XLS", "personal.xls"]:
                p_path = os.path.join(startup_path, fname)
                if os.path.exists(p_path):
                    opened = False
                    for wb in xl.Workbooks:
                        if wb.Name.lower() == fname.lower():
                            opened = True
                            break
                    if not opened:
                        try:
                            xl.Workbooks.Open(p_path)
                            print(f"[DEBUG] Loaded personal macro book: {fname}")
                            loaded_any = True
                        except Exception as ex:
                            if "800ac472" in str(ex):
                                time.sleep(0.5)
                                try:
                                    xl.Workbooks.Open(p_path)
                                    print(f"[DEBUG] Loaded personal macro book (retry): {fname}")
                                    loaded_any = True
                                except Exception as ex2:
                                    print(f"[DEBUG] Failed to load personal macro book {fname} after retry: {ex2}")
                            else:
                                print(f"[DEBUG] Failed to load personal macro book {fname}: {ex}")
                    break
    except Exception as ex:
        print(f"[DEBUG] Failed to load personal macro book: {ex}")

    # ロード処理が走った場合は待機
    if loaded_any:
        time.sleep(1.0)

    try:
        xl.DisplayAlerts = True
    except Exception:
        pass


def _get_active_excel():
    """起動中の Excel.Application に late-binding で接続する(gencache 非経由)。

    win32com.client.GetActiveObject は gen_py キャッシュを通るため、キャッシュ破損
    (例: module '...' has no attribute 'CLSIDToClassMap')で例外になり、Excel が実際は
    開いているのに「起動していない」と誤判定する。pythoncom.GetActiveObject で生の
    インスタンスを掴んで dynamic.Dispatch で包めばキャッシュに一切依存しない。
    Excel 未起動時は例外を送出する(呼び出し側で捕捉する)。
    """
    clsid = pywintypes.IID("Excel.Application")
    unk = pythoncom.GetActiveObject(clsid)
    disp = unk.QueryInterface(pythoncom.IID_IDispatch)
    return win32com.client.dynamic.Dispatch(disp)


def _running_excel_workbooks():
    """Running Object Table を走査し、起動中の全 Excel の全ブックを返す。

    GetActiveObject は ROT 先頭の1インスタンスしか返さず、非表示ゾンビ Excel
    (ブック0個)を掴むことがある。ROT を直接舐めれば、実際にブックを開いている
    インスタンスだけ拾える(ゾンビはブックを持たないので現れない)。
    失敗時は [] を返し、呼び出し側は GetActiveObject にフォールバックする。
    """
    wbs = []
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        monikers = list(rot)
    except Exception:
        return wbs
    for mon in monikers:
        try:
            disp = mon.GetDisplayName(ctx, None)
        except Exception:
            continue
        # Excel のブックはフルパス(...\xxx.xlsm 等)で ROT 登録される
        if not disp or not re.search(r'\.xl\w{1,4}$', disp, re.IGNORECASE):
            continue
        try:
            obj = rot.GetObject(mon)
            wb = win32com.client.dynamic.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
            _ = wb.FullName  # ブックか確認(違えば例外)
            wbs.append(wb)
        except Exception:
            continue
    return wbs


# get_workbook の接続キャッシュ。通常の1コマンド1プロセスでは1回しか呼ばれないので
# 実質無関係だが、batch モードでは全行が同じCOM接続を使い回せる（1コマンド毎の
# COM再接続＝ROT走査が一番重い、という実測への構造的な解）。
_wb_cache = {}


def get_workbook(target_file_arg=None, load_addins=False, readonly=False):
    """get_workbook（接続キャッシュつきの入口）。戻り値: (xl, wb)

    readonly=True は診断系コマンド用。未起動ブックを自動で開く場合に
    「読み取り専用＋イベント無効」で開く（Workbook_Open 等を起こさない健診モード）。
    既に開いているブックには影響しない。
    """
    key = "__active__"
    if target_file_arg:
        resolved = smart_path_resolve(target_file_arg)
        key = resolved.lower() if resolved else target_file_arg.lower()
    if key in _wb_cache:
        xl, wb, was_ro = _wb_cache[key]
        try:
            _ = wb.Name          # 生存確認（閉じられていたら再接続）
            if key == "__active__":
                # shell/batch/MCP 等の1接続セッション中にユーザーが Excel 側で
                # 別ブックをアクティブにした場合、キャッシュした旧ブックのまま
                # 破壊コマンドが走ると対象取り違えになる。毎回同一性を確認して追従する
                try:
                    cur = xl.ActiveWorkbook
                    if cur is not None and cur.Name != wb.Name:
                        print(f"対象ブック: {cur.Name}（アクティブブックの切替に追従）")
                        wb = cur
                        _wb_cache[key] = (xl, wb, was_ro)
                except Exception:
                    pass
            if was_ro and not readonly:
                # 健診モード（読み取り専用・イベント無効）で「このツールが」開いた
                # ブックを、同じセッション（batch/shell/MCP は接続キャッシュを持ち越す）で
                # 書き込み系に渡すと、編集が分かりにくい COM エラーで失敗する。
                # 例: get <path> → replace-procedure <path>
                # 警告だけでは直らないので、書き込みが要るときは通常モードで開き直す。
                # ※ was_ro はツール自身が開いたときにしか立たない（_last_open_by_tool）。
                #   ユーザーが読み取り専用で開いていたブックをここで閉じてはならない。
                print("（診断用に読み取り専用で開いていたブックを、書き込みのため通常モードで開き直します）")
                try:
                    wb.Close(SaveChanges=False)
                except Exception as ex:
                    # 閉じられない（ビジー・モーダル表示中等）のに開き直したことにすると、
                    # 読み取り専用のまま書き込みに進み、Save で落ちるか黙って捨てられる
                    print(f"エラー: 読み取り専用で開いたブックを閉じられませんでした: {ex}")
                    print("  Excel がビジー（マクロ実行中・ダイアログ表示中）の可能性があります。")
                    print("  少し待ってから再実行してください。")
                    raise
                try:
                    xl.EnableEvents = True   # 健診モードで切っていたイベントを戻す
                except Exception as ex:
                    print(f"⚠ EnableEvents を戻せませんでした: {ex}", file=sys.stderr)
                    print("  このExcelではイベントマクロ(Worksheet_Change等)が動きません。",
                          file=sys.stderr)
                del _wb_cache[key]
                xl, wb = _get_workbook_uncached(target_file_arg, load_addins, False)
                # 本当に書き込み可能になったかを実測する（「開き直しました」と言いながら
                # 読み取り専用のまま、を防ぐ）
                try:
                    if bool(wb.ReadOnly):
                        raise Exception(
                            "開き直しましたが、まだ読み取り専用です"
                            "（ファイルが読み取り専用属性、または他で開かれています）")
                except pywintypes.com_error:
                    pass
                _wb_cache[key] = (xl, wb, False)
                return xl, wb
            if load_addins:
                load_excel_addins_and_personal(xl)
            return xl, wb
        except Exception:
            del _wb_cache[key]
    xl, wb = _get_workbook_uncached(target_file_arg, load_addins, readonly)
    # 健診モードの印は「このツールが今このブックを開いた場合」だけ立てる。
    # wb.ReadOnly だけで判定すると、ユーザーが読み取り専用で開いていたブック
    # （読み取り専用属性・読み取り専用推奨・他者がロック中・「読み取り専用で開く」を
    # 選択）を健診モードと誤認し、後で書き込み系が来たときに勝手に Close して
    # 開き直す＝ユーザーのウィンドウが消える（2026-07-14 実機で再現）。
    opened_ro = readonly and target_file_arg is not None and _last_open_by_tool
    try:
        opened_ro = opened_ro and bool(wb.ReadOnly)
    except Exception:
        pass
    _wb_cache[key] = (xl, wb, opened_ro)
    return xl, wb


def _get_workbook_uncached(target_file_arg=None, load_addins=False, readonly=False):
    """
    target_file_arg が None/空 → アクティブExcelブックを自動使用
    それ以外 → 既に開いているか確認、なければ新規オープン
    戻り値: (xl, wb)

    副作用: グローバル _last_open_by_tool に「このツールが今開いたのか
    （True）／既に開いていたブックを掴んだだけか（False）」を記録する。
    健診モード(readonly)の後始末は自分が開いたブックにしか許されないため
    （ユーザーが読み取り専用で開いていたブックを閉じる事故の防止）。
    """
    global _last_open_by_tool
    _last_open_by_tool = False
    pythoncom.CoInitialize()

    if not target_file_arg:
        # ROT 全インスタンス横断で、実ブックを持つ Excel を選ぶ(ゾンビ自動回避)
        wb = None
        for cand in _running_excel_workbooks():
            try:
                app = cand.Application
                if app.Visible and app.ActiveWorkbook is not None:
                    wb = app.ActiveWorkbook   # 可視インスタンスのアクティブブックを最優先
                    break
                if wb is None:
                    wb = cand                  # 暫定: 実ブックを持つ最初のもの
            except Exception:
                continue
        if wb is None:
            # ROT に出ない稀ケース(未保存ブック等)は GetActiveObject で再挑戦
            try:
                wb = _get_active_excel().ActiveWorkbook
            except Exception:
                wb = None
        if wb is None:
            raise Exception(
                "起動中の Excel に開いているブックが見つかりません。\n"
                "  ・Excel で対象ブックを開いてから再実行してください。\n"
                "  ・非表示のゾンビ EXCEL.EXE が残っている場合があります。"
                "タスクマネージャーで余分な EXCEL.EXE を終了し、対象ブックを開いて再実行してください。\n"
                "  ※ COM 接続できないからといって .bas を手書きスクリプトで処理しないこと"
                "(改行二重化の原因)。")
        xl = wb.Application
        print(f"対象ブック: {wb.Name}  (アクティブブック自動検出)")
        if load_addins:
            load_excel_addins_and_personal(xl)
        return xl, wb

    target_path = smart_path_resolve(target_file_arg)
    if not target_path:
        raise Exception(f"ファイルが見つかりません: {target_file_arg}")

    # 既に開いているか確認(全 Excel インスタンス横断)
    excel_running = False
    for wb in _running_excel_workbooks():
        excel_running = True
        try:
            if wb.FullName.lower() == target_path.lower():
                xl = wb.Application
                print(f"対象ブック: {wb.Name}  (既に開いています)")
                if load_addins:
                    load_excel_addins_and_personal(xl)
                return xl, wb
        except Exception:
            continue
    if not excel_running:
        try:
            xl_fallback = _get_active_excel()
            excel_running = True
            # ROT 走査が失敗していても実際は開いているケースの二重オープン防止:
            # GetActiveObject で掴んだインスタンスの Workbooks も確認する
            try:
                for wb in xl_fallback.Workbooks:
                    if wb.FullName.lower() == target_path.lower():
                        print(f"対象ブック: {wb.Name}  (既に開いています)")
                        if load_addins:
                            load_excel_addins_and_personal(xl_fallback)
                        return xl_fallback, wb
            except Exception:
                pass
        except Exception:
            excel_running = False

    # 新規オープン
    # ★必ず DispatchEx を使う。Dispatch は既存インスタンスがあるとそこに接続してしまい、
    #   「自分が起動した Excel」と誤認 → 後始末でユーザーの Excel ごと閉じる大事故になる
    #   （2026-07-03 実害。ユーザーのブックを巻き込んで Quit した）
    xl = win32com.client.DispatchEx("Excel.Application")
    global _created_xl, _created_xl_pid
    _created_xl = xl
    try:
        import win32process
        hwnd = xl.Hwnd
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        _created_xl_pid = pid
        print(f"[DEBUG] Excel process PID detected: {_created_xl_pid}")
    except Exception as ex:
        _created_xl_pid = None
        print(f"[DEBUG] Failed to detect Excel PID: {ex}")
    # 起動した全台を記録（batch/shell で複数起動しても cleanup で全部閉じるため）
    _created_instances.append({"xl": xl, "pid": _created_xl_pid})
    xl.Visible = "--visible" in sys.argv or "-v" in sys.argv

    if load_addins:
        load_excel_addins_and_personal(xl)

    if readonly:
        # 健診モード: Workbook_Open 等の自動実行を起こさず、リンク更新もせず、
        # 読み取り専用で開く（「診察に行ったら患者を起こしてしまった」の防止）
        try:
            xl.EnableEvents = False
        except Exception:
            pass
        wb = xl.Workbooks.Open(target_path, 0, True)   # UpdateLinks=0, ReadOnly=True
        _last_open_by_tool = True
        print(f"対象ブック: {wb.Name}  (新規オープン・読み取り専用・イベント無効=健診モード)")
        return xl, wb

    wb = xl.Workbooks.Open(target_path)
    _last_open_by_tool = True
    print(f"対象ブック: {wb.Name}  (新規オープン)")
    if not excel_running and not load_addins:
        # COM起動のExcelは起動処理が走らず、アドインや PERSONAL.XLSB が読み込まれない
        print("注意: Excelが未起動だったため自動化用に新規起動しました。")
        print("      このExcelにはアドイン(秀.xlam等)・PERSONAL.XLSB が読み込まれていません。")
        print("      普段使いにはこのウィンドウを閉じて、手動起動した Excel で開き直してください。")
    return xl, wb


def make_backup(wb_fullname, label):
    """バックアップを作成（タイムスタンプ付き・同系列は直近5世代まで保持）。

    成功時はバックアップパス、失敗時は None を返す。
    破壊的操作（replace-procedure / replace-module）の呼び出し元は
    None のとき停止する（--force で続行可）。
    """
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ext = os.path.splitext(wb_fullname)[1] or '.xlsm'
    stamp = time.strftime("%Y%m%d_%H%M%S")
    # フォルダ違いの同名ブックが同じ系列に混ざると、世代の間引きが
    # もう一方のブックのバックアップを消してしまう。フォルダの短いタグで系列を分ける
    dirtag = format(zlib.crc32(os.path.dirname(os.path.abspath(wb_fullname))
                               .lower().encode('utf-8')) & 0xffff, '04x')
    prefix = os.path.basename(wb_fullname) + f".backup_before_{label}_{dirtag}_"
    backup_path = os.path.join(BACKUP_DIR, prefix + stamp + ext)
    # 同一秒内の連続バックアップ（batch での連続置換等）を黙って上書きしない
    seq = 1
    while os.path.exists(backup_path):
        seq += 1
        backup_path = os.path.join(BACKUP_DIR, f"{prefix}{stamp}_{seq}{ext}")
    try:
        shutil.copy2(wb_fullname, backup_path)
        print(f"バックアップ作成: backups/{os.path.basename(backup_path)}")
    except Exception as e:
        print(f"警告: バックアップ失敗 ({e})")
        return None
    # 同じ系列（同ブック・同ラベル）の古い世代を間引いて5世代までにする
    try:
        olds = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(prefix))
        for f in olds[:-5]:
            os.remove(os.path.join(BACKUP_DIR, f))
    except Exception:
        pass
    return backup_path


def _remove_export_artifacts(bas_path):
    """一時 Export の後始末。

    フォーム（UserForm）を Export すると同名の .frx が必ず併産されるが、
    従来は .bas しか消しておらず _tmp_*.frx が溜まり続けていた（構造的な掃除漏れ）。
    ペアで削除する。
    """
    for p in (bas_path, os.path.splitext(bas_path)[0] + '.frx'):
        try:
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass


def read_code_file(path):
    """コードファイルを UTF-8 / CP932 で読み込む（改行は \\n に正規化）。

    改行二重化(\\r\\r\\n)の水際検知＝多層防御の①層。テキストモード読みでは
    \\r\\r\\n が \\n\\n（空行）に化けて検知不能になるため、バイトで検知してから復号する。
    外部エディタ等で二重化済みのファイルもここで畳まれ、素通しでモジュールに入らない。
    """
    with open(path, 'rb') as f:
        raw = f.read()
    if b'\r\r' in raw:
        print(f"⚠ {os.path.basename(path)} に改行の二重化(\\r\\r\\n)を検知しました。正規化して読み込みます。")
    for enc in ['utf-8-sig', 'utf-8', 'cp932']:
        try:
            text = raw.decode(enc)
        except Exception:
            continue
        # \r\n / \r\r\n / lone \r をすべて \n に畳む（従来のテキストモード読みと互換）
        return re.sub(r'\r+\n?', '\n', text)
    raise Exception(f"ファイルを読み込めません: {path}")


def validate_bas_encoding(path):
    """.bas が CP932 で安全にインポートできるかの水際チェック。

    過去に繰り返された「CP932 の .bas を UTF-8 で上書きして日本語を壊す」事故を
    インポート前に機械的に検知する。日本語を含む CP932 のバイト列が偶然
    完全な UTF-8 として解釈できることはまず無いので、
    『非ASCIIを含むのに UTF-8 として読める』＝ UTF-8 化の疑い濃厚、として弾く。
    """
    with open(path, 'rb') as f:
        data = f.read()
    name = os.path.basename(path)
    if data.startswith(b'\xef\xbb\xbf'):
        print(f"エラー: {name} に UTF-8 BOM があります。")
        print("  この .bas は UTF-8 で保存されています。CP932(Shift-JIS) に戻してから実行してください。")
        return False
    if any(b > 0x7F for b in data):
        try:
            data.decode('cp932')
            cp932_ok = True
        except UnicodeDecodeError:
            cp932_ok = False
        try:
            data.decode('utf-8')
            utf8_ok = True
        except UnicodeDecodeError:
            utf8_ok = False
        if not cp932_ok:
            print(f"エラー: {name} は CP932 として読めません。エンコーディングを確認してください。")
            return False
        if utf8_ok:
            print(f"エラー: {name} は UTF-8 で保存されている疑いが濃厚です。")
            print("  このままインポートすると日本語が文字化けします。CP932(Shift-JIS) に変換してから実行してください。")
            return False
    return True


def normalize_bas_newlines(path):
    """インポート前に .bas の改行を正規 CRLF に揃える（改行二重化アーティファクトの水際修正）。

    「export → 編集 → テキストモードで書き戻し」の経路では、\\r\\n の各 \\n の前に
    余分な \\r が足されて \\r\\r\\n になることがある。これを VBA の Import に通すと
    1行おきに空行が挟まり、モジュールの行数が倍に膨れる（過去に繰り返した二重化事故の正体）。
    validate_bas_encoding が「文字コード事故」を弾くのと対になる「改行事故」の水際チェック。

    返り値 (fixed_bytes, raw_bytes, was_fixed):
        was_fixed=False なら元と完全一致＝無加工。True なら二重化等を検知・修正済み。

    クリーンな CRLF ファイルには冪等（バイト列が変わらないので was_fixed=False）。
    空行 (\\r\\n\\r\\n) は各 \\r\\n が個別にマッチするため保持される。
    """
    with open(path, 'rb') as f:
        raw = f.read()
    text = raw.decode('cp932')
    # \r を1個以上含む改行（\r\n / \r\r\n / \r\r\r\n / lone \r）を一旦 \n に畳む
    norm = re.sub(r'\r+\n?', '\n', text)
    # lone \n（Unix改行）も含め、すべて正規 CRLF へ揃える
    norm = norm.replace('\n', '\r\n')
    fixed = norm.encode('cp932')
    return fixed, raw, (fixed != raw)


def validate_vba_code(code, force=False):
    """VBAコードの簡易バリデーション（構文・エンコード）"""
    # 1. CP932エンコード検証
    try:
        code.encode('cp932')
    except UnicodeEncodeError as e:
        bad_char = code[e.start:e.end]
        print(f"エラー: CP932(Shift_JIS)でエンコードできない文字 '{bad_char}' (インデックス: {e.start}) が含まれています。")
        print("VBAマクロでは文字化けやインポートエラーの原因となるため、修正してください。")
        if not force:
            return False
        print("警告: --force が指定されているため、検証エラーを無視して処理を続行します。")

    # 2. 簡易構文チェック (Sub/End Sub, Function/End Function の対のチェック)
    # コメント行を除外して検索。マルチステートメント行（Sub x(): End Sub 等）は
    # ':' で文に分割してから数える（1行書きを「End Sub 不足」と誤警告しない）
    clean_lines = []
    for line in code.split('\n'):
        stripped = line.strip()
        if stripped.startswith("'") or stripped.lower().startswith("rem "):
            continue
        blanked = re.sub(r'"[^"]*"', '""', stripped)   # 文字列内の ':' で誤分割しない
        for seg in blanked.split(':'):
            clean_lines.append(seg.strip())
    clean_code = '\n'.join(clean_lines)

    decl_sub = len(re.findall(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?\bSub\b', clean_code, re.IGNORECASE | re.MULTILINE))
    end_sub = len(re.findall(r'^\s*End\s+Sub\b', clean_code, re.IGNORECASE | re.MULTILINE))
    decl_func = len(re.findall(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?\bFunction\b', clean_code, re.IGNORECASE | re.MULTILINE))
    end_func = len(re.findall(r'^\s*End\s+Function\b', clean_code, re.IGNORECASE | re.MULTILINE))

    errors = []
    if decl_sub != end_sub:
        errors.append(f"Sub宣言数 ({decl_sub}) と End Sub数 ({end_sub}) が一致しません。")
    if decl_func != end_func:
        errors.append(f"Function宣言数 ({decl_func}) と End Function数 ({end_func}) が一致しません。")

    # 3. プロシージャ名の識別子チェック（`_tmp検証` 級の先頭 _ 注入が繰り返された定番事故。
    #    AddFromString/InsertLines は構文検査をせず成功報告になるため、ここが水際）
    for ln, _name, reason in _find_invalid_procedure_names(code):
        errors.append(f"行{ln}: プロシージャ名が VBA の識別子規則に反しています: {reason}")

    if errors:
        for err in errors:
            print(f"構文エラー警告: {err}")
        if not force:
            print("エラー: 構文チェックに失敗しました。(--force で無視して実行可能)")
            return False
        print("警告: --force が指定されているため、構文エラーを無視して処理を続行します。")
    return True


def make_module_backup(wb, module_name):
    """モジュール単位のバックアップを .bas 形式で保存"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    base_wb_name = os.path.splitext(os.path.basename(wb.FullName))[0]
    backup_filename = f"{base_wb_name}_{module_name}_{timestamp}.bas"
    backup_path = os.path.join(BACKUP_DIR, backup_filename)

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() == module_name.lower():
            try:
                comp.Export(backup_path)
                print(f"モジュールバックアップ作成: backups/{backup_filename}")
                return backup_path
            except Exception as e:
                print(f"警告: モジュールバックアップ失敗 ({e})")
    return None


class ModuleNameCollisionError(Exception):
    """Remove+Import の名前衝突で、モジュールを期待名で取り込めなかったことを示す。

    Import 自体は成功しており、置換後のコードは actual_name（連番付き別名）側に
    存在している。呼び出し側はバックアップの再 Import で「復旧」してはいけない
    （同じ VB_Name の .bas を重ねると連番モジュールがさらに増えるだけ）。
    """
    def __init__(self, expected_name, actual_name, message):
        super().__init__(message)
        self.expected_name = expected_name
        self.actual_name = actual_name


def _find_component(wb, name):
    """VBComponent を名前（大小無視）で探す。無ければ None"""
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == name.lower():
            return c
    return None


def _wait_component_gone(wb, name, timeout=15.0, interval=0.25):
    """Remove 発行後、同名コンポーネントが VBProject から実際に消えるまで待つ。

    VBE の Remove は、対象モジュールのプロシージャが実行中（メニューや
    ショートカット経由の呼び出し中）などの場合に遅延完了する。消える前に
    Import すると名前衝突で「shu0051」のような連番付き別名で取り込まれる
    （2026-07-11 深夜の shu005 消滅事故の直接原因）。
    戻り値: 消えたら True / timeout まで残っていたら False
    """
    deadline = time.monotonic() + timeout
    while True:
        pythoncom.PumpWaitingMessages()
        if _find_component(wb, name) is None:
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(interval)


def _reregister_shortcuts_from_bas(wb, import_path):
    """Import した .bas の Attribute VB_Invoke_Func からショートカットをその場で再登録する。

    ショートカット（Attribute VB_ProcData.VB_Invoke_Func）が Excel に登録されるのは
    「ブックを開いた瞬間」だけ。Remove+Import はモジュールごと登録を剥がすため、
    Import 直後のセッションではショートカットが無反応になり、閉じて開き直すまで
    直らない（2026-07-12 発覚。修正のたびにマクロが効かなくなる症状の正体）。
    ここで MacroOptions により実行中セッションへ再登録して、開き直し不要にする。
    大文字=Ctrl+Shift+キー / 小文字=Ctrl+キー（MacroOptions の仕様どおり渡す）。
    """
    try:
        with open(import_path, 'rb') as f:
            text = f.read().decode('cp932', errors='replace')
    except Exception:
        return
    pairs = re.findall(
        r'^Attribute\s+([^\s.]+)\.VB_ProcData\.VB_Invoke_Func\s*=\s*"(.)\\n14"',
        text, re.MULTILINE)
    if not pairs:
        return
    xl = wb.Application
    for proc, key in pairs:
        if not key.strip():
            continue  # 空白キー（キー未割当の名残）は再登録しない
        try:
            xl.MacroOptions(Macro=f"'{wb.Name}'!{proc}",
                            HasShortcutKey=True, ShortcutKey=key)
            label = f"Ctrl+Shift+{key.upper()}" if key.isupper() else f"Ctrl+{key}"
            print(f"  ショートカット再登録: {label} → {proc}")
        except Exception as ex:
            print(f"  ⚠ ショートカット再登録に失敗: {proc} ({ex})")


def _import_module_verified(wb, import_path, expected_name,
                            ghost_timeout=15.0, rename_timeout=20.0, settle=1.5):
    """Import ＋ 取り込み実名の検証。Remove+Import 系3経路の共用ガード。

    2026-07-11 深夜の実害: replace-procedure（Attribute経路）で shu005 を置換した際、
    Remove の遅延完了中に Import が走って名前衝突し「shu0051」として取り込まれた。
    ツールは成功と報告し、shu005 の消滅は翌日まで発覚しなかった。
    同じ事故を二度と起こさないため、ここで
      (1) Import 前: expected_name の残骸が消えたことを確認（遅延 Remove 対策）。
          timeout しても Import 自体は行う（見送ると遅延 Remove だけが後から完了して
          モジュールが完全消滅するため、衝突覚悟で取り込んで (3) で回復する）。
      (2) Import 後: 返ってきた VBComponent の実名を expected_name と照合。
      (3) 不一致（連番付き別名等）なら旧名の消滅を待って改名で自動回復。
          回復できなければ ModuleNameCollisionError（黙って成功と報告しない）。
    成功時は Import 済みコンポーネントを返す（Save は呼び出し側の責務）。
    """
    if not _wait_component_gone(wb, expected_name, timeout=ghost_timeout):
        print(f"⚠ 旧 '{expected_name}' の Remove がまだ完了していません"
              f"（同モジュールのコードが実行中の可能性）。実名検証つきで Import を続行します。")
    imported = wb.VBProject.VBComponents.Import(import_path)
    if settle:
        time.sleep(settle)
    pythoncom.PumpWaitingMessages()
    actual = imported.Name
    if actual.lower() == expected_name.lower():
        _reregister_shortcuts_from_bas(wb, import_path)
        return imported
    # 名前衝突を検出（例: shu005 → shu0051）。旧名が空くのを待って改名で回復する
    print(f"⚠ 名前衝突を検出: '{expected_name}' が '{actual}' として取り込まれました。"
          f"旧モジュールの消滅を待って改名で回復します...")
    if not _wait_component_gone(wb, expected_name, timeout=rename_timeout):
        raise ModuleNameCollisionError(
            expected_name, actual,
            f"旧 '{expected_name}' が残ったままのため '{actual}' を改名できません")
    try:
        imported.Name = expected_name
    except Exception as ex:
        raise ModuleNameCollisionError(
            expected_name, actual, f"'{actual}' → '{expected_name}' の改名に失敗: {ex}")
    if imported.Name.lower() != expected_name.lower():
        raise ModuleNameCollisionError(
            expected_name, actual, f"改名後の実名が '{imported.Name}' のままです")
    print(f"回復成功: '{actual}' → '{expected_name}' に改名しました。")
    _reregister_shortcuts_from_bas(wb, import_path)
    return imported


def _print_collision_guidance(ex, module_name, module_backup, err=False):
    """名前衝突が自動回復できなかったときの案内（3経路共通・黙って成功にしない）"""
    out = sys.stderr if err else sys.stdout
    print(f"エラー: 名前衝突からの自動回復に失敗しました: {ex}", file=out)
    print(f"  ⚠ 置換後のコードは別名モジュール '{ex.actual_name}' として存在しています。", file=out)
    print(f"  ⚠ バックアップの再 Import はしないでください（VB_Name 衝突で連番モジュールが増えます）。", file=out)
    print(f"  対処: 旧 '{module_name}' が消えているのを確認してから、VBE のプロパティウィンドウで", file=out)
    print(f"        '{ex.actual_name}' の (オブジェクト名) を '{module_name}' に改名してください。", file=out)
    if module_backup:
        print(f"  置換前の内容のバックアップ: {module_backup}", file=out)


def _save_with_retry(wb, attempts=5, delay=0.6):
    """Import 直後の保存だけはビジー拒否で諦めない。

    Excel はセル編集中・メニュー展開中・モーダル表示中に COM 呼び出しを
    RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER / VBA_E_IGNORE で弾く
    （_com_is_busy 参照）。Remove+Import 直後にこれを食らうと、モジュールは
    正しく入れ替わっているのにブックだけ未保存で残る。一過性の拒否がほとんど
    なので、少し待って撃ち直す。恒久的な失敗（読み取り専用・ロック等）は
    そのまま送出して呼び出し側に判断させる。
    """
    last = None
    for i in range(attempts):
        try:
            wb.Save()
            return
        except Exception as ex:
            last = ex
            if not _com_is_busy(ex) or i == attempts - 1:
                raise
            print(f"  Excel がビジーのため保存を再試行します（{i + 1}/{attempts - 1}）...")
            time.sleep(delay)
            pythoncom.PumpWaitingMessages()
    if last:
        raise last


def _print_save_failed_guidance(module_name, err=False):
    """Import は成功したが、その後の『保存』で失敗したときの案内。

    2026-07-14 に発見した実害筋: Remove+Import 系3経路の except が removed
    フラグしか見ておらず、Import 成功後に wb.Save() が失敗すると「モジュールが
    消えた」と誤認してバックアップを再 Import していた。期待名のモジュールは
    既に正しく存在するので _wait_component_gone は空振りし（15秒）、VB_Name 衝突で
    連番別名として取り込まれ、改名待ちも空振りして（20秒）例外になる——つまり
    ツール自身が「旧コード入りの連番モジュール」を生んで 35 秒待たせていた。
    ここに来たら再 Import は絶対にしない。壊れているのは『保存』だけである。
    """
    out = sys.stderr if err else sys.stdout
    print(f"  ⚠ モジュール '{module_name}' の取り込みは成功しています"
          f"（開いているブックの中身は新しいコードに置き換わっています）。", file=out)
    print(f"  ⚠ 失敗したのは『保存』だけです。バックアップの再 Import はしません"
          f"（すると連番モジュールが増えるだけです）。", file=out)
    print(f"  対処: Excel のダイアログ・セル編集モードを解除してから、", file=out)
    print(f"        Excel で保存する（Ctrl+S）か、同じコマンドをもう一度実行してください。", file=out)
    print(f"  ⚠ 保存せずにブックを閉じると、この置換内容は失われます。", file=out)




# ================================================================
# 共有ヘルパー（分割時に後続パートから移設・前方参照解消）
# ================================================================

def check_vba_identifier(name):
    """VBA 識別子として無効なら理由（文字列）を返す。有効なら None。

    先頭 `_` の名前（例: `_tmp検証`）は VBE が黙って受け入れるがコンパイルで死ぬ。
    AddFromString / InsertLines は構文検査をしないため、注入自体は成功報告になり
    事故が繰り返された。VBA の識別子は英字か日本語などの文字で始まる必要があり、
    `_`・数字・記号では始められない。注入前にここで機械的に止める。
    """
    if not name:
        return "名前が空です"
    if name[0] == '_':
        suggestion = name.lstrip('_') or 'tmp'
        return (f"'{name}' は _ 始まりです。VBA の識別子は _ で始められません"
                f"（英字か日本語で始める。例: '{suggestion}'）")
    if name[0].isdigit():
        return f"'{name}' は数字始まりです。VBA の識別子は英字か日本語で始めてください"
    bad = re.sub(r'\w', '', name)
    if bad:
        return f"'{name}' に識別子に使えない文字が含まれています: {bad}"
    if len(name) > 255:
        return f"'{name}' が長すぎます（{len(name)}文字。VBA の上限は255文字）"
    return None

def _find_invalid_procedure_names(norm_text):
    """Sub/Function/Property 宣言の名前が VBA 識別子規則に反するものを列挙。

    コメント行は対象外。Declare 宣言は行頭トークンが合わないので元から素通り
    （外部 API 名は別規則のため対象にしない）。
    戻り値: [(行番号, 名前, 理由), ...]
    """
    decl_pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s\(]+)',
        re.IGNORECASE)
    hits = []
    for idx, line in enumerate(norm_text.split('\n'), 1):
        s = line.strip()
        if s.startswith("'") or s.lower().startswith('rem '):
            continue
        m = decl_pattern.match(line)
        if m:
            reason = check_vba_identifier(m.group(1))
            if reason:
                hits.append((idx, m.group(1), reason))
    return hits

def _col_letter(n):
    """列番号(1始まり)を A, B, ... Z, AA, ... に変換"""
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def _reject_extra_args(rest, used, usage):
    """位置引数の食い残しをエラーにする。

    黙って捨てると `clear-range Sheet1 A1:B2` のように「第2引数のつもりの範囲」が
    無視され、シート全域が対象になる事故（過去に実害）につながる。
    """
    if len(rest) > used:
        print(f"エラー: 余分な引数があります: {' '.join(str(a) for a in rest[used:])}")
        print(f"  {usage}")
        return True
    return False

def _cell_str(v):
    """セル値を表示用文字列に"""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime.datetime):
        # pywintypes の日時は str() だと "+00:00"（TZ）が付き、TSV 往復の
        # 書き戻しで日付がテキスト化する。素直な表記に整える
        # （_coerce_cell がこの表記を日付として復元する＝往復が一致する）
        if (v.hour, v.minute, v.second) == (0, 0, 0):
            return v.strftime('%Y/%m/%d')
        return v.strftime('%Y/%m/%d %H:%M:%S')
    return str(v)

def _range_values_2d(rng, use_formula=False):
    """Range の値を 2次元リストへ正規化（--tsv / --json 用）"""
    raw = rng.Formula if use_formula else rng.Value
    if raw is None:
        return [['']]
    if not isinstance(raw, tuple):
        return [[raw]]
    return [list(r) if isinstance(r, tuple) else [r] for r in raw]

_LAST_VALUES_FILE = os.path.join(SCRIPT_DIR, '_last_values.tsv')   # write-range のグリッド入力

_LAST_SNAPSHOT_FILE = os.path.join(SCRIPT_DIR, '_last_snapshot.json')  # snapshot の意味構造JSON出力先

def _coerce_cell(s):
    """文字列をセル値に変換。'='始まりは数式、数値は数値、空は None。

    数値化しても文字列に戻したい場合（郵便番号 "007" 等）は write-range --raw を使う。
    """
    if s is None or s == '':
        return None
    if s.startswith('='):
        return s                      # 数式 (.Value への代入で Excel が数式と解釈)
    # 日付表記（_cell_str の出力と同じ形）は datetime で書き戻す。
    # 文字列のまま .Value に入れるとテキスト格納になり、TSV 往復で日付列が壊れる
    m = re.fullmatch(
        r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})'
        r'(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?', s)
    if m:
        try:
            return datetime.datetime(
                int(m.group(1)), int(m.group(2)), int(m.group(3)),
                int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0))
        except ValueError:
            return s                  # 2026/99/99 のような非実在日付は文字列のまま
    if re.fullmatch(r'-?\d+', s):
        try:
            return int(s)
        except ValueError:
            return s
    # Python の float() は "1_000" や前後空白も受理してしまい、
    # "2026_07" のようなテキストIDが黙って数値化される。桁区切り表記や
    # 空白付きは文字列のまま扱う
    if '_' in s or s != s.strip():
        return s
    try:
        f = float(s)
    except ValueError:
        return s
    # "nan"/"inf" は float 化すると Excel 上でエラー値になるため文字列のまま
    if f != f or f in (float('inf'), float('-inf')):
        return s
    return f


# ================================================================
# 保護シート対策（2026-07-13）
#
# UserInterfaceOnly:=True で保護したシートでも、外部COM(このツール)からの
# ClearContents/Clear/CopyPicture・Shapes操作の一部は例外的にブロックされる
# （Value代入は素通りするが挙動が不統一）。VBAマクロ実行(run-macro)経由の
# 変更だけは正しく素通りするため、これは「マクロ実行か外部COMか」で
# UserInterfaceOnly の適用され方が違うことに起因する。
# 対策として、保護されたシートのある操作の前後で一時解除→元の設定で再保護する。
# ================================================================

# Protect()/Protection オブジェクトの真偽値フラグ（DrawingObjects/Contents/Scenarios/
# UserInterfaceOnly の4つ以外）。Worksheet.Protection.AllowXxx で現在値を読める。
# これを記録・復元しないと、並べ替え許可や行列挿入許可などブック固有のカスタム設定が
# 一時解除→再保護の往復で既定値(False=不許可)に巻き戻ってしまう。
_PROTECTION_ALLOW_FLAGS = (
    'AllowFormattingCells', 'AllowFormattingColumns', 'AllowFormattingRows',
    'AllowInsertingColumns', 'AllowInsertingRows', 'AllowInsertingHyperlinks',
    'AllowDeletingColumns', 'AllowDeletingRows',
    'AllowSorting', 'AllowFiltering', 'AllowUsingPivotTables',
)


def _unprotect_all_sheets(wb):
    """ブック内の保護されたシートを全て記録して一時解除する。

    戻り値は再保護用の (ws, drawing, contents, scenarios, ui_only, allow_kwargs) の
    リスト。allow_kwargs は並べ替え許可・行列挿入許可など細かい許可設定の辞書
    （Protect() にそのまま **allow_kwargs で渡せる）。
    パスワード保護等で解除できないシートは諦めて記録しない（そのシートは
    保護されたままなので、そのシートを触る操作は従来どおり失敗しうる）。
    """
    saved = []
    try:
        sheets = list(wb.Worksheets)
    except Exception:
        return saved
    for ws in sheets:
        try:
            protected = bool(ws.ProtectContents or ws.ProtectDrawingObjects
                              or ws.ProtectScenarios)
        except Exception:
            protected = False
        if not protected:
            continue
        try:
            drawing = bool(ws.ProtectDrawingObjects)
        except Exception:
            drawing = False
        try:
            contents = bool(ws.ProtectContents)
        except Exception:
            contents = False
        try:
            scenarios = bool(ws.ProtectScenarios)
        except Exception:
            scenarios = False
        try:
            ui_only = bool(ws.ProtectionMode)
        except Exception:
            ui_only = False
        allow_kwargs = {}
        try:
            prot = ws.Protection
            for flag in _PROTECTION_ALLOW_FLAGS:
                try:
                    allow_kwargs[flag] = bool(getattr(prot, flag))
                except Exception:
                    pass
        except Exception:
            pass
        try:
            ws.Unprotect()
            saved.append((ws, drawing, contents, scenarios, ui_only, allow_kwargs))
        except Exception as ex:
            # 解除できなかった＝そのシートは保護されたまま。以後の操作はそこで失敗するが、
            # 従来は完全に無言だったため「なぜ失敗したのか」が分からなかった
            print(f"⚠ シート '{_ws_name(ws)}' の保護を一時解除できませんでした"
                  f"（パスワード保護の可能性）: {ex}", file=sys.stderr)
            print("  このシートを触る操作は失敗します。Excel 側で先に保護を解除してください。",
                  file=sys.stderr)
    return saved


def _ws_name(ws):
    """シート名（取れなければ '?'）。エラーメッセージ用"""
    try:
        return ws.Name
    except Exception:
        return "?"


def _reprotect_sheets(saved):
    """_unprotect_all_sheets で記録した設定どおりに再保護する。

    再保護に失敗したら必ず報告する。黙って諦めると「ブックの保護が外れたまま
    コマンドは成功終了」になり、保護されているつもりのブックが無防備になる。
    """
    for ws, drawing, contents, scenarios, ui_only, allow_kwargs in saved:
        try:
            ws.Protect(DrawingObjects=drawing, Contents=contents,
                       Scenarios=scenarios, UserInterfaceOnly=ui_only,
                       **allow_kwargs)
        except Exception:
            try:
                # allow_kwargs の一部が今の Excel バージョンで受理されない場合の保険
                ws.Protect(DrawingObjects=drawing, Contents=contents,
                           Scenarios=scenarios, UserInterfaceOnly=ui_only)
                print(f"⚠ シート '{_ws_name(ws)}' を再保護しましたが、"
                      "細かい許可設定（並べ替え許可など）は復元できませんでした。",
                      file=sys.stderr)
            except Exception as ex:
                print(f"⚠ シート '{_ws_name(ws)}' の再保護に失敗しました: {ex}",
                      file=sys.stderr)
                print("  このシートは保護が外れたままです。Excel 側で保護し直してください。",
                      file=sys.stderr)


# 実行中の protected_sheets_guard（forget_protection から参照する）
_active_guards = []


def forget_protection(ws):
    """このシートを「ガードの再保護対象」から外す（保護解除そのものが目的の操作用）。

    ガードは入口で全保護シートを一時解除し、出口で記録どおり再保護する。
    そのため sheet unprotect のように「保護を外すこと自体が目的」の操作は、
    出口で保護を戻されて無言で効かなくなる。この関数で記録から落としておく。
    """
    try:
        name = ws.Name
    except Exception:
        return
    for g in _active_guards:
        kept = []
        for entry in g._saved:
            try:
                if entry[0].Name == name:
                    continue
            except Exception:
                pass
            kept.append(entry)
        g._saved = kept


class protected_sheets_guard:
    """保護されたシートのあるブックに対する外部COM操作を安全に行う with 文。

    シート保護に加えてブック構造保護（Protect Structure。シートの追加・削除・
    移動・改名を外部COMからもブロックする）も同じ流儀で一時解除→復元する。
    パスワード付きで解除できない場合は諦めて従来どおり（操作は失敗しうる）。

    使い方: with protected_sheets_guard(wb): ...操作...
    """
    def __init__(self, wb):
        self.wb = wb
        self._saved = []
        self._wb_saved = None  # (structure, windows) 解除できた場合のみ

    def __enter__(self):
        try:
            structure = bool(self.wb.ProtectStructure)
            windows = bool(self.wb.ProtectWindows)
        except Exception:
            structure = windows = False
        if structure or windows:
            try:
                self.wb.Unprotect()
                self._wb_saved = (structure, windows)
            except Exception:
                pass
        self._saved = _unprotect_all_sheets(self.wb)
        _active_guards.append(self)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            _active_guards.remove(self)
        except ValueError:
            pass
        _reprotect_sheets(self._saved)
        if self._wb_saved:
            try:
                self.wb.Protect(Structure=self._wb_saved[0],
                                Windows=self._wb_saved[1])
            except Exception:
                pass
        return False


def protect_safe(cmd_func):
    """cmd_* 関数を「対象ブックの保護シートを一時解除してから実行」に変える decorator。

    対象ブックは他の cmd_* 関数と同じ流儀（posargs先頭がExcelファイルなら
    それを対象、なければアクティブブック）で解決する。get_workbook は
    接続キャッシュを持つため、ここでの解決が二重コストにはならない。
    """
    import functools

    @functools.wraps(cmd_func)
    def wrapper(args):
        try:
            target_file, _rest = parse_target_and_rest(getattr(args, 'posargs', []) or [])
            xl, wb = get_workbook(target_file)
        except Exception:
            # ブック解決に失敗した場合は元の関数にそのまま委ね、
            # そちらのエラーメッセージを出させる
            return cmd_func(args)
        with protected_sheets_guard(wb):
            return cmd_func(args)
    return wrapper


__all__ = [
    'BACKUP_DIR',
    'LAST_PROC_FILE',
    'ModuleNameCollisionError',
    'SCRIPT_DIR',
    'XL_EXTS',
    '_LAST_SNAPSHOT_FILE',
    '_LAST_VALUES_FILE',
    '_cell_str',
    '_coerce_cell',
    '_col_letter',
    '_created_instances',
    '_created_xl',
    '_created_xl_pid',
    '_last_open_by_tool',
    '_find_component',
    '_find_invalid_procedure_names',
    '_get_active_excel',
    '_get_workbook_uncached',
    '_import_module_verified',
    '_print_collision_guidance',
    '_print_save_failed_guidance',
    '_save_with_retry',
    '_range_values_2d',
    '_reject_extra_args',
    '_remove_export_artifacts',
    '_running_excel_workbooks',
    '_wait_component_gone',
    '_wb_cache',
    '_unprotect_all_sheets',
    '_reprotect_sheets',
    '_active_guards',
    '_com_is_busy',
    '_pid_is_excel',
    '_ws_name',
    'forget_protection',
    'protected_sheets_guard',
    'protect_safe',
    'argparse',
    'check_vba_identifier',
    'cleanup_excel',
    'release_created_instances',
    'datetime',
    'get_workbook',
    'load_excel_addins_and_personal',
    'looks_like_xl_file',
    'make_backup',
    'make_module_backup',
    'normalize_bas_newlines',
    'os',
    'parse_target_and_rest',
    'pythoncom',
    'pywintypes',
    're',
    'read_code_file',
    'setup_encoding',
    'shutil',
    'smart_path_resolve',
    'sys',
    'time',
    'unicodedata',
    'validate_bas_encoding',
    'validate_vba_code',
    'win32com',
    'zlib',
]
