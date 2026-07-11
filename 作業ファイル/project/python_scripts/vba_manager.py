"""
VBAマネージャー (アクティブブック対応版)

【特徴】
- target_file を省略するとアクティブなExcelブックを自動使用
- get で取得したコードは _last_proc.vba に保存 → Claudeが読み取り・修正
- replace-procedure は _last_proc.vba を自動使用 (--code-file 省略時)
- replace-module は Remove+Import で Attribute を正しく処理

【コマンド一覧】
  list            [excel_file]                     マクロ一覧
  list-modules    [excel_file]                     モジュール一覧
  get             [excel_file] <macro_name>        プロシージャのコード取得
  replace-procedure [excel_file] [--code-file f]  プロシージャを置換
  replace-module  [excel_file] <module> <bas_file> モジュール全体を置換
  export-module   [excel_file] <module>            モジュールを .bas にエクスポート
  diag                                             動作確認

  reorder-macro   <macro> <up|down>                マクロの表示順を入れ替え
  list-shortcuts  [excel_file]                      ショートカットキー一覧

【目コマンド（シート状態の読み取り）】
  read-range     [excel_file] [range] [--formula]  セル値（--formulaで数式）をテキスト格子で読む
  read-selection [excel_file] [--formula]           今選択している範囲を読む
  sheet-info     [excel_file]                       シート構成・使用範囲の一覧
  screenshot     [excel_file] [range] [--out f]    範囲を画像(PNG)で書き出す

【手コマンド（シートの編集・整形・構造操作／開いたままのブックに直接書込）】
  write-range    [excel_file] <range> [値]          値・数式を書込（グリッドは --tsv / _last_values.tsv）
  clear-range    [excel_file] <range>               範囲をクリア（--contents/--formats/--all）
  format-range   [excel_file] <range> [書式opt...]  フォント・色・罫線・書式・列幅等
  sheet          <add|delete|rename|copy|activate|show|hide|very-hide|visibility|tab-color>
  table          <create|list|delete|column|filter|filter-values|filter-clear|filters|sort|sort-multi|ref>
  name           [excel_file] <add|list|delete>     名前付き範囲

  -- 編集の足回り --
  row            <insert|delete> <行番号> [本数]    行の挿入・削除
  col            <insert|delete> <列文字> [本数]    列の挿入・削除
  copy-range     <src> <dst> [--values]            範囲コピー
  fill           <range> [--right]                  オートフィル（既定は下）
  sort           <range> [--key 列][--desc][--header] 並べ替え
  autofilter     [range] [--off]                    オートフィルタ
  -- 検索・置換 --
  find           <文字> [--book][--whole][--formula] セル検索（番地を返す）
  find-replace   <検索> <置換> [range] [--whole]     一括置換
  -- 保存・印刷 --
  save           [excel_file]                       上書き保存
  save-as        <path>                             別名保存
  print-setup    [--area R][--title-rows 1:3]...    印刷設定
  -- 仕上げ --
  cond-format    <range> --gt 100 --bg '#...'        条件付き書式
  hyperlink      <cell> <url> [--text t]             ハイパーリンク
  validation     <range> --list 'A,B,C'              入力規則(ドロップダウン)
  freeze         <cell> | off                        ウィンドウ枠固定
  comment        <cell> <text>                       セルコメント
  -- 重量級 --
  chart          <create|list|delete>                グラフ（column/bar/line/pie/scatter/area）
  chart-config   <set-title|set-type|legend|style|axis-scale|data-labels|add-series|trendline...>  グラフ詳細設定
  pivot          <create|list|delete>                ピボットテーブル（--rows/--cols/--values/--func）
  pivot-field    <list|add-row|add-col|add-value|remove|set-func|sort|group-date|group-numeric...>  フィールド管理
  pivot-calc     <get-data|calc-field|layout|subtotals|grand-totals>  計算フィールド・レイアウト
  slicer         <add|list|delete>                   スライサー（ピボット/テーブルに紐づけ）
  calc-mode      [manual|auto|recalc]                計算モード確認・切替・再計算
  powerquery     <list|refresh|add|edit|delete|load>  PowerQueryの一覧・更新・作成・書換・削除・読込配線(--to sheet|model)
  connection     <list|refresh|delete> [name]        ブック接続の一覧・更新・削除
  datamodel      <list|relation|measure>             データモデル一覧／リレーション・メジャー(DAX)の作成削除
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

        # 新規起動したPIDが存在する場合は強制クリーンアップ
        if pid is not None:
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
                print("⚠ このブックは診断用に読み取り専用で開いています。"
                      "編集するには Excel で通常モードで開き直してください。")
            if load_addins:
                load_excel_addins_and_personal(xl)
            return xl, wb
        except Exception:
            del _wb_cache[key]
    xl, wb = _get_workbook_uncached(target_file_arg, load_addins, readonly)
    # readonly フラグは「このツールが自動で開いた場合」だけ意味を持つ
    opened_ro = readonly and target_file_arg is not None
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
    """
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
        print(f"対象ブック: {wb.Name}  (新規オープン・読み取り専用・イベント無効=健診モード)")
        return xl, wb

    wb = xl.Workbooks.Open(target_path)
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


# ================================================================
# コマンド実装
# ================================================================

def _find_duplicate_procedures(norm_text):
    """.bas 内の Sub/Function 名の重複を機械的に検出（重複プロシージャ挿入の検知）。

    Property Get/Let/Set は同名が正常なので対象外（Sub/Function のみ）。
    戻り値: {名前: [行番号, ...], ...}（重複のあるものだけ）
    """
    sub_pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?P<kind>Sub|Function)\s+(?P<name>[^\s\(]+)',
        re.IGNORECASE
    )
    seen = {}
    for idx, line in enumerate(norm_text.split('\n'), 1):
        m = sub_pattern.match(line)
        if m:
            seen.setdefault(m.group('name'), []).append(idx)
    return {name: lns for name, lns in seen.items() if len(lns) > 1}


def _find_consecutive_dup_lines(norm_text):
    """連続して同一の非空コード行を検出（重複挿入の臭い。例: On Error Resume Next ×2）。

    空行・コメント行は対象外。空行を挟むとリセット（空行連続は正常）。
    入れ子で正常に連続しうるブロック終端等（End If / End With / Next / Loop / Else / Wend）は
    重複扱いしない（実モジュールでの誤検知を避ける）。
    戻り値: [(行番号, 行内容), ...]
    """
    struct = re.compile(r'^(end\b|else\b|elseif\b|next\b|loop\b|wend\b)', re.IGNORECASE)
    hits = []
    prev = None
    for idx, raw in enumerate(norm_text.split('\n'), 1):
        s = raw.strip()
        if s and not s.startswith("'") and s == prev and not struct.match(s):
            hits.append((idx, s))
        prev = s if s else None
    return hits


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


def cmd_check_bas(args):
    """.bas を取り込む前の単体検査（COM不要）。複数ファイル可。

    バイパス経路（vba_manager を通さず手書きスクリプトで .bas を作る）でも、取り込み前に
    この1コマンドで「文字コード事故 / 改行二重化 / プロシージャ重複 / 識別子規則違反 /
    連続重複行」を機械的に検査できる。COM接続が落ちていても動くのが要点（安全確認を不安全な手順と同じ手数にする）。
    --fix を付けると改行二重化だけ CP932 のまま自動修正する
    （重複は判断が要るので自動修正しない＝Pythonは機械的検査まで）。
    """
    if not args.posargs:
        print("使い方: py vba_manager.py check-bas <file.bas> [file2.bas ...] [--fix] [--json]")
        return False
    results = []
    ok_all = True
    for p in args.posargs:
        ok = _check_bas_one(p, fix=getattr(args, 'fix', False))
        results.append({"file": p, "ok": bool(ok)})
        if not ok:
            ok_all = False
    if len(args.posargs) > 1:
        print(f"===== 一括検査: {len(args.posargs)}本 → {'すべて取り込み可' if ok_all else '⚠ NGあり'} =====")
    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": ok_all, "files": results}, ensure_ascii=False),
              file=sys.stdout)
    return ok_all


def _check_bas_one(path, fix=False):
    """check-bas の1ファイルぶんの検査本体（True=取り込み可）"""
    if not os.path.isfile(path):
        print(f"エラー: ファイルが見つかりません: {path}")
        return False

    name = os.path.basename(path)
    print(f"===== .bas 取り込み前検査: {name} =====")
    problems = 0
    warnings = 0

    # 1. 文字コード事故（UTF-8化 / BOM）
    if not validate_bas_encoding(path):
        problems += 1
    else:
        print("  [OK] 文字コード: CP932 として安全")

    # 2. 改行二重化（\r\r\n）
    try:
        fixed_bytes, raw_bytes, was_doubled = normalize_bas_newlines(path)
    except Exception as e:
        print(f"エラー: 改行検査に失敗 ({e})")
        return False
    if was_doubled:
        before = len(re.split(r'\r\n|\r|\n', raw_bytes.decode('cp932')))
        after = len(re.split(r'\r\n|\r|\n', fixed_bytes.decode('cp932')))
        if fix:
            # in-place 書き換えなので、書く前に元バイト列を退避しておく（undo 導線）
            try:
                bak = path + f".bak_{time.strftime('%Y%m%d_%H%M%S')}"
                with open(bak, 'wb') as f:
                    f.write(raw_bytes)
                print(f"  退避: {bak}")
            except Exception as e:
                print(f"  [WARN] 退避に失敗しました（{e}）。修正は続行します")
            with open(path, 'wb') as f:
                f.write(fixed_bytes)
            print(f"  [FIXED] 改行二重化を修正しました: {before}行 → {after}行")
        else:
            print(f"  [NG] 改行二重化を検知: {before}行 → {after}行（--fix で修正可）")
            problems += 1
    else:
        print("  [OK] 改行: 正規 CRLF（二重化なし）")

    # 3/4 の検査は現在のファイル内容（--fix 後を反映）に対して行う
    with open(path, 'rb') as f:
        norm_text = re.sub(r'\r\n|\r', '\n', f.read().decode('cp932'))

    # 3. プロシージャ名の重複（重複挿入の検知・自動修正しない）
    dups = _find_duplicate_procedures(norm_text)
    if dups:
        print("  [NG] Sub/Function 名の重複を検知（重複挿入の疑い・自動修正しません）:")
        for nm, lns in dups.items():
            print(f"        {nm}  (行 {', '.join(map(str, lns))})")
        problems += 1
    else:
        print("  [OK] プロシージャ名: 重複なし")

    # 3b. プロシージャ名の識別子規則（先頭 _ 等。VBE は黙って受け入れコンパイルで死ぬ）
    bad_names = _find_invalid_procedure_names(norm_text)
    if bad_names:
        print("  [NG] VBA の識別子規則に反するプロシージャ名を検知（コンパイルエラーになります）:")
        for ln, _nm, reason in bad_names:
            print(f"        行{ln}: {reason}")
        problems += 1
    else:
        print("  [OK] プロシージャ名: 識別子規則OK")

    # 4. 連続する同一コード行（On Error Resume Next ×2 等の臭い）
    cdl = _find_consecutive_dup_lines(norm_text)
    if cdl:
        print("  [WARN] 連続する同一コード行（重複挿入の臭い・要確認）:")
        for ln, s in cdl[:20]:
            disp = s if len(s) <= 60 else s[:60] + '…'
            print(f"        行{ln}: {disp}")
        if len(cdl) > 20:
            print(f"        … 他 {len(cdl) - 20} 件")
        warnings += 1
    else:
        print("  [OK] 連続重複行: なし")

    print(f"----- 結果: 問題 {problems} / 警告 {warnings} -----")
    if problems:
        print("  ⚠ 問題があります。修正してから replace-module / replace-procedure で取り込んでください。")
        return False
    print("  取り込み可（手書きでも、取り込みは replace-module / replace-procedure 経由を推奨）。")
    return True


def cmd_check(args):
    """VBAコードの静的解析・診断を行う"""
    target_file, _ = parse_target_and_rest(args.posargs)
    # 診断系なので readonly（閉じているブックを通常モードで開くと
    # Workbook_Open / Auto_Open が発火して副作用が走る）
    xl, wb = get_workbook(target_file, readonly=True)
    
    is_json = getattr(args, 'json', False)

    if not is_json:
        print(f"\n===== VBA診断を実行中: {wb.Name} =====")
    
    results = {
        "success": True,
        "file": wb.Name,
        "modules": [],
        "duplicates": [],
        "summary": {"errors": 0, "warnings": 0}
    }
    
    all_procedures = {}  # proc_name -> [module_name, ...]
    
    for comp in wb.VBProject.VBComponents:
        comp_name = comp.Name
        cm = comp.CodeModule
        count_lines = cm.CountOfLines
        
        type_names = {1: '標準モジュール', 2: 'クラスモジュール',
                      3: 'フォーム', 100: 'シート/ThisWorkbook'}
        tname = type_names.get(comp.Type, f'Type={comp.Type}')
        
        mod_info = {
            "name": comp_name,
            "type": tname,
            "type_id": comp.Type,
            "warnings": [],
            "errors": [],
            "skipped": False
        }
        
        if count_lines == 0:
            mod_info["skipped"] = True
            results["modules"].append(mod_info)
            continue
            
        code = cm.Lines(1, count_lines)
        code = code.replace('\r\n', '\n').replace('\r', '\n')
        lines = code.split('\n')
        
        # 1. Option Explicit チェック
        has_option_explicit = False
        for line in lines:
            stripped = line.strip().lower()
            if not stripped:
                continue
            if stripped.startswith("'") or stripped.startswith("rem "):
                continue
            if stripped.startswith("option explicit"):
                has_option_explicit = True
                break
            if re.match(r'^(?:(?:public|private|friend)\s+)?(?:static\s+)?(?:sub|function|property)\s+', stripped) or stripped.startswith("dim ") or stripped.startswith("const "):
                break
        
        if not has_option_explicit:
            mod_info["warnings"].append("Option Explicit が記述されていません。変数宣言の強制を推奨します。")
            
        # 2. Sub/Function 閉じ忘れチェック
        decl_sub = 0
        end_sub = 0
        decl_func = 0
        end_func = 0
        
        proc_pattern = re.compile(
            r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
            r'(Sub|Function)\s+([^\s\(\)]+)',
            re.IGNORECASE
        )
        
        local_variables = []  # (var_name, line_idx)
        variable_usage = {}
        
        # プロシージャごとの詳細診断用状態管理
        current_proc_name = None
        current_proc_start_idx = None
        current_proc_has_error_handler = False
        current_proc_kind = None  # 'sub' or 'function'

        for idx, line in enumerate(lines):
            # 文字列リテラルを潰してからコメントを落とす。先に ' で切ると
            # 文字列内のアポストロフィ（"Don't" 等）で行が切断され、
            # 同じ行の End Sub を見失って正常コードが閉じ忘れ ERROR になる
            clean_line = re.sub(r'"[^"]*"', '""', line).split("'")[0].strip()
            if clean_line.lower().startswith("rem "):
                continue

            m = proc_pattern.match(clean_line)
            if m:
                # 別のプロシージャの中にいる状態で新しい宣言を見つけた場合（前のプロシージャがEnd Subなしで閉じた等）、
                # 簡易クリア（閉じ忘れ警告は後続 of if decl_sub != end_sub で処理）
                kind = m.group(1).lower()
                name = m.group(2)
                current_proc_name = name
                current_proc_start_idx = idx
                current_proc_has_error_handler = False
                current_proc_kind = kind

                if name not in all_procedures:
                    all_procedures[name] = []
                all_procedures[name].append(comp_name)

                if kind == 'sub':
                    decl_sub += 1
                elif kind == 'function':
                    decl_func += 1

            # プロシージャ内における警告チェック
            if current_proc_name:
                # On Error の検出
                if "on error " in clean_line.lower():
                    current_proc_has_error_handler = True
                # SendKeys の検出
                if "sendkeys" in clean_line.lower():
                    mod_info["warnings"].append(f"プロシージャ '{current_proc_name}' 内で危険な SendKeys が使用されています (行 {idx + 1})")

            # 終了チェック。「Sub x(): End Sub」の1行書きは End が行頭に来ないため、
            # validate_vba_code と同じく ':' で文に分割してから数える
            # （文字列内の ':' で誤分割しないよう先に潰す）
            blanked = re.sub(r'"[^"]*"', '""', clean_line)
            segs = [s.strip() for s in blanked.split(':')]
            n_end_sub = sum(1 for s in segs
                            if re.match(r'^end\s+sub\b', s, re.IGNORECASE))
            n_end_func = sum(1 for s in segs
                             if re.match(r'^end\s+function\b', s, re.IGNORECASE))
            is_end_sub = n_end_sub > 0
            is_end_func = n_end_func > 0

            end_sub += n_end_sub
            end_func += n_end_func

            if current_proc_name and (
                (current_proc_kind == 'sub' and is_end_sub) or
                (current_proc_kind == 'function' and is_end_func)
            ):
                if not current_proc_has_error_handler:
                    mod_info["warnings"].append(f"プロシージャ '{current_proc_name}' にエラーハンドリング (On Error) がありません (行 {current_proc_start_idx + 1})")

                current_proc_name = None
                current_proc_start_idx = None
                current_proc_kind = None

            # Dim宣言の簡易スキャン
            dim_match = re.match(r'^\s*Dim\s+(.+)$', clean_line, re.IGNORECASE)
            if dim_match:
                dim_body = dim_match.group(1)
                parts = dim_body.split(',')
                for p in parts:
                    p = p.strip()
                    var_part = re.split(r'\s+As\s+', p, flags=re.IGNORECASE)[0].strip()
                    var_name = re.sub(r'\(.*\)', '', var_part).strip()
                    var_name = re.sub(r'[%&\$#!@]$', '', var_name).strip()
                    if var_name and re.match(r'^[a-zA-Z_][a-zA-Z0-9_]*$', var_name):
                        local_variables.append((var_name, idx))
                        variable_usage[var_name] = 0

        if decl_sub != end_sub:
            mod_info["errors"].append(f"Sub の閉じ忘れがあります (宣言数: {decl_sub}, End Sub数: {end_sub})")
        if decl_func != end_func:
            mod_info["errors"].append(f"Function の閉じ忘れがあります (宣言数: {decl_func}, End Function数: {end_func})")

        # 未使用変数のカウントスキャン
        if local_variables:
            for idx, line in enumerate(lines):
                # 上と同じ理由で、文字列を潰してからコメントを落とす
                clean_line = re.sub(r'"[^"]*"', '""', line).split("'")[0]
                for var_name, decl_idx in local_variables:
                    if idx == decl_idx:
                        continue
                    if re.search(r'\b' + re.escape(var_name) + r'\b', clean_line, re.IGNORECASE):
                        variable_usage[var_name] += 1

            for var_name, decl_idx in local_variables:
                if variable_usage[var_name] == 0:
                    mod_info["warnings"].append(f"未使用変数 Dim {var_name} があります (行 {decl_idx + 1})")

        results["modules"].append(mod_info)

        # 画面出力 (JSON指定でない場合のみ)
        if not is_json:
            print(f"\n📄 モジュール: {comp_name} ({tname})")
            if not has_option_explicit:
                print("  [WARNING] Option Explicit が記述されていません。変数宣言の強制を推奨します。")
            if mod_info["errors"]:
                for err in mod_info["errors"]:
                    print(f"  [ERROR] {err}")
            if mod_info["warnings"]:
                for warn in mod_info["warnings"]:
                    # Option Explicitの警告はすでに出力しているのでスキップ
                    if "Option Explicit" in warn:
                        continue
                    print(f"  [WARNING] {warn}")
            print("  モジュール診断完了")

    # 重複チェックの集計
    for proc_name, mods in all_procedures.items():
        if len(mods) > 1:
            results["duplicates"].append({
                "procedure": proc_name,
                "modules": mods
            })
            
    # サマリーの集計
    err_total = sum(len(m["errors"]) for m in results["modules"])
    warn_total = sum(len(m["warnings"]) for m in results["modules"]) + len(results["duplicates"])
    results["summary"]["errors"] = err_total
    results["summary"]["warnings"] = warn_total

    # JSON出力
    if is_json:
        import json
        print(json.dumps(results, ensure_ascii=False), file=sys.stdout)
        return err_total == 0

    # 通常出力
    print("\n===== ブック全体の重複診断 =====")
    if results["duplicates"]:
        for dup in results["duplicates"]:
            print(f"  [WARNING] 重複プロシージャ名 '{dup['procedure']}' が複数のモジュールに存在します:")
            for m in dup["modules"]:
                print(f"    - {m}")
    else:
        print("  プロシージャ名の重複はありません。")

    print(f"\n===== 診断サマリー =====")
    print(f"  エラー数  : {err_total}")
    print(f"  警告数    : {warn_total}")
    
    if err_total > 0:
        print("  [RESULT] 致命的な構文エラーがあります。修正してください。")
        return False
    elif warn_total > 0:
        print("  [RESULT] 警告項目がありますが、実行は可能です。品質向上のため修正を推奨します。")
        return True
    else:
        print("  [RESULT] すべてのチェックを通過しました。良好な状態です。")
        return True


def cmd_diag(args):
    """動作確認"""
    print("Syntax OK")
    try:
        xl = _get_active_excel()
        wb = xl.ActiveWorkbook
        if wb:
            print(f"アクティブブック: {wb.Name}")
        else:
            print("アクティブブック: なし")
    except Exception:
        print("Excelは起動していません")
    return True

def cmd_list_open(args):
    """現在開いているExcelファイルを一覧表示"""
    as_json = getattr(args, 'json', False)
    books = []
    seen = set()

    def _add(wb):
        try:
            full = wb.FullName
            name = wb.Name
        except Exception:
            return
        if full.lower() in seen:
            return
        seen.add(full.lower())
        books.append({"name": name, "fullname": full})

    # ROT 全走査: GetActiveObject は ROT 先頭の1インスタンスしか返さず、
    # 非表示ゾンビ(ブック0個)を掴んで「ブックなし」と誤報することがある。
    # 点呼コマンドこそ全インスタンスの全ブックを数える必要がある
    for wb in _running_excel_workbooks():
        _add(wb)
    excel_running = bool(books)
    # 未保存ブック(Book1等)はパスを持たず ROT に載らないことがあるため、
    # GetActiveObject 側の列挙でも補完する
    try:
        xl = _get_active_excel()
        excel_running = True
        for wb in xl.Workbooks:
            _add(wb)
    except Exception:
        pass

    if as_json:
        import json
        print(json.dumps({"success": True, "excel_running": excel_running,
                          "workbooks": books}, ensure_ascii=False), file=sys.stdout)
        return True
    if not excel_running:
        print('Excelは起動していません')
        return
    if not books:
        print('（開いているブックはありません）')
        return True
    for b in books:
        print(b["fullname"])
    return True



def _select_addin_project(all_projects, sel):
    """--addin の対象アドインを選ぶ。sel は True（無指定）または名前の一部。

    特定のアドイン名を優先するハードコードはしない（汎用原則）。
    複数ロード時は名前指定を促す。見つからなければ None（メッセージ出力済み）。
    """
    found = []
    for p in all_projects:
        try:
            fname = os.path.basename(p.Filename).lower()
            if fname.endswith(('.xlam', '.xla')):
                found.append(p)
        except Exception:
            continue
    if not found:
        print("エラー: アドインブック (.xlam / .xla) がロードされていません。", file=sys.stderr)
        return None
    if isinstance(sel, str):
        hits = [p for p in found
                if sel.lower() in os.path.basename(p.Filename).lower()]
        if not hits:
            names = ", ".join(os.path.basename(p.Filename) for p in found)
            print(f"エラー: '{sel}' に一致するアドインがありません。ロード中: {names}",
                  file=sys.stderr)
            return None
        if len(hits) > 1:
            names = ", ".join(os.path.basename(p.Filename) for p in hits)
            print(f"エラー: '{sel}' に複数一致します: {names}", file=sys.stderr)
            return None
        return hits[0]
    if len(found) > 1:
        names = ", ".join(os.path.basename(p.Filename) for p in found)
        print("エラー: 複数のアドインがロードされています。"
              "--addin 名前 で対象を指定してください。", file=sys.stderr)
        print(f"  ロード中: {names}", file=sys.stderr)
        return None
    return found[0]


def cmd_list(args):
    """マクロ(プロシージャ)一覧"""
    load_addins = getattr(args, 'personal', False) or getattr(args, 'addin', False) or getattr(args, 'all', False)
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, default_wb = get_workbook(target_file, load_addins=load_addins)

    # 全プロジェクトをリスト化 (ビジーエラー対策としてリトライ)
    import time
    all_projects = []
    for attempt in range(5):
        try:
            all_projects = []
            for p in xl.VBE.VBProjects:
                all_projects.append(p)
            break
        except Exception as ex:
            if "800ac472" in str(ex) and attempt < 4:
                time.sleep(0.5)
                continue
            print(f"エラー: VBAプロジェクトモデルへのアクセスが拒否されました: {ex}", file=sys.stderr)
            return False

    # 対象のプロジェクトを選択
    target_projects = []
    if getattr(args, 'all', False):
        target_projects = all_projects
    elif getattr(args, 'personal', False):
        found = None
        for p in all_projects:
            try:
                fname = os.path.basename(p.Filename).lower()
                if fname in ("personal.xlsb", "personal.xls"):
                    found = p
                    break
            except Exception:
                continue
        if not found:
            print("エラー: 個人用マクロブック (PERSONAL.XLSB) がロードされていません。", file=sys.stderr)
            return False
        target_projects.append(found)
    elif getattr(args, 'addin', False):
        target_addin = _select_addin_project(all_projects, args.addin)
        if target_addin is None:
            return False
        target_projects.append(target_addin)
    else:
        found = None
        try:
            for p in all_projects:
                try:
                    if p.Filename.lower() == default_wb.FullName.lower():
                        found = p
                        break
                except Exception:
                    continue
        except Exception:
            pass
        if not found:
            try:
                found = default_wb.VBProject
            except Exception as ex:
                print(f"エラー: VBProjectの取得に失敗しました: {ex}", file=sys.stderr)
                return False
        target_projects.append(found)

    pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE
    )

    def get_project_display_name(p):
        try:
            if p.Filename:
                return os.path.basename(p.Filename)
        except Exception:
            pass
        return p.Name

    is_all = getattr(args, 'all', False)
    if is_all:
        all_results = {}
        for proj in target_projects:
            macros = []
            proj_name = get_project_display_name(proj)
            try:
                for comp in proj.VBComponents:
                    if getattr(args, 'standard', False) and comp.Type != 1:
                        continue
                    cm = comp.CodeModule
                    if cm.CountOfLines == 0:
                        continue
                    for m in pattern.finditer(cm.Lines(1, cm.CountOfLines)):
                        name = m.group(1)
                        if name not in macros:
                            macros.append(name)
            except Exception as ex:
                print(f"[DEBUG] Failed to access VBComponents of {proj_name}: {ex}", file=sys.stderr)
                continue
            all_results[proj_name] = macros

        if getattr(args, 'json', False):
            import json
            print(json.dumps({"success": True, "file": "all", "macros": all_results}, ensure_ascii=False), file=sys.stdout)
            return True

        for p_name, macros in all_results.items():
            print(f"\n--- {p_name} ---")
            print(f"マクロ数: {len(macros)}")
            for name in macros:
                print(f"MACRO:{name}")
        return True

    proj = target_projects[0]
    proj_name = get_project_display_name(proj)
    mod_filter = getattr(args, 'module_opt', None)
    detail = getattr(args, 'detail', False)
    macros = []            # 従来互換の名前リスト
    details = []           # --detail / --json 用
    try:
        for comp in proj.VBComponents:
            if getattr(args, 'standard', False) and comp.Type != 1:
                continue
            if mod_filter and comp.Name.lower() != mod_filter.lower():
                continue
            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                continue
            code = cm.Lines(1, cm.CountOfLines)
            lines = code.split('\r\n')
            for m in pattern.finditer(code):
                name = m.group(1)
                if name in macros:
                    continue
                macros.append(name)
                if not (detail or getattr(args, 'json', False)):
                    continue
                info = {'module': comp.Name, 'name': name}
                try:
                    info['lines'] = cm.ProcCountLines(name, 0)
                    start = cm.ProcStartLine(name, 0)
                    body_start = cm.ProcBodyLine(name, 0)
                    # 宣言の次行が先頭コメントならそれを1行だけ添える（機械的抽出）
                    if body_start < len(lines):
                        first = lines[body_start].strip()   # body_start は1始まり＝宣言行、次行は index body_start
                        if first.startswith("'"):
                            info['comment'] = first.lstrip("'").strip()
                except Exception:
                    pass
                details.append(info)
    except Exception as ex:
        print(f"エラー: VBComponentsへのアクセスに失敗しました: {ex}", file=sys.stderr)
        return False

    if getattr(args, 'json', False):
        import json
        payload = {"success": True, "file": proj_name, "macros": macros}
        if details:
            payload["details"] = details
        print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
        return True

    print(f"対象ブック: {proj_name}")
    print(f"マクロ数: {len(macros)}")
    if detail:
        for d in details:
            extra = f", {d['lines']}行" if 'lines' in d else ""
            cmt = f"  '{d['comment']}" if 'comment' in d else ""
            print(f"MACRO:[{d['module']}] {d['name']}{extra}{cmt}")
    else:
        for name in macros:
            print(f"MACRO:{name}")
    return True


def cmd_list_modules(args):
    """モジュール一覧"""
    load_addins = getattr(args, 'personal', False) or getattr(args, 'addin', False) or getattr(args, 'all', False)
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, default_wb = get_workbook(target_file, load_addins=load_addins)

    # 全プロジェクトをリスト化 (ビジーエラー対策としてリトライ)
    import time
    all_projects = []
    for attempt in range(5):
        try:
            all_projects = []
            for p in xl.VBE.VBProjects:
                all_projects.append(p)
            break
        except Exception as ex:
            if "800ac472" in str(ex) and attempt < 4:
                time.sleep(0.5)
                continue
            print(f"エラー: VBAプロジェクトモデルへのアクセスが拒否されました: {ex}", file=sys.stderr)
            return False

    # 対象のプロジェクトを選択
    target_projects = []
    if getattr(args, 'all', False):
        target_projects = all_projects
    elif getattr(args, 'personal', False):
        found = None
        for p in all_projects:
            try:
                fname = os.path.basename(p.Filename).lower()
                if fname in ("personal.xlsb", "personal.xls"):
                    found = p
                    break
            except Exception:
                continue
        if not found:
            print("エラー: 個人用マクロブック (PERSONAL.XLSB) がロードされていません。", file=sys.stderr)
            return False
        target_projects.append(found)
    elif getattr(args, 'addin', False):
        target_addin = _select_addin_project(all_projects, args.addin)
        if target_addin is None:
            return False
        target_projects.append(target_addin)
    else:
        found = None
        try:
            for p in all_projects:
                try:
                    if p.Filename.lower() == default_wb.FullName.lower():
                        found = p
                        break
                except Exception:
                    continue
        except Exception:
            pass
        if not found:
            try:
                found = default_wb.VBProject
            except Exception as ex:
                print(f"エラー: VBProjectの取得に失敗しました: {ex}", file=sys.stderr)
                return False
        target_projects.append(found)

    type_names = {1: '標準モジュール', 2: 'クラスモジュール',
                  3: 'フォーム', 100: 'シート/ThisWorkbook'}

    def get_project_display_name(p):
        try:
            if p.Filename:
                return os.path.basename(p.Filename)
        except Exception:
            pass
        return p.Name

    is_all = getattr(args, 'all', False)
    if is_all:
        all_results = {}
        for proj in target_projects:
            modules = []
            proj_name = get_project_display_name(proj)
            try:
                for comp in proj.VBComponents:
                    tname = type_names.get(comp.Type, f'Type={comp.Type}')
                    modules.append({"name": comp.Name, "type": tname, "type_id": comp.Type})
            except Exception as ex:
                print(f"[DEBUG] Failed to access VBComponents of {proj_name}: {ex}", file=sys.stderr)
                continue
            all_results[proj_name] = modules

        if getattr(args, 'json', False):
            import json
            print(json.dumps({"success": True, "file": "all", "modules": all_results}, ensure_ascii=False), file=sys.stdout)
            return True

        for p_name, modules in all_results.items():
            print(f"\n--- {p_name} ---")
            for m in modules:
                print(f"MODULE:{m['name']}  ({m['type']})")
        return True

    proj = target_projects[0]
    proj_name = get_project_display_name(proj)
    modules = []
    try:
        for comp in proj.VBComponents:
            tname = type_names.get(comp.Type, f'Type={comp.Type}')
            modules.append({"name": comp.Name, "type": tname, "type_id": comp.Type})
    except Exception as ex:
        print(f"エラー: VBComponentsへのアクセスに失敗しました: {ex}", file=sys.stderr)
        return False

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": True, "file": proj_name, "modules": modules}, ensure_ascii=False), file=sys.stdout)
        return True

    print(f"対象ブック: {proj_name}")
    for m in modules:
        print(f"MODULE:{m['name']}  ({m['type']})")
    return True


def _all_procedure_names(wb):
    """ブック内の全プロシージャ名を列挙（did-you-mean 用の機械的リスト）"""
    pat = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE)
    names = []
    try:
        for comp in wb.VBProject.VBComponents:
            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                continue
            for m in pat.finditer(cm.Lines(1, cm.CountOfLines)):
                if m.group(1) not in names:
                    names.append(m.group(1))
    except Exception:
        pass
    return names


def _suggest_similar(name, candidates, label="もしかして"):
    """タイポ候補の提示（difflib による機械的な近似のみ・判断はしない）"""
    import difflib
    close = difflib.get_close_matches(name, candidates, n=3, cutoff=0.6)
    if close:
        print(f"  {label}: {' / '.join(close)}")
    print("  py vba_manager.py list で一覧を確認できます。")


def _extract_proc(wb, module_name, macro_name):
    """1プロシージャのコードを取り出す。

    戻り値: (comp_name, clean_code) / 見つからなければ (None, None)。
    同名複数（module_name 未指定時）は例外 ValueError(候補リスト) を投げる。
    """
    # モジュール未指定時：同名プロシージャが複数モジュールにある場合はエラー
    # （違うフォームの同名イベントを黙って掴む事故を防ぐ。replace-procedure と同じ流儀）
    if not module_name:
        matched = []
        for comp in wb.VBProject.VBComponents:
            try:
                comp.CodeModule.ProcStartLine(macro_name, 0)
                matched.append(comp.Name)
            except Exception:
                pass
        if len(matched) > 1:
            raise ValueError(matched)

    for comp in wb.VBProject.VBComponents:
        if module_name and comp.Name.lower() != module_name.lower():
            continue
        cm = comp.CodeModule
        try:
            proc_start = cm.ProcStartLine(macro_name, 0)
            count      = cm.ProcCountLines(macro_name, 0)
            # ProcStartLine 起点の領域には宣言の上のコメントも含まれる。
            # replace-procedure と対称にし、get→replace の往復で
            # ヘッダーコメントが消えないようにする。
            code = cm.Lines(proc_start, count)
        except Exception:
            continue

        lines = code.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n').split('\n')
        # 先頭の空行を除去（プロシージャ間の区切り空行は領域に含まれるため）
        while lines and lines[0].strip() == '':
            lines.pop(0)
        # 末尾の空行と、紛れ込んだ次プロシージャの宣言行を除去
        # （「Sub B(): 処理: End Sub」のような1行完結プロシージャは正当な本体なので
        #   対象外。cmd_replace_procedure の混入除去と同じ条件に揃える）
        while len(lines) > 1:
            last = lines[-1].strip()
            if last == '':
                lines.pop()
            elif (re.match(
                    r'^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+',
                    last, re.IGNORECASE)
                  and not re.search(r'\bEnd\s+(?:Sub|Function)\b', last, re.IGNORECASE)):
                lines.pop()
            else:
                break
        return comp.Name, '\n'.join(lines) + '\n'
    return None, None


def cmd_get(args):
    """プロシージャのコードを取得・表示・ファイル保存

    書式:
      get <macro_name>                       全モジュールから検索
      get <module_name> <macro_name>         モジュール指定（スペース区切り）
      get <module_name>.<macro_name>         モジュール指定（ドット区切り）
      get <名1> <名2> <名3> ...              3個以上は複数取得（各要素にドット記法可）
                                             ※出力は連結。書き戻しは従来どおり1本ずつ
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: get [excel_file] <macro_name>  または  get [excel_file] <module_name> <macro_name>")
        return False

    # 取得リクエストの解析
    requests = []                     # [(module_name or None, macro_name), ...]
    if len(rest) >= 3:
        # 複数取得モード（1回のCOM接続でまとめ読み）。各要素は 名前 or モジュール.名前
        for token in rest:
            if '.' in token:
                mn, pn = token.split('.', 1)
                requests.append((mn, pn))
            else:
                requests.append((None, token))
    elif len(rest) == 2 and '.' in rest[0] and '.' in rest[1] and not looks_like_xl_file(rest[1]):
        # 両方ドット記法なら複数取得（get A.x B.y を「モジュールA.x のマクロ B.y」と
        # 誤解釈しないため）
        for token in rest:
            mn, pn = token.split('.', 1)
            requests.append((mn, pn))
    elif len(rest) == 2 and not looks_like_xl_file(rest[1]):
        requests.append((rest[0], rest[1]))          # get <module> <macro>
        print(f"モジュール指定: {rest[0]}")
    elif len(rest) == 1 and '.' in rest[0] and not looks_like_xl_file(rest[0]):
        mn, pn = rest[0].split('.', 1)
        requests.append((mn, pn))                    # get <module>.<macro>
        print(f"モジュール指定: {mn}")
    else:
        requests.append((None, rest[0]))

    xl, wb = get_workbook(target_file)

    results = []
    for module_name, macro_name in requests:
        try:
            comp_name, clean = _extract_proc(wb, module_name, macro_name)
        except ValueError as e:
            print(f"エラー: '{macro_name}' が複数のモジュールに存在します:")
            for mn in e.args[0]:
                print(f"  - {mn}")
            print(f"  モジュールを指定してください。例: py vba_manager.py get {e.args[0][0]} {macro_name}")
            return False
        if comp_name is None:
            print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
            _suggest_similar(macro_name, _all_procedure_names(wb))
            return False
        results.append({'module': comp_name, 'name': macro_name, 'code': clean})

    out_path = getattr(args, 'out_opt', None)
    save_path = os.path.abspath(out_path) if out_path else LAST_PROC_FILE
    joined = '\n'.join(r['code'] for r in results)
    with open(save_path, 'w', encoding='utf-8') as f:
        f.write(joined)

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": True, "file": wb.Name, "saved": save_path,
                          "procs": results}, ensure_ascii=False), file=sys.stdout)
        return True

    for r in results:
        print(f"モジュール  : {r['module']}")
        print(f"プロシージャ: {r['name']}")
        print(f"保存先      : {save_path}")
        print("=" * 60)
        print(r['code'])
        print("=" * 60)
    if len(results) > 1:
        print(f"（{len(results)}本を連結して保存しました。replace-procedure での書き戻しは1本ずつ）")
    return True


def cmd_replace_procedure(args):
    """プロシージャを置換 (コードファイル省略時は _last_proc.vba を使用)"""
    target_file, rest = parse_target_and_rest(args.posargs)

    # コードファイルの決定: --code-file > 位置引数 > _last_proc.vba
    code_file = (getattr(args, 'code_file_opt', None)
                 or (rest[0] if rest else None)
                 or LAST_PROC_FILE)

    resolved = smart_path_resolve(code_file)
    if not resolved or not os.path.exists(resolved):
        print(f"エラー: コードファイルが見つかりません: {code_file}")
        return False

    new_code = read_code_file(resolved)

    # 簡易構文チェック・エンコード検証
    if not validate_vba_code(new_code, getattr(args, 'force', False)):
        return False

    # プロシージャ名を特定
    pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE
    )
    m = pattern.search(new_code)
    if not m:
        print("エラー: コードファイルに Sub/Function 宣言が見つかりません")
        return False
    macro_name = m.group(1)

    # new_code の末尾に次のプロシージャの Sub/Function 宣言が混入していたら除去。
    # ただし「Sub B(): 処理: End Sub」のような1行完結のプロシージャは正当なコード
    # なので対象外（End を伴わない裸の宣言行だけが混入）
    code_lines = new_code.rstrip('\n').split('\n')
    while len(code_lines) > 1:
        last = code_lines[-1].strip()
        if last == '':
            code_lines.pop()
        elif (pattern.match(last)
              and not re.search(r'\bEnd\s+(?:Sub|Function)\b', last, re.IGNORECASE)
              and code_lines[-1].strip() != code_lines[0].strip()):
            code_lines.pop()
        else:
            break
    new_code = '\n'.join(code_lines) + '\n'

    xl, wb = get_workbook(target_file)

    # --module 未指定時：同名プロシージャが複数モジュールにある場合はエラー
    module_opt = getattr(args, 'module_opt', None)
    if not module_opt:
        matched_modules = []
        for comp in wb.VBProject.VBComponents:
            try:
                comp.CodeModule.ProcStartLine(macro_name, 0)
                matched_modules.append(comp.Name)
            except Exception:
                pass
        if len(matched_modules) > 1:
            print(f"エラー: '{macro_name}' が複数のモジュールに存在します:")
            for mn in matched_modules:
                print(f"  - {mn}")
            print(f"  --module オプションで対象を指定してください。")
            print(f"  例: py vba_manager.py replace-procedure --module {matched_modules[0]}")
            return False

    if make_backup(wb.FullName, macro_name) is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        print("  ※ 未保存の新規ブックはバックアップできません。一度保存してから実行してください。")
        return False

    # 対象プロシージャの確認と差分表示
    target_comp = None
    proc_start = 0
    proc_count = 0
    for comp in wb.VBProject.VBComponents:
        if module_opt and comp.Name.lower() != module_opt.lower():
            continue
        cm = comp.CodeModule
        try:
            proc_start = cm.ProcStartLine(macro_name, 0)
            proc_count = cm.ProcCountLines(macro_name, 0)
            target_comp = comp
            break
        except Exception:
            continue

    if not target_comp:
        print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
        _suggest_similar(macro_name, _all_procedure_names(wb))
        print("  新規追加なら add-procedure を使ってください。")
        return False

    # 変更前コードの取得と差分表示
    # ProcStartLine 起点の領域には前後の空行が含まれるため、実置換範囲は
    # 空行を除いて絞る（プロシージャ間の区切り空行を消さないため）
    old_code = target_comp.CodeModule.Lines(proc_start, proc_count)
    old_all = old_code.replace('\r\n', '\n').split('\n')
    lead = 0
    while lead < len(old_all) and old_all[lead].strip() == '':
        lead += 1
    trail = 0
    while trail < len(old_all) - lead and old_all[len(old_all) - 1 - trail].strip() == '':
        trail += 1
    eff_start = proc_start + lead
    eff_count = proc_count - lead - trail

    import difflib
    old_lines = old_all[lead:len(old_all) - trail]
    new_lines = new_code.replace('\r\n', '\n').split('\n')
    if new_lines and new_lines[-1] == '': new_lines.pop()

    diff = list(difflib.unified_diff(
        old_lines, new_lines,
        fromfile=f"Current: {target_comp.Name}.{macro_name}",
        tofile=f"New: {macro_name}",
        lineterm=""
    ))

    if diff:
        print("\n--- 変更差分 (Diff) ---")
        for line in diff:
            print(line)
        print("----------------------\n")
    else:
        # 変更ゼロなら置換しない（Attribute経路だと無変更でも Remove+Import が走り、
        # 無用なリスクを負うだけのため）
        print("変更はありません。置換をスキップしました。")
        return True

    # 確認プロンプト
    if not getattr(args, 'yes', False):
        ans = input(f"プロシージャ '{macro_name}' を置換しますか？ (y/N): ")
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False

    # モジュール単位のバックアップ（Attribute経路の Import 失敗時の復旧素材を兼ねる）
    module_backup = make_module_backup(wb, target_comp.Name)

    print(f"プロシージャ '{macro_name}' を置換中...")

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() != target_comp.Name.lower():
            continue
        cm = comp.CodeModule

        # モジュールをエクスポートして Attribute行の有無を確認
        module_name = comp.Name
        tmp_bas = os.path.join(SCRIPT_DIR, f"_tmp_{module_name}.bas")
        comp.Export(tmp_bas)

        with open(tmp_bas, 'rb') as f:
            bas_content = f.read().decode('cp932')
        bas_lines = bas_content.split('\r\n')

        # .bas 内で対象プロシージャの Sub/Function 宣言行を探す
        proc_pattern = re.compile(
            r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
            r'(?:Sub|Function)\s+' + re.escape(macro_name) + r'\s*[\(\s]',
            re.IGNORECASE
        )
        # 末尾コメント（End Sub 'xxx）を許容しないと次のプロシージャの End Sub まで
        # スキャンが伸び、置換範囲が次のプロシージャを丸ごと巻き込む（消失事故）
        end_pattern = re.compile(
            r"^\s*End\s+(?:Sub|Function)\s*(?:'.*)?$", re.IGNORECASE
        )

        sub_line_idx = None
        proc_end_idx = None
        attr_block = []

        for idx, line in enumerate(bas_lines):
            if sub_line_idx is None and proc_pattern.match(line):
                sub_line_idx = idx
                # Sub宣言の直後の Attribute行を収集
                # （宣言が行継続 " _" で複数行の場合は継続行を読み飛ばしてから収集。
                #   読み飛ばさないと Attribute を見逃し、ショートカット定義が失われる）
                check = idx + 1
                while check < len(bas_lines) and re.search(r'\s_\s*$', bas_lines[check - 1]):
                    check += 1
                while check < len(bas_lines) and bas_lines[check].strip().startswith('Attribute '):
                    attr_block.append(bas_lines[check])
                    check += 1
                # 「Sub X(): 処理: End Sub」の1行完結は宣言行自身で閉じている。
                # 次の End Sub まで探すと後続プロシージャを巻き込んで消すため、
                # ここで終端を確定する（随伴 Attribute 行までが置換対象）
                no_str = re.sub(r'"[^"]*"', '""', line)
                if re.search(r'\bEnd\s+(?:Sub|Function)\b', no_str, re.IGNORECASE):
                    proc_end_idx = check - 1
                    break
            elif sub_line_idx is not None and end_pattern.match(line):
                proc_end_idx = idx
                break

        if sub_line_idx is None or proc_end_idx is None:
            _remove_export_artifacts(tmp_bas)
            continue

        if not attr_block:
            # Attribute行なし → 従来の InsertLines 方式（高速・モジュール順維持）
            # InsertLines は末尾改行を余分な空行として挿入するため取り除く
            _remove_export_artifacts(tmp_bas)
            cm.DeleteLines(eff_start, eff_count)
            cm.InsertLines(eff_start, new_code.rstrip('\n'))
            wb.Save()
            print(f"置換完了: [{comp.Name}] '{macro_name}' → 保存しました")
            return True

        # Attribute行あり → .bas編集 → replace-module 方式
        print(f"  (Attribute行検出 → replace-module方式で処理)")

        # new_code の行を準備（Sub宣言の直後に Attribute行を挿入）
        # この方式は .bas の「宣言行〜End Sub」だけを差し替えるため、
        # 宣言より上のコメントは .bas 側の既存行をそのまま維持し、
        # 新コード側の宣言より上の行は使わない（使うと二重になる）。
        new_lines = new_code.rstrip('\n').split('\n')
        sub_decl_pattern = re.compile(
            r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+',
            re.IGNORECASE
        )
        decl_idx = None
        for ni, nl in enumerate(new_lines):
            if sub_decl_pattern.match(nl):
                decl_idx = ni
                break
        if decl_idx is None:
            _remove_export_artifacts(tmp_bas)
            print("エラー: 置換コードに Sub/Function 宣言が見つかりません")
            return False
        if decl_idx > 0:
            print("  (宣言より上のコメント行は、Attribute方式では .bas 側の既存行を維持します)")
            new_lines = new_lines[decl_idx:]
        for ai, al in enumerate(attr_block):
            new_lines.insert(1 + ai, al)

        # .bas 内の対象プロシージャを置換（Sub宣言行から End Sub まで）
        bas_lines[sub_line_idx:proc_end_idx + 1] = new_lines
        new_bas = '\r\n'.join(bas_lines)

        with open(tmp_bas, 'wb') as f:
            f.write(new_bas.encode('cp932'))

        # Remove + Import（例外時も DisplayAlerts を戻し、一時ファイルを残さない）
        xl.DisplayAlerts = False
        removed = False
        try:
            wb.Save()
            time.sleep(0.5)
            pythoncom.PumpWaitingMessages()
            wb.VBProject.VBComponents.Remove(comp)
            removed = True
            time.sleep(1.5)
            pythoncom.PumpWaitingMessages()
            # 実名検証つき Import（Remove 遅延完了→名前衝突→shu0051 化事故のガード）
            _import_module_verified(wb, tmp_bas, module_name)
            wb.Save()
        except ModuleNameCollisionError as ex:
            # コードは連番付き別名側に生きている。バックアップ再 Import は三重化するので禁止
            _print_collision_guidance(ex, module_name, module_backup)
            return False
        except Exception as ex:
            print(f"エラー: 置換中に失敗しました: {ex}")
            if removed:
                # Remove 成功後の Import 失敗＝開いているブックからモジュール消失。
                # 直前のモジュールバックアップ（置換前の内容）から自動復旧を試みる。
                print(f"⚠ モジュール '{module_name}' は Remove 済みです。バックアップから復旧を試みます...")
                try:
                    if module_backup and os.path.exists(module_backup):
                        _import_module_verified(wb, module_backup, module_name)
                        print(f"復旧成功: {module_backup} を再インポートしました（置換前の内容に戻っています）")
                    else:
                        raise RuntimeError("モジュールバックアップがありません")
                except ModuleNameCollisionError as ex2:
                    print(f"復旧の再 Import で名前衝突: {ex2}")
                    print(f"  置換前の内容は別名モジュール '{ex2.actual_name}' 側にあります。")
                    print(f"  旧 '{module_name}' が消えているのを確認してから '{ex2.actual_name}' を改名してください。")
                except Exception as ex2:
                    print(f"復旧失敗: {ex2}")
                    print("  ⚠ このままブックを保存するとモジュールがファイルからも消えます。")
                    print("  対処: ブックを『保存せずに閉じて』開き直せば置換前の状態に戻ります。")
                    if module_backup:
                        print(f"  または backups のバックアップを手動で Import: {module_backup}")
            return False
        finally:
            xl.DisplayAlerts = True
            if os.path.exists(tmp_bas):
                _remove_export_artifacts(tmp_bas)
        # ここまで来たら Import 済み・実名検証済み（黙って成功と報告しない、の実装）
        print(f"置換完了: [{module_name}] '{macro_name}' → 保存しました (Attribute保持)")
        return True

    print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
    return False


def cmd_add_procedure(args):
    """新規プロシージャをモジュール末尾に追加: add-procedure [excel_file] <モジュール名>

    コードは _last_proc.vba（または --code-file）から。replace-procedure が
    既存置換専用なのに対し、こちらは「新しい Sub を1本足す」軽量経路
    （InsertLines のみ・Remove+Import 不要）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: add-procedure [excel_file] <モジュール名> [--code-file f] [-y]")
        print("  追加するコードは _last_proc.vba（または --code-file）に置く")
        return False
    module_name = rest[0]
    if _reject_extra_args(rest, 1, '使い方: add-procedure [excel_file] <モジュール名>'):
        return False

    code_file = getattr(args, 'code_file_opt', None) or LAST_PROC_FILE
    resolved = smart_path_resolve(code_file)
    if not resolved or not os.path.exists(resolved):
        print(f"エラー: コードファイルが見つかりません: {code_file}")
        return False
    new_code = read_code_file(resolved)
    if not validate_vba_code(new_code, force=getattr(args, 'force', False)):
        return False
    m = re.search(r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
                  r'(?:Sub|Function)\s+([^\s\(]+)', new_code,
                  re.MULTILINE | re.IGNORECASE)
    if not m:
        print("エラー: コードに Sub/Function 宣言が見つかりません")
        return False
    proc_name = m.group(1)

    xl, wb = get_workbook(target_file)
    comp = None
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == module_name.lower():
            comp = c
            break
    if comp is None:
        print(f"エラー: モジュール '{module_name}' が見つかりません（list-modules で確認）")
        return False
    cm = comp.CodeModule
    # 同名の重複挿入を防止（ブック全体はマクロ実行時の曖昧さになるだけだが、
    # 同一モジュール内はコンパイルエラーになるため必ず止める）
    try:
        cm.ProcStartLine(proc_name, 0)
        exists = True
    except Exception:
        exists = False
    if exists:
        print(f"エラー: '{proc_name}' は [{comp.Name}] に既に存在します。修正なら replace-procedure を使ってください。")
        return False

    print(f"--- 追加するプロシージャ: [{comp.Name}] {proc_name} ---")
    print(new_code.rstrip('\n'))
    print("-" * 40)
    if not getattr(args, 'yes', False):
        ans = input(f"モジュール '{comp.Name}' の末尾に追加しますか？ (y/N): ")
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    if make_backup(wb.FullName, f"add_{proc_name}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False

    body = new_code.rstrip('\n')
    if cm.CountOfLines > 0:
        body = '\n' + body     # 既存コードとの区切りの空行
    cm.InsertLines(cm.CountOfLines + 1, body)
    wb.Save()
    print(f"追加完了: [{comp.Name}] '{proc_name}' → 保存しました")
    return True


def cmd_add_module(args):
    """新規モジュールを追加: add-module [excel_file] <モジュール名> [--type std|class|form]

    add-procedure/replace-module は既存モジュール前提のため、まっさらなブックには
    器を作れなかった。これはその穴を埋める純機械コマンド（VBComponents.Add）。
    std=標準モジュール / class=クラスモジュール / form=UserForm。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: add-module [excel_file] <モジュール名> [--type std|class|form]")
        return False
    module_name = rest[0]
    if _reject_extra_args(rest, 1, '使い方: add-module [excel_file] <モジュール名> [--type std|class|form]'):
        return False

    type_opt = (getattr(args, 'type_opt', None) or 'std').lower()
    type_map = {'std': 1, 'standard': 1, 'class': 2, 'cls': 2, 'form': 3, 'userform': 3}
    if type_opt not in type_map:
        print(f"エラー: 不明な --type '{type_opt}'（std|class|form のいずれか）")
        return False
    comp_type = type_map[type_opt]

    reason = check_vba_identifier(module_name)
    if reason:
        print(f"エラー: モジュール名が VBA の識別子規則に反しています: {reason}")
        return False

    xl, wb = get_workbook(target_file)
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == module_name.lower():
            print(f"エラー: モジュール '{c.Name}' は既に存在します。（別名にするか delete-module で消してから）")
            return False

    if make_backup(wb.FullName, f"add_module_{module_name}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False

    try:
        comp = wb.VBProject.VBComponents.Add(comp_type)
    except Exception as e:
        print(f"エラー: モジュールを追加できませんでした: {e}")
        print("  VBProject へのアクセスが信頼されているか確認してください（setup-check）。")
        return False
    try:
        comp.Name = module_name
    except Exception as e:
        print(f"エラー: モジュール名 '{module_name}' を設定できませんでした: {e}")
        # 追加済みの既定名（Module1等）のまま残すと、後で保存されたときに
        # ゴミモジュールがブックに焼き付くため撤去する
        try:
            wb.VBProject.VBComponents.Remove(comp)
        except Exception:
            print("警告: 追加途中のモジュールを撤去できませんでした"
                  "（既定名のモジュールが残っていたら手で削除してください）")
        return False
    wb.Save()
    label = {1: '標準モジュール', 2: 'クラスモジュール', 3: 'ユーザーフォーム'}[comp_type]
    print(f"追加完了: {label} '{comp.Name}' → 保存しました")
    return True


def cmd_delete_procedure(args):
    """プロシージャを削除: delete-procedure [excel_file] <Sub名>

    同名が複数モジュールにある場合は --module で明示（get/replace と同じ流儀）。
    削除対象のコードを表示してから確認（-y でスキップ）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: delete-procedure [excel_file] <Sub名> [--module 名] [-y]")
        return False
    macro_name = rest[0]
    if _reject_extra_args(rest, 1, '使い方: delete-procedure [excel_file] <Sub名> [--module 名]'):
        return False
    module_opt = getattr(args, 'module_opt', None)

    xl, wb = get_workbook(target_file)

    # 対象特定（同名複数はエラーで候補列挙＝対象取り違え防止）
    matches = []
    for comp in wb.VBProject.VBComponents:
        if module_opt and comp.Name.lower() != module_opt.lower():
            continue
        try:
            start = comp.CodeModule.ProcStartLine(macro_name, 0)
            count = comp.CodeModule.ProcCountLines(macro_name, 0)
            matches.append((comp, start, count))
        except Exception:
            continue
    if not matches:
        print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
        _suggest_similar(macro_name, _all_procedure_names(wb))
        return False
    if len(matches) > 1:
        print(f"エラー: '{macro_name}' は複数のモジュールにあります。--module で指定してください:")
        for comp, _, _ in matches:
            print(f"  {comp.Name}")
        return False

    comp, start, count = matches[0]
    cm = comp.CodeModule
    print(f"--- 削除するプロシージャ: [{comp.Name}] {macro_name} ({count}行) ---")
    print(cm.Lines(start, count).rstrip())
    print("-" * 40)
    if not getattr(args, 'yes', False):
        ans = input(f"[{comp.Name}] から '{macro_name}' を削除しますか？ (y/N): ")
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    if make_backup(wb.FullName, f"delete_{macro_name}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False
    make_module_backup(wb, comp.Name)

    cm.DeleteLines(start, count)
    wb.Save()
    print(f"削除完了: [{comp.Name}] '{macro_name}' → 保存しました")
    return True


def cmd_delete_module(args):
    """モジュール丸ごと削除: delete-module [excel_file] <モジュール名> [-y]

    削除前に中身の要約を表示して確認。ブック＋モジュールの自動バックアップつき
    （戻すのは restore）。ThisWorkbook / シートモジュールは削除できない（VBAの制約）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: delete-module [excel_file] <モジュール名> [-y]")
        return False
    module_name = rest[0]
    if _reject_extra_args(rest, 1, '使い方: delete-module [excel_file] <モジュール名>'):
        return False

    xl, wb = get_workbook(target_file)
    comp = None
    for c in wb.VBProject.VBComponents:
        if c.Name.lower() == module_name.lower():
            comp = c
            break
    if comp is None:
        print(f"エラー: モジュール '{module_name}' が見つかりません")
        print("  存在するモジュール: " + ', '.join(c.Name for c in wb.VBProject.VBComponents))
        return False
    if int(comp.Type) == 100:
        print(f"エラー: '{comp.Name}' はブック/シートのモジュールなので削除できません（VBAの制約）")
        return False

    cm = comp.CodeModule
    n = cm.CountOfLines
    kind = {1: '標準モジュール', 2: 'クラス', 3: 'フォーム'}.get(int(comp.Type), '?')
    print(f"削除対象: [{comp.Name}]（{kind}・{n}行）")
    if n > 0:
        head = cm.Lines(1, min(n, 8)).replace('\r\n', '\n')
        print("  --- 先頭8行 ---")
        for ln in head.split('\n'):
            print(f"  {ln}")
        if n > 8:
            print(f"  … 他 {n - 8}行")

    if not getattr(args, 'yes', False):
        ans = input(f"モジュール '{comp.Name}' を丸ごと削除しますか？ (y/N): ")
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False
    if make_backup(wb.FullName, f"delmod_{comp.Name}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False
    backup = make_module_backup(wb, comp.Name)

    wb.VBProject.VBComponents.Remove(comp)
    wb.Save()
    print(f"削除完了: モジュール '{module_name}' → 保存しました")
    if backup:
        print(f"  戻すには: py vba_manager.py restore {os.path.basename(backup)}")
    return True


def cmd_replace_module(args):
    """モジュール全体を Remove+Import で置換 (Attribute を正しく処理)"""
    target_file, rest = parse_target_and_rest(args.posargs)

    if len(rest) < 2:
        print("使い方: replace-module [excel_file] <module_name> <bas_file>")
        return False
    module_name, code_file = rest[0], rest[1]

    resolved = smart_path_resolve(code_file)
    if not resolved or not os.path.exists(resolved):
        print(f"エラー: コードファイルが見つかりません: {code_file}")
        return False

    # UTF-8化事故（CP932で書くべき .bas をUTF-8で保存）の水際チェック
    if not validate_bas_encoding(resolved):
        return False

    # フォーム（.frm）はレイアウトを .frx に持ち、Import は .frm と同名の .frx を
    # 同じ場所に要求する。.frx 無しで Import するとレイアウトが空のフォームになるため止める。
    is_form_file = resolved.lower().endswith('.frm')
    src_frx = os.path.splitext(resolved)[0] + '.frx'
    if is_form_file and not os.path.exists(src_frx):
        print(f"エラー: フォームの相方 {os.path.basename(src_frx)} が見つかりません。")
        print("  .frm と .frx は同じフォルダにペアで置いてください（.frx が無いとレイアウトが失われます）。")
        return False

    # 改行二重化（\r\r\n 化）の水際チェック＆修正。
    # 過去の二重化事故は、外部で作られた .bas が既に倍増した状態で replace-module に
    # 渡され、それを無検査で Import したのが原因だった。ここで修正してから取り込む。
    import_path = resolved
    tmp_norm = None
    fixed_bytes, raw_bytes, was_fixed = normalize_bas_newlines(resolved)
    if was_fixed:
        # VBA は \r\r\n を「行＋空行」と解釈して行数が倍に見える。実症状に合わせて
        # \r\n / \r / \n のいずれでも行が区切られる前提で数える。
        before_lines = len(re.split(r'\r\n|\r|\n', raw_bytes.decode('cp932')))
        after_lines = len(re.split(r'\r\n|\r|\n', fixed_bytes.decode('cp932')))
        print(f"⚠ 改行の二重化を検知しました。インポート前に修正します: {before_lines}行 → {after_lines}行")
        tmp_norm = os.path.join(SCRIPT_DIR, f"_norm_{os.path.basename(resolved)}")
        with open(tmp_norm, 'wb') as f:
            f.write(fixed_bytes)
        if is_form_file:
            # 正規化後の .frm から Import する場合も .frx を随伴させる
            # （コピーしないとレイアウト無しで取り込まれる穴だった）
            shutil.copy2(src_frx, os.path.splitext(tmp_norm)[0] + '.frx')
        import_path = tmp_norm

    # .bas の VB_Name と指定モジュール名の照合（別モジュール取り違えの防止）。
    # VB_Name が無い .bas は Import 時に Module1 等の別名で入り「Xを消してYが増える」
    # 事故になるため、ここで止める。
    with open(import_path, 'rb') as f:
        bas_head = f.read().decode('cp932', errors='replace')
    m_name = re.search(r'^Attribute\s+VB_Name\s*=\s*"([^"]*)"', bas_head,
                       re.MULTILINE | re.IGNORECASE)
    if not m_name:
        print(f"エラー: {os.path.basename(resolved)} に Attribute VB_Name 行がありません。")
        print("  このまま Import すると別名モジュールとして取り込まれます。")
        print("  export-module で出力した .bas をベースに編集してください。")
        if tmp_norm and os.path.exists(tmp_norm):
            _remove_export_artifacts(tmp_norm)
        return False
    if m_name.group(1).lower() != module_name.lower():
        print(f"エラー: .bas の VB_Name '{m_name.group(1)}' が指定モジュール名 '{module_name}' と一致しません。")
        print("  別モジュールの .bas を取り込もうとしている可能性があります（対象取り違え防止のため停止）。")
        if tmp_norm and os.path.exists(tmp_norm):
            _remove_export_artifacts(tmp_norm)
        return False

    # プロシージャ名の識別子チェック（先頭 _ 等は Import 自体は通るがコンパイルで死ぬ）
    bad_names = _find_invalid_procedure_names(re.sub(r'\r\n|\r', '\n', bas_head))
    if bad_names and not getattr(args, 'force', False):
        print("エラー: VBA の識別子規則に反するプロシージャ名があります（--force で強行可）:")
        for ln, _nm, reason in bad_names:
            print(f"  行{ln}: {reason}")
        if tmp_norm and os.path.exists(tmp_norm):
            _remove_export_artifacts(tmp_norm)
        return False

    xl, wb = get_workbook(target_file)
    if make_backup(wb.FullName, f"module_{module_name}") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        print("  ※ 未保存の新規ブックはバックアップできません。一度保存してから実行してください。")
        return False
    # モジュール単位のバックアップ（Import 失敗時の復旧素材を兼ねる）
    module_backup = make_module_backup(wb, module_name)
    print(f"モジュール '{module_name}' を Remove+Import で置換中...")

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() == module_name.lower():
            xl.DisplayAlerts = False
            removed = False
            try:
                wb.Save()
                time.sleep(0.5)
                pythoncom.PumpWaitingMessages()
                wb.VBProject.VBComponents.Remove(comp)
                removed = True
                time.sleep(1.5)
                pythoncom.PumpWaitingMessages()
                # 実名検証つき Import（Remove 遅延完了→名前衝突→shu0051 化事故のガード）
                _import_module_verified(wb, import_path, module_name)
                wb.Save()
            except ModuleNameCollisionError as ex:
                # コードは連番付き別名側に生きている。バックアップ再 Import は三重化するので禁止
                _print_collision_guidance(ex, module_name, module_backup)
                return False
            except Exception as ex:
                print(f"エラー: 置換中に失敗しました: {ex}")
                if removed:
                    # Remove だけ成功して Import に失敗＝開いているブックからモジュール消失。
                    # 直前のモジュールバックアップから自動復旧を試みる。
                    print(f"⚠ モジュール '{module_name}' は Remove 済みです。バックアップから復旧を試みます...")
                    try:
                        if module_backup and os.path.exists(module_backup):
                            _import_module_verified(wb, module_backup, module_name)
                            print(f"復旧成功: {module_backup} を再インポートしました（置換前の内容に戻っています）")
                        else:
                            raise RuntimeError("モジュールバックアップがありません")
                    except ModuleNameCollisionError as ex2:
                        print(f"復旧の再 Import で名前衝突: {ex2}")
                        print(f"  置換前の内容は別名モジュール '{ex2.actual_name}' 側にあります。")
                        print(f"  旧 '{module_name}' が消えているのを確認してから '{ex2.actual_name}' を改名してください。")
                    except Exception as ex2:
                        print(f"復旧失敗: {ex2}")
                        print("  ⚠ このままブックを保存するとモジュールがファイルからも消えます。")
                        print("  対処: ブックを『保存せずに閉じて』開き直せば置換前の状態に戻ります。")
                        if module_backup:
                            print(f"  または backups のバックアップを手動で Import: {module_backup}")
                return False
            finally:
                xl.DisplayAlerts = True
                if tmp_norm and os.path.exists(tmp_norm):
                    _remove_export_artifacts(tmp_norm)
            print(f"置換完了: モジュール '{module_name}' → 保存しました")
            return True

    if tmp_norm and os.path.exists(tmp_norm):
        _remove_export_artifacts(tmp_norm)
    print(f"エラー: モジュール '{module_name}' が見つかりません")
    return False


def _parse_module_blocks(bas_text):
    """
    .bas モジュールの本文を ヘッダー / Sub・Functionブロック群 / 末尾 に分解する。
    各ブロックは前ブロックの End Sub/Function の次行から自分の End Sub/Function 行までを所有。
    Attribute 行（ショートカット定義）は Sub 内に含まれるので自動的にブロック内に入る。
    戻り値: (header_lines, blocks, trailing_lines)
        block: {'name': str, 'kind': 'Sub'|'Function', 'lines': [str,...]}
    """
    sub_pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?P<kind>Sub|Function|Property\s+(?:Get|Let|Set))\s+(?P<name>[^\s\(]+)',
        re.IGNORECASE
    )
    # 末尾コメント（End Sub 'xxx）を許容（cmd_replace_procedure の end_pattern と同じ理由）
    end_pattern = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\s*(?:'.*)?$", re.IGNORECASE)

    lines = bas_text.split('\r\n')
    n = len(lines)

    first_sub = None
    for i, line in enumerate(lines):
        if sub_pattern.match(line):
            first_sub = i
            break

    if first_sub is None:
        return lines, [], []

    header = lines[:first_sub]

    blocks = []
    cur_start = first_sub
    i = first_sub
    while i < n:
        m = sub_pattern.match(lines[i])
        if m:
            # 「Sub X(): 処理: End Sub」の1行完結は宣言行自身で閉じている。
            # 次の End 行まで探すと後続プロシージャを同一ブロックに巻き込むため、
            # その行（＋随伴 Attribute 行）でブロックを閉じる
            no_str = re.sub(r'"[^"]*"', '""', lines[i])
            if re.search(r'\bEnd\s+(?:Sub|Function|Property)\b', no_str, re.IGNORECASE):
                j = i
                while j + 1 < n and lines[j + 1].strip().startswith('Attribute '):
                    j += 1
            else:
                j = i + 1
                while j < n and not end_pattern.match(lines[j]):
                    j += 1
                if j >= n:
                    break
            blocks.append({
                'name':  m.group('name'),
                'kind':  ' '.join(m.group('kind').split()).lower(),
                'lines': lines[cur_start:j + 1],
            })
            cur_start = j + 1
            i = j + 1
        else:
            i += 1

    trailing = lines[cur_start:]
    return header, blocks, trailing


def _write_module(header, blocks, trailing):
    parts = list(header)
    for b in blocks:
        parts.extend(b['lines'])
    parts.extend(trailing)
    return '\r\n'.join(parts)


def cmd_reorder_macro(args):
    """
    マクロを同モジュール内で 1 つ上 / 下のマクロと入れ替える。

    終了コード:
        0 : 成功
        1 : 引数エラー / その他
        2 : マクロが見つからない
        3 : 既にモジュール内の最初/最後（境界）
    """
    rest = list(args.posargs)
    if len(rest) < 2:
        print("使い方: reorder-macro <macro_name> <up|down|top|bottom|位置番号>")
        sys.exit(1)
    macro_name = rest[0]
    direction  = rest[1].lower()
    if direction not in ('up', 'down', 'top', 'bottom') and not direction.isdigit():
        print("方向は up|down|top|bottom または移動先の位置番号(1始まり)を指定してください")
        sys.exit(1)

    xl, wb = get_workbook(None)  # ActiveWorkbook 自動検出

    # 対象マクロを含む標準モジュールを特定
    # （replace-procedure / delete-procedure と同様、同名複数は取り違え防止で停止）
    matched_comps = []
    for comp in wb.VBProject.VBComponents:
        if comp.Type != 1:  # 標準モジュールのみ対象
            continue
        try:
            comp.CodeModule.ProcStartLine(macro_name, 0)
            matched_comps.append(comp)
        except Exception:
            continue

    if not matched_comps:
        print(f"エラー: マクロ '{macro_name}' が標準モジュールに見つかりません")
        sys.exit(2)
    if len(matched_comps) > 1:
        names = ', '.join(c.Name for c in matched_comps)
        print(f"エラー: 同名マクロ '{macro_name}' が複数のモジュールにあります: {names}")
        print("  対象を特定できないため中止しました（重量操作の取り違え防止）。")
        sys.exit(2)
    target_comp = matched_comps[0]

    module_name = target_comp.Name

    # モジュールをエクスポートして CP932 のまま読み込む
    tmp_bas = os.path.join(SCRIPT_DIR, f"_tmp_reorder_{module_name}.bas")
    target_comp.Export(tmp_bas)
    with open(tmp_bas, 'rb') as f:
        bas_text = f.read().decode('cp932')

    header, blocks, trailing = _parse_module_blocks(bas_text)

    # 対象ブロックの index（VBA の名前は大小無視なので比較も合わせる）
    target_idx = None
    for i, b in enumerate(blocks):
        if b['name'].lower() == macro_name.lower():
            target_idx = i
            break
    if target_idx is None:
        _remove_export_artifacts(tmp_bas)
        print(f"エラー: モジュール {module_name} に Sub '{macro_name}' が見つかりません")
        sys.exit(2)

    # 並べ替えの位置は Sub 単位で数える（Function は位置の数に入れない）
    visible_indices = [
        i for i, b in enumerate(blocks) if b['kind'] == 'sub'
    ]

    if target_idx not in visible_indices:
        _remove_export_artifacts(tmp_bas)
        print(f"エラー: '{macro_name}' は Sub ではないため並べ替えの対象外です")
        sys.exit(2)

    vis_pos = visible_indices.index(target_idx)

    if direction == 'up':
        if vis_pos == 0:
            _remove_export_artifacts(tmp_bas)
            print(f"BOUNDARY: [{module_name}] 内で既に最初です")
            sys.exit(3)
        swap_block_idx = visible_indices[vis_pos - 1]
        # ブロック単位で入れ替え（Attribute 行は各ブロック内に含まれているので一緒に動く）
        blocks[target_idx], blocks[swap_block_idx] = blocks[swap_block_idx], blocks[target_idx]
    elif direction == 'down':
        if vis_pos == len(visible_indices) - 1:
            _remove_export_artifacts(tmp_bas)
            print(f"BOUNDARY: [{module_name}] 内で既に最後です")
            sys.exit(3)
        swap_block_idx = visible_indices[vis_pos + 1]
        blocks[target_idx], blocks[swap_block_idx] = blocks[swap_block_idx], blocks[target_idx]
    else:
        # top / bottom / 位置番号: 一発で目的位置へ（up/down を N回＝重量処理N回、の代わり）
        if direction == 'top':
            new_pos = 0
        elif direction == 'bottom':
            new_pos = len(visible_indices) - 1
        else:
            new_pos = max(0, min(int(direction) - 1, len(visible_indices) - 1))
        if new_pos == vis_pos:
            _remove_export_artifacts(tmp_bas)
            print(f"BOUNDARY: [{module_name}] 既にその位置です")
            sys.exit(3)
        # 可視ブロックの並びだけを組み替え、非表示ブロックの位置は維持する
        vis_blocks = [blocks[i] for i in visible_indices]
        blk = vis_blocks.pop(vis_pos)
        vis_blocks.insert(new_pos, blk)
        for i, b in zip(visible_indices, vis_blocks):
            blocks[i] = b

    new_text = _write_module(header, blocks, trailing)
    with open(tmp_bas, 'wb') as f:
        f.write(new_text.encode('cp932'))

    # 破壊操作（Remove+Import）なので他コマンドと同じくバックアップ必須。
    # 取れなければ停止（--force で強行）
    if make_backup(wb.FullName, f"reorder_{macro_name}") is None and not getattr(args, 'force', False):
        _remove_export_artifacts(tmp_bas)
        print("エラー: バックアップが取れなかったため中止しました（--force で強行可能）")
        return False
    # モジュール単位のバックアップ（Import 失敗時の自動復旧素材。replace-module と同型）
    module_backup = make_module_backup(wb, module_name)
    print(f"並べ替え中: [{module_name}] '{macro_name}' を {direction}")

    # replace-module と同じ安定化手順（sleep + PumpWaitingMessages）に揃える
    xl.DisplayAlerts = False
    removed = False
    success = False
    try:
        wb.Save()
        time.sleep(0.5)
        pythoncom.PumpWaitingMessages()
        wb.VBProject.VBComponents.Remove(target_comp)
        removed = True
        time.sleep(1.5)
        pythoncom.PumpWaitingMessages()
        # 実名検証つき Import（Remove 遅延完了→名前衝突→shu0051 化事故のガード）
        _import_module_verified(wb, tmp_bas, module_name)
        wb.Save()
        success = True
    except ModuleNameCollisionError as ex:
        # コードは連番付き別名側に生きている。バックアップ再 Import は三重化するので禁止
        _print_collision_guidance(ex, module_name, module_backup, err=True)
        return False
    except Exception as ex:
        # Remove 成功後に Import が失敗するとモジュールが消えたままになる。
        # replace-module と同じく、モジュールバックアップからの自動復旧を先に試みる
        if removed:
            print(f"エラー: 並べ替え中に失敗しました（モジュール '{module_name}' が"
                  f"開いているブックから外れた可能性があります）: {ex}", file=sys.stderr)
            try:
                if module_backup and os.path.exists(module_backup):
                    _import_module_verified(wb, module_backup, module_name)
                    print(f"復旧成功: {module_backup} を再インポートしました"
                          f"（並べ替え前の内容に戻っています）", file=sys.stderr)
                else:
                    raise RuntimeError("モジュールバックアップがありません")
            except ModuleNameCollisionError as ex2:
                print(f"復旧の再 Import で名前衝突: {ex2}", file=sys.stderr)
                print(f"  並べ替え前の内容は別名モジュール '{ex2.actual_name}' 側にあります。", file=sys.stderr)
                print(f"  旧 '{module_name}' が消えているのを確認してから"
                      f" '{ex2.actual_name}' を改名してください。", file=sys.stderr)
            except Exception as ex2:
                print(f"復旧失敗: {ex2}", file=sys.stderr)
                print("  ⚠ このままブックを保存するとモジュールがファイルからも消えます。", file=sys.stderr)
                print("  対処: ブックを『保存せずに閉じて』開き直せば並べ替え前の状態に戻ります。", file=sys.stderr)
                if module_backup:
                    print(f"  または並べ替え前のモジュールを restore: {module_backup}", file=sys.stderr)
                print(f"  並べ替え後のコードも残してあります（restore に渡せます）: {tmp_bas}", file=sys.stderr)
        else:
            print(f"エラー: 並べ替えに失敗しました: {ex}", file=sys.stderr)
        return False
    finally:
        xl.DisplayAlerts = True
        # 「Remove 済みで Import 失敗」のときだけ復旧用に残し、それ以外は掃除する
        if not (removed and not success) and os.path.exists(tmp_bas):
            _remove_export_artifacts(tmp_bas)

    print(f"完了: [{module_name}] '{macro_name}' を {direction}に移動")
    sys.exit(0)


def cmd_export_module(args):
    """モジュールを .bas ファイルにエクスポート"""
    target_file, rest = parse_target_and_rest(args.posargs)

    if not rest:
        print("使い方: export-module [excel_file] <module_name>")
        return False
    module_name = rest[0]

    xl, wb = get_workbook(target_file)

    # export-all と同じ Type 別拡張子を使う。フォームを .bas で書き出すと、
    # replace-module の「.frm なのに .frx が無い」ガードをすり抜けてしまう
    ext_map = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() == module_name.lower():
            # 表記ゆれ（大小文字）でファイル名が実モジュール名とズレないよう comp.Name を使う
            ext = ext_map.get(int(comp.Type), '.bas')
            out_path = os.path.join(SCRIPT_DIR, f"{comp.Name}{ext}")
            if os.path.exists(out_path):
                print(f"（既存の {os.path.basename(out_path)} を上書きします）")
            comp.Export(out_path)
            print(f"エクスポート完了: {out_path}")
            return True

    print(f"エラー: モジュール '{module_name}' が見つかりません")
    print("  存在するモジュール: " + ', '.join(c.Name for c in wb.VBProject.VBComponents))
    return False


def cmd_export_all(args):
    """全モジュールを一括エクスポート: export-all [excel_file] [--dir 出力先] [--check]

    1回のCOM接続で全 VBComponents を書き出す（1コマンドずつ回すと数分かかる
    ことが実測済みの作業を1コマンド化）。--check で書き出した .bas/.frm に
    check-bas 相当の機械検査（文字コード/改行二重化/重複）をその場でかける。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    out_dir = getattr(args, 'dir_opt', None) or SCRIPT_DIR
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    ext_map = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}

    xl, wb = get_workbook(target_file)
    exported = []
    skipped = 0
    for comp in wb.VBProject.VBComponents:
        ctype = int(comp.Type)
        if ctype not in ext_map:
            skipped += 1
            continue
        if ctype == 100 and comp.CodeModule.CountOfLines == 0:
            skipped += 1               # 空の ThisWorkbook / Sheet モジュールは省く
            continue
        out_path = os.path.join(out_dir, comp.Name + ext_map[ctype])
        comp.Export(out_path)
        exported.append(out_path)
        print(f"  {os.path.basename(out_path)}")
    print(f"エクスポート完了: {len(exported)}本 → {out_dir}"
          + (f"（空モジュール等 {skipped}本はスキップ）" if skipped else ""))

    if getattr(args, 'check', False):
        print("----- 取り込み前検査 (check-bas 相当) -----")
        ng = 0
        for p in exported:
            if not p.lower().endswith(('.bas', '.frm', '.cls')):
                continue
            ok = validate_bas_encoding(p)
            _, _, doubled = normalize_bas_newlines(p)
            with open(p, 'rb') as f:
                norm_text = re.sub(r'\r\n|\r', '\n', f.read().decode('cp932', errors='replace'))
            dups = _find_duplicate_procedures(norm_text)
            if ok and not doubled and not dups:
                print(f"  [OK] {os.path.basename(p)}")
            else:
                ng += 1
                marks = []
                if not ok:
                    marks.append("文字コード")
                if doubled:
                    marks.append("改行二重化")
                if dups:
                    marks.append(f"重複({', '.join(dups)})")
                print(f"  [NG] {os.path.basename(p)}: {' / '.join(marks)}")
        print(f"----- 検査結果: NG {ng} / {len(exported)}本 -----")
        return ng == 0
    return True


def cmd_list_backups(args):
    """バックアップの一覧: list-backups [キーワード]（COM不要）

    backups フォルダの内容を新しい順に表示。restore の対象選びに使う。
    """
    kw = args.posargs[0] if args.posargs else None
    if not os.path.isdir(BACKUP_DIR):
        print(f"バックアップフォルダがありません: {BACKUP_DIR}")
        return False
    entries = []
    for name in os.listdir(BACKUP_DIR):
        path = os.path.join(BACKUP_DIR, name)
        if not os.path.isfile(path):
            continue
        if kw and kw.lower() not in name.lower():
            continue
        entries.append((os.path.getmtime(path), name, os.path.getsize(path)))
    entries.sort(reverse=True)
    if not entries:
        print("該当するバックアップはありません。" + (f"（キーワード: {kw}）" if kw else ""))
        return True
    limit = getattr(args, 'max_hits', None) or 30
    print(f"--- バックアップ一覧（新しい順・{min(limit, len(entries))}/{len(entries)}件） ---")
    for mtime, name, size in entries[:limit]:
        stamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
        low = name.lower()
        if low.endswith(('.bas', '.frm', '.cls')):
            kind = "モジュール"
        elif low.endswith(('.xlsm', '.xlsx', '.xlsb', '.xls', '.xlam')):
            kind = "ブック"
        else:
            kind = "その他"
        print(f"  {stamp}  [{kind}] {name}  ({size:,} bytes)")
    if len(entries) > limit:
        print(f"  …他 {len(entries) - limit}件（--max で上限変更可）")
    print(f"場所: {BACKUP_DIR}")
    print("戻すには: py vba_manager.py restore <ファイル名>   （モジュール .bas/.frm のみ）")
    return True


def cmd_restore(args):
    """モジュールバックアップを開いているブックに書き戻す: restore <バックアップ.bas>

    対象モジュール名はファイル内の Attribute VB_Name から機械的に取得し、
    replace-module と同じ経路（照合・ガード・自動復旧つき）で適用する。
    どの世代に戻すかの判断はユーザー/AI側（list-backups で選ぶ）。
    ブック丸ごと（.xlsm）のバックアップはこのコマンドでは扱わない
    （開いているブックへの上書きになるため。必要ならExcelを閉じて手動コピー）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: restore [excel_file] <バックアップファイル名>   （list-backups で確認）")
        return False
    name = rest[0]
    if name.lower().endswith(('.xlsm', '.xlsx', '.xlsb', '.xls')):
        print("エラー: ブック丸ごとのバックアップは restore では扱いません。")
        print("  （開いているブックそのものの上書きになるため。Excelを閉じて手動でコピーしてください）")
        return False
    path = name if os.path.isabs(name) else os.path.join(BACKUP_DIR, name)
    if not os.path.exists(path):
        print(f"エラー: バックアップが見つかりません: {path}")
        print("  list-backups で名前を確認してください。")
        return False
    with open(path, 'rb') as f:
        head = f.read().decode('cp932', errors='replace')
    m = re.search(r'^Attribute\s+VB_Name\s*=\s*"([^"]*)"', head, re.MULTILINE | re.IGNORECASE)
    if not m:
        print(f"エラー: {os.path.basename(path)} に VB_Name がありません（モジュールバックアップではない可能性）")
        return False
    module_name = m.group(1)
    print(f"復元: モジュール '{module_name}' ← backups/{os.path.basename(path)}")
    import argparse as _ap
    ns = _ap.Namespace(posargs=([target_file] if target_file else []) + [module_name, path],
                       force=getattr(args, 'force', False))
    return cmd_replace_module(ns)


def cmd_list_shortcuts(args):
    """ショートカットキーが設定されているマクロの一覧表示"""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)

    shortcuts = []

    for comp in wb.VBProject.VBComponents:
        if comp.Type not in (1, 2, 3, 100):
            continue

        tmp_file = os.path.join(SCRIPT_DIR, f"_tmp_sc_{comp.Name}.bas")
        try:
            comp.Export(tmp_file)
            with open(tmp_file, 'rb') as f:
                content = f.read().decode('cp932', errors='replace')
        except Exception:
            continue
        finally:
            _remove_export_artifacts(tmp_file)

        # Attribute マクロ名.VB_ProcData.VB_Invoke_Func = "キー\n14"
        pattern = re.compile(
            r'Attribute\s+([^.\s]+)\.VB_ProcData\.VB_Invoke_Func\s*=\s*"([^"]+)"',
            re.IGNORECASE | re.DOTALL
        )
        for m in pattern.finditer(content):
            macro_name = m.group(1)
            raw_val = m.group(2)
            # 文字列としての "\n" や "\r" を本物の改行コードに変換してから分割
            raw_val_clean = raw_val.replace('\\n', '\n').replace('\\r', '\r')
            key_char = raw_val_clean.split('\n')[0].split('\r')[0]
            if not key_char:
                continue

            if len(key_char) == 1:
                if key_char.isupper():
                    shortcut_str = f"Ctrl + Shift + {key_char}"
                else:
                    shortcut_str = f"Ctrl + {key_char}"
            else:
                shortcut_str = f"Ctrl + {key_char}"

            shortcuts.append({
                'module': comp.Name,
                'macro': macro_name,
                'shortcut': shortcut_str
            })

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": True, "file": wb.Name, "shortcuts": shortcuts}, ensure_ascii=False), file=sys.stdout)
        return True

    if not shortcuts:
        print("ショートカットキーが設定されているマクロはありません。")
        return True

    print(f"設定されているショートカットキー一覧 (数: {len(shortcuts)})")
    print("-" * 60)
    for item in shortcuts:
        print(f"[{item['module']}] {item['macro']} -> {item['shortcut']}")
    print("-" * 60)
    return True


def cmd_setup_check(args):
    """導入セルフ診断: setup-check

    「会話するだけでマクロが直る」環境に必要なものが揃っているかを○×で表示する。
    初心者が最初に打つ1コマンド。Excel を起動していなくても動く
    （VBOM 信頼設定のチェックだけは Excel 起動中に実施）。
    """
    results = []          # (ok: bool|None, 項目, 詳細, 対処)

    # 1. Python 本体
    v = sys.version_info
    bits = 64 if sys.maxsize > 2 ** 32 else 32
    results.append((True, "Python",
                    f"{v.major}.{v.minor}.{v.micro} ({bits}bit)", None))

    # 2. pywin32
    try:
        import win32com  # noqa: F401  （先頭 import 済みだが診断として明示確認）
        try:
            from importlib.metadata import version as _ver
            pv = _ver("pywin32")
        except Exception:
            pv = "(バージョン不明)"
        results.append((True, "pywin32", f"インストール済み {pv}", None))
    except ImportError:
        results.append((False, "pywin32", "見つかりません",
                        "py -m pip install pywin32 を実行してください（AI に「入れて」でも可）"))

    # 3. Excel のインストール（レジストリ確認・起動はしない）
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer") as k:
            curver = winreg.QueryValueEx(k, None)[0]     # 例: Excel.Application.16
        results.append((True, "Excel", f"インストール済み ({curver})", None))
        excel_installed = True
    except Exception:
        results.append((False, "Excel", "インストールが確認できません",
                        "Excel (デスクトップ版) が必要です"))
        excel_installed = False

    # 4. Excel の起動状態と VBOM（VBAプロジェクトへのアクセス信頼）
    xl = None
    if excel_installed:
        try:
            xl = _get_active_excel()
        except Exception:
            xl = None
    if xl is None:
        results.append((None, "Excel起動", "起動していません",
                        "VBOM 設定の診断には、Excel でブックを開いてから再実行してください"))
    else:
        try:
            wb_names = [w.Name for w in xl.Workbooks]
        except Exception:
            wb_names = []
        results.append((True, "Excel起動",
                        f"起動中（開いているブック: {', '.join(wb_names) or 'なし'}）", None))
        # VBOM: VBE にアクセスできるか（ブロックされていると例外になる）
        try:
            _ = xl.VBE.VBProjects.Count
            results.append((True, "VBOM信頼設定", "有効（VBAプロジェクトにアクセス可能）", None))
        except Exception:
            results.append((False, "VBOM信頼設定", "無効（VBAプロジェクトにアクセスできません）",
                            "Excel の [ファイル > オプション > トラストセンター > トラストセンターの設定 > "
                            "マクロの設定] で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」に"
                            "チェックを入れてください"))

    # 5. gen_py キャッシュの健全性（破損すると「Excelは起動していません」と誤報する既知問題）
    if xl is not None:
        try:
            win32com.client.GetActiveObject("Excel.Application")
            results.append((True, "COMキャッシュ(gen_py)", "正常", None))
        except Exception:
            results.append((False, "COMキャッシュ(gen_py)", "破損の疑い（キャッシュ経由の接続に失敗）",
                            r"%LOCALAPPDATA%\Temp\gen_py フォルダを削除すると自動再生成されます"))

    # 6. ツール一式の存在
    missing = [f for f in ("form_builder.py", "form_inspect.py",
                           "form_layout.py", "form_tool.py")
               if not os.path.exists(os.path.join(SCRIPT_DIR, f))]
    if missing:
        results.append((None, "ツール一式", f"見つからないファイル: {', '.join(missing)}",
                        "フォーム機能を使う場合は同じフォルダに配置してください（マクロ管理だけなら不要）"))
    else:
        results.append((True, "ツール一式", "vba_manager + フォーム4ツールが揃っています", None))

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": all(r[0] is not False for r in results),
                          "checks": [{"ok": r[0], "item": r[1], "detail": r[2],
                                      "fix": r[3]} for r in results]},
                         ensure_ascii=False), file=sys.stdout)
        return all(r[0] is not False for r in results)

    print("===== 導入セルフ診断 (setup-check) =====")
    ng = 0
    for ok, item, detail, fix in results:
        mark = "OK" if ok else ("--" if ok is None else "NG")
        if ok is False:
            ng += 1
        print(f"  [{mark}] {item}: {detail}")
        if fix and ok is not True:
            print(f"       → {fix}")
    print("-" * 44)
    if ng == 0:
        print("  問題なし。`py vba_manager.py list` から始められます。")
    else:
        print(f"  NG {ng}件。上の対処を実行してから再実行してください。")
        print("  分からないところは、導入しようとしている AI にこの出力を貼って聞いてください。")
    return ng == 0


def _collect_book_inventory(xl, wb, include_vba=True):
    """ブックの棚卸しデータを機械収集する（docs / call-graph の共通土台）。

    include_vba=False は VBA プロジェクトに触らない縮退モード
    （パスワード保護・VBOM 未信頼のブックをシート側だけ診るために使う）。
    """
    inv = {'name': wb.Name, 'fullname': wb.FullName}

    # シート
    sheets = []
    try:
        active = wb.ActiveSheet.Name
    except Exception:
        active = None
    for sh in wb.Sheets:
        d = {'name': sh.Name, 'active': sh.Name == active}
        try:
            d['visible'] = int(sh.Visible)
        except Exception:
            d['visible'] = -1
        try:
            ur = sh.UsedRange
            d['rows'] = ur.Rows.Count
            d['cols'] = ur.Columns.Count
            d['address'] = ur.Address
        except Exception:
            d['rows'] = d['cols'] = 0
            d['address'] = None
        sheets.append(d)
    inv['sheets'] = sheets

    # モジュールとプロシージャ（コード全文も保持＝call-graph が使う）
    proc_pat = re.compile(
        r'^\s*(?P<vis>Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?'
        r'(?P<kind>Sub|Function)\s+(?P<name>[^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE)
    type_names = {1: '標準', 2: 'クラス', 3: 'フォーム', 100: 'ブック/シート'}
    modules = []
    for comp in (wb.VBProject.VBComponents if include_vba else ()):
        cm = comp.CodeModule
        n = cm.CountOfLines
        code = cm.Lines(1, n) if n > 0 else ""
        lines = code.split('\r\n') if code else []
        procs = []
        for m in proc_pat.finditer(code):
            name = m.group('name')
            info = {'name': name,
                    'kind': m.group('kind').capitalize(),
                    'private': bool(m.group('vis') and 'private' in m.group('vis').lower())}
            try:
                info['lines'] = cm.ProcCountLines(name, 0)
                body = cm.ProcBodyLine(name, 0)
                if body < len(lines):
                    first = lines[body].strip()
                    if first.startswith("'"):
                        info['comment'] = first.lstrip("'").strip()
            except Exception:
                pass
            procs.append(info)
        modules.append({'name': comp.Name, 'type': int(comp.Type),
                        'type_name': type_names.get(int(comp.Type), str(comp.Type)),
                        'total_lines': n, 'procs': procs, 'code': code})
    inv['modules'] = modules

    # フォーム（コントロール数）
    forms = []
    for comp in (wb.VBProject.VBComponents if include_vba else ()):
        if int(comp.Type) != 3:
            continue
        d = {'name': comp.Name}
        try:
            d['caption'] = comp.Properties("Caption").Value
        except Exception:
            d['caption'] = None
        try:
            d['controls'] = comp.Designer.Controls.Count
        except Exception:
            d['controls'] = None
        forms.append(d)
    inv['forms'] = forms

    # ショートカット（Attribute 走査。エクスポート方式は list-shortcuts と同じ）
    shortcuts = {}
    attr_pat = re.compile(
        r'Attribute\s+([^.\s]+)\.VB_ProcData\.VB_Invoke_Func\s*=\s*"([^"]+)"',
        re.IGNORECASE | re.DOTALL)
    for comp in (wb.VBProject.VBComponents if include_vba else ()):
        if int(comp.Type) not in (1, 2, 3, 100):
            continue
        tmp = os.path.join(SCRIPT_DIR, f"_tmp_doc_{comp.Name}.bas")
        try:
            comp.Export(tmp)
            with open(tmp, 'rb') as f:
                content = f.read().decode('cp932', errors='replace')
            for m in attr_pat.finditer(content):
                raw = m.group(2).replace('\\n', '\n').replace('\\r', '\r')
                key = raw.split('\n')[0].split('\r')[0]
                if key:
                    # VB_Invoke_Func のキーが大文字なら Ctrl+Shift+キー の割当
                    if key.isalpha() and key == key.upper():
                        shortcuts[m.group(1)] = f"Ctrl+Shift+{key}"
                    else:
                        shortcuts[m.group(1)] = f"Ctrl+{key}"
        except Exception:
            pass
        finally:
            _remove_export_artifacts(tmp)
    inv['shortcuts'] = shortcuts

    # 図形・フォームコントロールに登録されたマクロ（OnAction）＝ボタンからの実行入口。
    # VBA ロック中でも読める（Shapes はシート側の情報）
    onaction = []
    for sh in wb.Worksheets:
        try:
            shapes = list(sh.Shapes)
        except Exception:
            continue
        for shp in shapes:
            try:
                oa = shp.OnAction
            except Exception:
                continue
            if oa:
                macro = oa.split('!')[-1].strip("'\" ")
                try:
                    shp_name = shp.Name
                except Exception:
                    shp_name = '(図形)'
                onaction.append((sh.Name, shp_name, macro))
    inv['onaction'] = onaction

    # テーブル・ピボット
    tables, pivots = [], []
    for sh in wb.Worksheets:
        try:
            for lo in sh.ListObjects:
                tables.append({'sheet': sh.Name, 'name': lo.Name,
                               'address': lo.Range.Address})
        except Exception:
            pass
        try:
            for pt in sh.PivotTables():
                pivots.append({'sheet': sh.Name, 'name': pt.Name})
        except Exception:
            pass
    inv['tables'] = tables
    inv['pivots'] = pivots

    # PowerQuery・接続・名前付き範囲
    queries = []
    try:
        for q in wb.Queries:
            queries.append({'name': q.Name})
    except Exception:
        pass
    inv['queries'] = queries
    conns = []
    try:
        for cn in wb.Connections:
            conns.append({'name': cn.Name})
    except Exception:
        pass
    inv['connections'] = conns
    names = []
    try:
        for nm in wb.Names:
            try:
                names.append({'name': nm.Name, 'refers_to': nm.RefersTo})
            except Exception:
                names.append({'name': nm.Name, 'refers_to': None})
    except Exception:
        pass
    inv['names'] = names
    return inv


def _inventory_or_explain(xl, wb):
    """棚卸しを試み、VBA に触れないブック（パスワード保護/VBOM未信頼）なら
    生の COM エラーで転ばず理由を説明して None を返す（docs/call-graph/impact 用）"""
    try:
        return _collect_book_inventory(xl, wb)
    except Exception as e:
        print("エラー: VBA プロジェクトに触れません（パスワード保護または VBOM 未信頼）。")
        print("  シート側だけの診断なら checkup（健康診断）が縮退モードで実行できます。")
        print(f"  詳細: {e}")
        return None


def cmd_docs(args):
    """ブックの取扱説明書を自動生成: docs [excel_file] [--out f.md]

    シート構成・モジュール別マクロ表（行数/ショートカット/先頭コメント）・
    フォーム・テーブル/ピボット/クエリ/接続/名前付き範囲を Markdown 1枚に棚卸しする。
    「このブックに何が入っているか」を機械が書く＝ブックと会話するための自己紹介文。
    """
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, readonly=True)   # 健診モード（診断は読むだけ）
    inv = _inventory_or_explain(xl, wb)
    if inv is None:
        return False

    if getattr(args, 'json', False):
        import json
        slim = {k: v for k, v in inv.items()}
        slim['modules'] = [{k: v for k, v in m.items() if k != 'code'}
                           for m in inv['modules']]
        print(json.dumps({"success": True, **slim}, ensure_ascii=False), file=sys.stdout)
        return True

    L = []
    total_procs = sum(len(m['procs']) for m in inv['modules'])
    total_lines = sum(m['total_lines'] for m in inv['modules'])
    L.append(f"# {inv['name']} の構成ドキュメント")
    L.append("")
    L.append(f"- 生成: {time.strftime('%Y-%m-%d %H:%M')}（vba_manager docs）")
    L.append(f"- パス: {inv['fullname']}")
    L.append(f"- シート {len(inv['sheets'])} / マクロ {total_procs}（{total_lines}行） / "
             f"フォーム {len(inv['forms'])} / テーブル {len(inv['tables'])} / "
             f"ピボット {len(inv['pivots'])} / クエリ {len(inv['queries'])}")
    L.append("")

    L.append("## シート")
    L.append("")
    L.append("| シート | 使用範囲 | 大きさ | 状態 |")
    L.append("|---|---|---|---|")
    vis_label = {-1: '', 0: '非表示', 2: '完全非表示'}
    for s in inv['sheets']:
        mark = '（アクティブ）' if s['active'] else ''
        L.append(f"| {s['name']}{mark} | {s['address'] or '-'} | "
                 f"{s['rows']}行×{s['cols']}列 | {vis_label.get(s['visible'], '')} |")
    L.append("")

    # --preview N: 各シートの先頭N行を Markdown 表で（初見ブックの中身の見取り）
    try:
        preview = int(getattr(args, 'preview', None) or 0)
    except (TypeError, ValueError):
        preview = 0
    if preview > 0:
        def _md_cell(v):
            return _cell_str(v).replace('|', '\\|').replace('\n', ' ')
        for sh in wb.Sheets:
            try:
                ur = sh.UsedRange
                nrows = min(preview, ur.Rows.Count)
                ncols = min(ur.Columns.Count, 12)   # 横に広すぎる表は12列で切る
                head = sh.Range(ur.Cells(1, 1), ur.Cells(nrows, ncols))
                rows_v = _range_values_2d(head)
            except Exception:
                continue
            L.append(f"### {sh.Name} の先頭{nrows}行")
            L.append("")
            start_col = ur.Column
            headers = [_col_letter(start_col + j) for j in range(ncols)]
            L.append("| 行 | " + " | ".join(headers) + " |")
            L.append("|---|" + "---|" * ncols)
            for ri, r in enumerate(rows_v):
                cells = [_md_cell(v) for v in r] + [''] * (ncols - len(r))
                L.append(f"| {ur.Row + ri} | " + " | ".join(cells[:ncols]) + " |")
            if ur.Columns.Count > ncols:
                L.append(f"（横は {ncols} 列まで表示・実際は {ur.Columns.Count} 列）")
            L.append("")

    L.append("## マクロ")
    L.append("")
    for m in inv['modules']:
        if not m['procs'] and m['total_lines'] == 0:
            continue
        L.append(f"### [{m['name']}]（{m['type_name']}・{len(m['procs'])}プロシージャ・{m['total_lines']}行）")
        if m['procs']:
            L.append("")
            L.append("| プロシージャ | 種別 | 行数 | ショートカット | 説明（先頭コメント） |")
            L.append("|---|---|---|---|---|")
            for p in m['procs']:
                sc = inv['shortcuts'].get(p['name'], '')
                priv = 'Private ' if p.get('private') else ''
                L.append(f"| {p['name']} | {priv}{p['kind']} | {p.get('lines', '')} | "
                         f"{sc} | {p.get('comment', '')} |")
        L.append("")

    if inv['forms']:
        L.append("## フォーム")
        L.append("")
        L.append("| フォーム | キャプション | コントロール数 |")
        L.append("|---|---|---|")
        for f in inv['forms']:
            L.append(f"| {f['name']} | {f.get('caption') or ''} | {f.get('controls') or ''} |")
        L.append("")

    def _simple_list(title, items, fmt):
        if not items:
            return
        L.append(f"## {title}")
        L.append("")
        for it in items:
            L.append(f"- {fmt(it)}")
        L.append("")

    _simple_list("テーブル", inv['tables'],
                 lambda t: f"{t['name']}（{t['sheet']} {t['address']}）")
    _simple_list("ピボットテーブル", inv['pivots'],
                 lambda t: f"{t['name']}（{t['sheet']}）")
    _simple_list("PowerQuery", inv['queries'], lambda t: t['name'])
    _simple_list("接続", inv['connections'], lambda t: t['name'])
    _simple_list("名前付き範囲", inv['names'],
                 lambda t: f"{t['name']} → {t.get('refers_to') or '?'}")

    out_path = getattr(args, 'out_opt', None)
    out_path = os.path.abspath(out_path) if out_path else os.path.join(SCRIPT_DIR, "_last_docs.md")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')
    print(f"構成ドキュメントを生成: {out_path}")
    print(f"  シート {len(inv['sheets'])} / マクロ {total_procs}（{total_lines}行） / "
          f"フォーム {len(inv['forms'])} / テーブル {len(inv['tables'])} / "
          f"ピボット {len(inv['pivots'])} / クエリ {len(inv['queries'])}")
    return True


def _analyze_calls(inv):
    """呼び出し関係の解析本体（call-graph / checkup 共用・COM 不要の純粋処理）。

    戻り値: {'known', 'edges', 'unresolved', 'orphans'}
    """
    known = {}
    form_modules = {m['name'] for m in inv['modules'] if m['type'] == 3}
    doc_modules = {m['name'] for m in inv['modules'] if m['type'] == 100}
    for m in inv['modules']:
        for p in m['procs']:
            known.setdefault(p['name'].lower(), (p['name'], m['name']))

    defn_pat = re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?'
                          r'(?:Sub|Function)\s+([^\s\(\)]+)', re.IGNORECASE)
    end_pat = re.compile(r'^\s*End\s+(?:Sub|Function)\b', re.IGNORECASE)
    call_pat = re.compile(r'\bCall\s+([^\s\(\):=]+)', re.IGNORECASE)
    # Declare 宣言された外部 API（Win32 等）は実在する呼び先＝未解決扱いにしない
    decl_pat = re.compile(r'^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?'
                          r'(?:Sub|Function)\s+([^\s\(]+)', re.IGNORECASE)
    api_names = set()
    for m in inv['modules']:
        for line in m['code'].split('\r\n'):
            dm = decl_pat.match(line)
            if dm:
                api_names.add(dm.group(1).lower())
    # Run の呼び先が丸ごと1つの文字列リテラルのときだけ静的に解決する。
    # 閉じクォート直後に & が続く（"…" & 変数）は動的呼び出し＝対象外
    any_run_pat = re.compile(r'Application\s*\.\s*Run\b', re.IGNORECASE)
    run_pat = re.compile(r'Application\s*\.\s*Run\s*\(?\s*"(?:[^"!]*!)?([^"]+)"(?!\s*&)',
                         re.IGNORECASE)
    # 裸呼び用: 全既知名の一括照合（ドット直後=他オブジェクトのメソッドは除外）
    if known:
        names_alt = '|'.join(sorted((re.escape(v[0]) for v in known.values()),
                                    key=len, reverse=True))
        # VBA の識別子は大文字小文字を区別しないため IGNORECASE で照合し、
        # マッチ後に known で正式名へ正規化する
        bare_pat = re.compile(r'(?<![\w.])(' + names_alt + r')(?![\w])', re.IGNORECASE)
    else:
        bare_pat = None

    edges = {}         # (caller_mod, caller_proc) -> {callee正式名, ...}
    unresolved = []    # (mod, proc, 呼び名, 行番号, 行テキスト)
    dynamic_runs = []  # (mod, proc, 行番号, 行テキスト) 呼び先が実行時に決まる Run
    for m in inv['modules']:
        cur = None
        for i, raw in enumerate(m['code'].split('\r\n'), 1):
            dm = defn_pat.match(raw)
            if dm:
                cur = dm.group(1)
                continue
            if end_pat.match(raw):
                cur = None
                continue
            # 文字列リテラルを潰してからコメントを落とす（' の誤爆防止）
            line = re.sub(r'"[^"]*"', '""', raw).split("'")[0]
            # Rem コメント行の Call/Run を生きた呼び出し扱いしない
            low = line.strip().lower()
            if low == 'rem' or low.startswith('rem '):
                continue
            caller = (m['name'], cur or '(宣言部)')

            for cm_ in call_pat.finditer(line):
                name = cm_.group(1)
                if '.' in name:
                    # Call obj.Method(...) はオブジェクトのメソッド。ただし
                    # Call モジュール名.マクロ名 の修飾呼びは末尾で解決を試みる
                    tail = name.rsplit('.', 1)[-1]
                    hit = known.get(tail.lower())
                    if hit:
                        edges.setdefault(caller, set()).add(hit[0])
                    continue
                hit = known.get(name.lower())
                if hit:
                    edges.setdefault(caller, set()).add(hit[0])
                elif name.lower() not in api_names:
                    unresolved.append((m['name'], cur, name, i, raw.strip()))
            ran_static = False
            # Run の名前は文字列内にあるのでコメントだけ除いた原文から拾う
            # （生 raw だとコメントアウトされた Run を誤検知する）
            for rm_ in run_pat.finditer(_strip_vba_comment(raw)):
                ran_static = True
                name = rm_.group(1)
                hit = known.get(name.lower())
                if hit is None and '.' in name:
                    # Run "モジュール名.マクロ名" のモジュール修飾は
                    # Call 側の修飾呼びと同じく末尾名で解決を試みる
                    hit = known.get(name.rsplit('.', 1)[-1].lower())
                if hit:
                    edges.setdefault(caller, set()).add(hit[0])
                else:
                    unresolved.append((m['name'], cur, f'Run "{name}"', i, raw.strip()))
            if not ran_static and any_run_pat.search(line):
                # リテラル1本で解決できない Run（"…" & 変数 / 変数のみ）＝動的呼び出し
                dynamic_runs.append((m['name'], cur or '(宣言部)', i, raw.strip()[:60]))
            if bare_pat:
                for bm in bare_pat.finditer(line):
                    hit = known.get(bm.group(1).lower())
                    name = hit[0] if hit else bm.group(1)   # 正式名へ正規化
                    if cur and name.lower() == cur.lower():
                        continue          # 自分自身の再帰は流れの把握には不要
                    edges.setdefault(caller, set()).add(name)

    # 図形・ボタンに登録されたマクロ（OnAction）＝ボタンからの実行入口
    onaction = {}
    for sheet, shp, macro in inv.get('onaction', ()):
        hit = known.get(macro.lower())
        if hit is None and '.' in macro:
            # 「モジュール名.マクロ名」形式の登録（同名マクロがあると Excel が
            # 自動でこの形式にする）も末尾名で解決する
            hit = known.get(macro.rsplit('.', 1)[-1].lower())
        if hit:
            onaction.setdefault(hit[0], []).append(f"{sheet}/{shp}")

    # 呼ばれる側の集合と孤立
    called = set()
    for callees in edges.values():
        called |= callees
    orphans = []
    for m in inv['modules']:
        for p in m['procs']:
            if p['name'] in called:
                continue
            if p['name'] in onaction:
                continue      # シート上のボタン/図形から呼ばれている＝孤立ではない
            # フォーム/ブック/シートのイベントプロシージャ、Private はイベント・内部用が
            # 多いので孤立には数えない（機械的な絞り込み）
            if m['name'] in form_modules or m['name'] in doc_modules:
                continue
            orphans.append((m['name'], p['name']))
    return {'known': known, 'edges': edges, 'onaction': onaction,
            'unresolved': unresolved, 'orphans': orphans,
            'dynamic_runs': dynamic_runs}


def _strip_vba_comment(raw):
    """行からコメント部を落とす（文字列リテラル内の ' は誤爆させない）"""
    in_str = False
    for i, ch in enumerate(raw):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return raw[:i]
    return raw


_AUTO_EXEC_STD = {'auto_open', 'auto_close'}
_AUTO_EXEC_PREFIXES = ('workbook_', 'worksheet_', 'chart_')


def _extra_code_scans(inv):
    """checkup の参考所見スキャン（COM 不要の純粋処理・事実の列挙のみ）。

    自動実行イベント / Option Explicit なし / On Error Resume Next /
    ハードコードされたパス / 長いプロシージャ / 破壊的な操作の所在 /
    ScreenUpdating・Calculation の戻し忘れ。いずれも要不要の判断はしない。
    """
    res = {'auto_exec': [], 'no_option_explicit': [], 'error_resume': [],
           'hardcoded_paths': [], 'long_procs': [], 'destructive': [],
           'no_restore': []}
    path_pat = re.compile(r'"((?:[A-Za-z]:\\|\\\\)[^"]{2,})"')
    oern_pat = re.compile(r'\bOn\s+Error\s+Resume\s+Next\b', re.IGNORECASE)
    defn_pat = re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?'
                          r'(?:Sub|Function)\s+([^\s\(\)]+)', re.IGNORECASE)
    end_pat = re.compile(r'^\s*End\s+(?:Sub|Function)\b', re.IGNORECASE)
    destr_pats = [
        (re.compile(r'(?<![\w.])(?:Kill|RmDir)\b', re.IGNORECASE), 'ファイル/フォルダ削除'),
        (re.compile(r'\.(?:DeleteFile|DeleteFolder|MoveFile|MoveFolder)\b',
                    re.IGNORECASE), 'FSOのファイル操作'),
        (re.compile(r'\b(?:Worksheets|Sheets)\s*\([^)]*\)\s*\.Delete\b'
                    r'|\bActiveSheet\s*\.Delete\b', re.IGNORECASE), 'シート削除'),
        (re.compile(r'(?:\bRows\b|\bColumns\b|\.EntireRow|\.EntireColumn)'
                    r'[^\n]*\.Delete\b', re.IGNORECASE), '行/列の削除'),
    ]
    su_off = re.compile(r'\bScreenUpdating\s*=\s*False\b', re.IGNORECASE)
    su_on = re.compile(r'\bScreenUpdating\s*=\s*True\b', re.IGNORECASE)
    ca_off = re.compile(r'\bCalculation\s*=\s*xl(?:Calculation)?Manual\b', re.IGNORECASE)
    ca_on = re.compile(r'\bCalculation\s*=\s*xl(?:Calculation)?Automatic\b', re.IGNORECASE)
    ee_off = re.compile(r'\bEnableEvents\s*=\s*False\b', re.IGNORECASE)
    ee_on = re.compile(r'\bEnableEvents\s*=\s*True\b', re.IGNORECASE)

    for m in inv['modules']:
        code = m['code']
        if not code:
            continue
        lines = code.split('\r\n')
        if not any(re.match(r'\s*Option\s+Explicit\b', ln, re.IGNORECASE)
                   for ln in lines):
            res['no_option_explicit'].append(m['name'])
        for p in m['procs']:
            nl = p['name'].lower()
            if (m['type'] == 100 and nl.startswith(_AUTO_EXEC_PREFIXES)) or \
               (m['type'] == 1 and nl in _AUTO_EXEC_STD):
                res['auto_exec'].append((m['name'], p['name'], p.get('lines')))
            if p.get('lines') and p['lines'] >= 150:
                res['long_procs'].append((m['name'], p['name'], p['lines']))

        cur = None
        state = {'su_off': False, 'su_on': False, 'ca_off': False, 'ca_on': False,
                 'ee_off': False, 'ee_on': False}

        def flush(proc, _m=m, _state=state):
            # プロシージャ末尾で ScreenUpdating/Calculation/EnableEvents の戻し忘れを確定する
            if proc:
                if _state['su_off'] and not _state['su_on']:
                    res['no_restore'].append(
                        (_m['name'], proc, 'ScreenUpdating を False にしたまま True に戻す行がない'))
                if _state['ca_off'] and not _state['ca_on']:
                    res['no_restore'].append(
                        (_m['name'], proc, 'Calculation を手動にしたまま自動に戻す行がない'))
                if _state['ee_off'] and not _state['ee_on']:
                    res['no_restore'].append(
                        (_m['name'], proc, 'EnableEvents を False にしたまま True に戻す行がない'
                                           '（イベントが死んだままになる）'))
            for k in _state:
                _state[k] = False

        for i, raw in enumerate(lines, 1):
            dm = defn_pat.match(raw)
            if dm:
                flush(cur)
                cur = dm.group(1)
            elif end_pat.match(raw):
                flush(cur)
                cur = None
            body = _strip_vba_comment(raw)
            blanked = re.sub(r'"[^"]*"', '""', body)
            for pm in path_pat.finditer(body):
                res['hardcoded_paths'].append((m['name'], cur or '(宣言部)', i,
                                               pm.group(1)))
            if oern_pat.search(blanked):
                res['error_resume'].append((m['name'], cur or '(宣言部)', i))
            for pat, label in destr_pats:
                if pat.search(blanked):
                    res['destructive'].append((m['name'], cur or '(宣言部)', i,
                                               label, body.strip()[:60]))
            if su_off.search(blanked):
                state['su_off'] = True
            if su_on.search(blanked):
                state['su_on'] = True
            if ca_off.search(blanked):
                state['ca_off'] = True
            if ca_on.search(blanked):
                state['ca_on'] = True
            if ee_off.search(blanked):
                state['ee_off'] = True
            if ee_on.search(blanked):
                state['ee_on'] = True
        flush(cur)
    res['long_procs'].sort(key=lambda t: -t[2])
    return res


CHECKUP_HISTORY_DIR = os.path.join(SCRIPT_DIR, "_checkup_history")


def _checkup_history_path(book_name):
    safe = re.sub(r'[\\/:*?"<>|]', '_', book_name)
    return os.path.join(CHECKUP_HISTORY_DIR, safe + ".json")


def _load_checkup_history(book_name):
    """ブック別の診断履歴（新しい順でなく古い順のリスト）を読む。無ければ空"""
    import json
    try:
        with open(_checkup_history_path(book_name), 'r', encoding='utf-8') as f:
            hist = json.load(f)
        if not isinstance(hist, list):
            return []
        # 要素の型まで守る（壊れた履歴ファイルで毎回診断が落ちるのを防ぐ）
        return [h for h in hist if isinstance(h, dict)]
    except (OSError, ValueError):
        return []


def _save_checkup_history(book_name, hist, keep=12):
    import json
    os.makedirs(CHECKUP_HISTORY_DIR, exist_ok=True)
    with open(_checkup_history_path(book_name), 'w', encoding='utf-8') as f:
        json.dump(hist[-keep:], f, ensure_ascii=False, indent=1)


# --- 確認済み（意図的）所見の記録 ---
# 「フォント混在」「同じ行のボタン幅」等、目で見て意図的なデザインと判断した所見は
# ここに登録すると次回以降の「所見サマリ」の件数・総合判定から除外される
# （消すのではなく「## 確認済み（意図的・件数から除外）」セクションに移すだけ＝事実は残す）。
CHECKUP_ACK_DIR = os.path.join(SCRIPT_DIR, "_checkup_ack")


def _checkup_ack_path(book_name):
    safe = re.sub(r'[\\/:*?"<>|]', '_', book_name)
    return os.path.join(CHECKUP_ACK_DIR, safe + ".json")


def _load_checkup_ack(book_name):
    """確認済み（意図的）として除外する所見キーの集合を読む。無ければ空集合"""
    import json
    try:
        with open(_checkup_ack_path(book_name), 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, list):
            return set()
        # 文字列以外が混ざると保存時の sorted() で TypeError になるため除外
        return set(x for x in data if isinstance(x, str))
    except (OSError, ValueError):
        return set()


def _save_checkup_ack(book_name, ack_set):
    import json
    os.makedirs(CHECKUP_ACK_DIR, exist_ok=True)
    with open(_checkup_ack_path(book_name), 'w', encoding='utf-8') as f:
        json.dump(sorted(ack_set), f, ensure_ascii=False, indent=1)


def _checkup_diff(prev, cur):
    """前回スナップショットとの差分（純粋処理）。prev が無ければ None。

    所見キーは行番号を含まない形にしてあるので、行ずれでは差分にならない。
    """
    if not prev:
        return None
    # 履歴ファイルは外部データ＝どちら側もキー欠損に耐える（cur も .get で読む）
    pk, ck = set(prev.get('keys', [])), set(cur.get('keys', []))
    d = {'prev_time': prev.get('time'),
         'new': sorted(ck - pk), 'resolved': sorted(pk - ck)}
    # 履歴 JSON は外部データ。procs の値が list 以外に汚損していても診断ごと
    # 落とさない（キー欠損への耐性と同じ方針で型汚損にも耐える）
    def _safe(d):
        return {m: set(v) for m, v in (d.get('procs') or {}).items()
                if isinstance(v, (list, tuple, set))}
    pp = _safe(prev)
    cp = _safe(cur)
    added, removed = [], []
    for m in sorted(cp):
        for p in sorted(cp[m] - pp.get(m, set())):
            added.append(f"[{m}] {p}")
    for m in sorted(pp):
        for p in sorted(pp[m] - cp.get(m, set())):
            removed.append(f"[{m}] {p}")
    d['procs_added'], d['procs_removed'] = added, removed
    cur_lines = cur.get('total_lines', 0)
    d['lines_delta'] = cur_lines - prev.get('total_lines', cur_lines)
    cur_sheets, cur_forms = set(cur.get('sheets', [])), set(cur.get('forms', []))
    d['sheets_added'] = sorted(cur_sheets - set(prev.get('sheets', [])))
    d['sheets_removed'] = sorted(set(prev.get('sheets', [])) - cur_sheets)
    d['forms_added'] = sorted(cur_forms - set(prev.get('forms', [])))
    d['forms_removed'] = sorted(set(prev.get('forms', [])) - cur_forms)
    d['changed'] = bool(d['new'] or d['resolved'] or added or removed
                        or d['lines_delta'] or d['sheets_added'] or d['sheets_removed']
                        or d['forms_added'] or d['forms_removed'])
    return d


def _vba_references(wb):
    """VBA 参照設定の一覧と破損（MISSING）の検出（読み取りのみ）。

    眠っていたブックが動かない原因の筆頭＝「参照不可: ライブラリが見つかりません」を
    VBE を開かずに検出する。
    """
    total, broken = 0, []
    try:
        refs = wb.VBProject.References
    except Exception:
        return {'total': 0, 'broken': []}
    for r in refs:
        total += 1
        try:
            name = r.Name
        except Exception:
            name = "(名前不明)"
        try:
            is_broken = bool(r.IsBroken)
        except Exception:
            is_broken = True      # IsBroken すら読めない参照は破損として報告
        if is_broken:
            try:
                path = r.FullPath
            except Exception:
                path = None
            broken.append((name, path))
    return {'total': total, 'broken': broken}


def _sheet_health(wb):
    """シート側の健診（読み取りのみ）: エラーセル数。

    ※ ゴースト（使用範囲と実データの差）検査は撤去した（2026-07-05・shuさん指示）。
    値の無い行は「意図して罫線を引いた記入欄」のことが普通にあり（例:
    ファイル一覧.xlsm の修正シート＝連続修正の記入欄100行）、機械には
    取り残しと区別できない。この種の自動判定を根拠にした肥大縮小で
    取説シートの結合を壊した実害もある（2026-06-19）。復活させないこと。
    """
    out = {'error_cells': []}
    for sh in wb.Worksheets:
        try:
            ur = sh.UsedRange
        except Exception:
            continue
        n_err = 0
        for cell_type in (-4123, 2):          # 数式(-4123) / 定数(2)
            try:
                n_err += ur.SpecialCells(cell_type, 16).Count   # 16 = xlErrors
            except Exception:
                pass                           # 該当なしは SpecialCells が例外を投げる
        if n_err:
            out['error_cells'].append((sh.Name, n_err))
    return out


def _broken_links(wb):
    """外部ブックへのリンクのうち、参照先ファイルが存在しないもの（読み取りのみ）"""
    broken = []
    try:
        links = wb.LinkSources(1)              # 1 = xlExcelLinks
    except Exception:
        links = None
    for src in (links or ()):
        s = str(src)
        if re.match(r'^[A-Za-z]:\\|^\\\\', s) and not os.path.exists(s):
            broken.append(s)
    return broken


def _checkup_rating(total_findings, n_critical):
    """総合判定（機械的な分類。判断はしない）:
    C=壊れた参照・呼び出しあり / B=その他の所見あり / A=所見なし"""
    if n_critical:
        return "C（要確認）"
    if total_findings:
        return "B（軽度所見）"
    return "A（異常なし）"


def cmd_call_graph(args):
    """マクロの呼び出し関係を解析: call-graph [excel_file] [--macro 名]

    Call 文・Application.Run・既知プロシージャ名の裸呼びを機械的に解析する。
    - 未解決 Call: **存在しないマクロを呼んでいる行**（コピペ残骸の一語バグ検出器）
    - 呼び出し関係: どのマクロがどのマクロを使っているか
    - 孤立: どこからも呼ばれていないマクロ（メニュー/イベント直実行の可能性があるため
      機械は事実だけ報告し、要不要の判断はしない）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    focus = getattr(args, 'macro_opt', None)
    xl, wb = get_workbook(target_file, readonly=True)   # 健診モード（診断は読むだけ）
    inv = _inventory_or_explain(xl, wb)
    if inv is None:
        return False
    res = _analyze_calls(inv)
    known, edges = res['known'], res['edges']
    unresolved, orphans = res['unresolved'], res['orphans']

    if getattr(args, 'json', False):
        import json
        print(json.dumps({
            "success": True, "file": inv['name'],
            "edges": [{"caller_module": k[0], "caller": k[1], "callees": sorted(v)}
                      for k, v in sorted(edges.items())],
            "unresolved": [{"module": u[0], "proc": u[1], "name": u[2],
                            "line": u[3], "text": u[4]} for u in unresolved],
            "orphans": [{"module": o[0], "name": o[1]} for o in orphans],
        }, ensure_ascii=False), file=sys.stdout)
        return not unresolved

    if getattr(args, 'mermaid', None) is not None:
        # Mermaid 図（GitHub / Qiita でそのまま描画される）。モジュール別 subgraph、
        # 未解決の呼び先は赤ノード。呼び出しのあるマクロだけを図に載せる
        out_path = (os.path.join(SCRIPT_DIR, "_last_callgraph.md")
                    if args.mermaid == '_DEFAULT_' else os.path.abspath(args.mermaid))
        node_ids = {}

        def nid(name):
            if name not in node_ids:
                node_ids[name] = f"n{len(node_ids)}"
            return node_ids[name]

        used = set()
        for (mod, proc), callees in edges.items():
            if proc == '(宣言部)' or not callees:
                continue
            used.add(proc)
            used |= callees
        # 未解決Callしか持たないマクロもノード定義に載せる（辺だけ出すと
        # Mermaid が無ラベルの自動ノードを作り、肝心の呼び元が図から読めない）
        for _, _proc, _n, _, _ in unresolved:
            if _proc and _proc != '(宣言部)':
                used.add(_proc)
        ml = [f"# {inv['name']} 呼び出し関係図", "",
              f"生成: {time.strftime('%Y-%m-%d %H:%M')}（vba_manager call-graph --mermaid）", "",
              "```mermaid", "flowchart LR"]
        by_mod = {}
        for m in inv['modules']:
            for p in m['procs']:
                if p['name'] in used:
                    by_mod.setdefault(m['name'], []).append(p['name'])
        for mod, procs in by_mod.items():
            ml.append(f'    subgraph {mod}')
            for p in procs:
                ml.append(f'        {nid(p)}["{p}"]')
            ml.append('    end')
        for (mod, proc), callees in sorted(edges.items()):
            if proc == '(宣言部)':
                continue
            for c in sorted(callees):
                ml.append(f'    {nid(proc)} --> {nid(c)}')
        bad_ids = []
        for _, proc, name, _, _ in unresolved:
            if proc:
                bid = nid(f"？{name}")
                bad_ids.append(bid)
                ml.append(f'    {bid}["{name}（存在しない）"]')
                ml.append(f'    {nid(proc)} -.-> {bid}')
        for b in set(bad_ids):
            ml.append(f'    style {b} fill:#ffcccc,stroke:#cc0000')
        ml.append("```")
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(ml) + '\n')
        print(f"Mermaid 図を生成: {out_path}  "
              f"（ノード {len(node_ids)} / 未解決 {len(unresolved)}）")
        return not unresolved

    print(f"===== 呼び出し関係の解析: {inv['name']} =====")

    # 1. 未解決（最重要＝存在しないマクロを呼んでいる）
    if unresolved:
        print(f"\n⚠ 未解決の呼び出し（{len(unresolved)}件）— 存在しないマクロを呼んでいます:")
        for mod, proc, name, ln, text in unresolved:
            print(f"  [{mod}] {proc or '(宣言部)'} :{ln}  →  {name}")
            print(f"      {text}")
    else:
        print("\n未解決の呼び出し: なし（Call/Run はすべて実在のマクロを指しています）")

    # 2. 呼び出し関係（focus 指定ならそのマクロを起点にツリー展開）
    if focus:
        hit = known.get(focus.lower())
        if not hit:
            print(f"\nエラー: マクロ '{focus}' が見つかりません")
            _suggest_similar(focus, [v[0] for v in known.values()])
            return False
        root = hit[0]
        print(f"\n--- {root} からの呼び出しツリー ---")

        # 正式名→callee集合（モジュール横断で合算）
        by_name = {}
        for (mod, proc), callees in edges.items():
            by_name.setdefault(proc, set()).update(callees)

        def walk(name, depth, seen):
            for callee in sorted(by_name.get(name, ())):
                loop = "（循環）" if callee in seen else ""
                print("  " * depth + f"└ {callee}{loop}")
                if not loop and depth < 6:
                    walk(callee, depth + 1, seen | {callee})
        walk(root, 1, {root})
        callers = sorted({f"[{k[0]}] {k[1]}" for k, v in edges.items() if root in v})
        print(f"--- {root} を呼んでいるマクロ ---")
        for c in callers or ["  (なし)"]:
            print(f"  {c}" if not c.startswith("  ") else c)
    else:
        callers_with_edges = [(k, v) for k, v in sorted(edges.items())
                              if v and k[1] != '(宣言部)']
        print(f"\n--- 呼び出し関係（{len(callers_with_edges)}マクロが他マクロを使用） ---")
        for (mod, proc), callees in callers_with_edges:
            print(f"  [{mod}] {proc} → {', '.join(sorted(callees))}")

        if orphans:
            print(f"\n--- どこからも呼ばれていないマクロ（{len(orphans)}件・"
                  "メニュー/ショートカット直実行の可能性あり） ---")
            for mod, name in orphans[:40]:
                sc = inv['shortcuts'].get(name)
                print(f"  [{mod}] {name}" + (f"  ({sc})" if sc else ""))
            if len(orphans) > 40:
                print(f"  … 他 {len(orphans) - 40}件")
    return not unresolved


def _checkup_one(xl, wb, note=None, ack_all=False):
    """1ブックぶんの健康診断（収集・検査・履歴比較まで）。

    note はカルテのメモ＝今回のスナップショットに添付して履歴に残す
    （「何をどう直した回か」を後から思い出すための治療欄）。
    ack_all=True のとき、今回のコード/フォーム所見を全て「確認済み（意図的）」
    として登録し、以降の所見サマリ・総合判定から除外する
    （存在しないマクロ呼び出し等の致命的所見は対象外＝常に表示する）。
    戻り値: {'name','rating','total','counts','diff','lines','json'}
    レポートの書き出しと画面表示は cmd_checkup 側で行う。
    """
    # VBA プロジェクトに触れないブック（パスワード保護・VBOM 未信頼）は
    # シート側だけの縮退診断に切り替える（入口で転ばない）
    vba_error = None
    try:
        inv = _collect_book_inventory(xl, wb)
    except Exception as e:
        vba_error = str(e).strip() or "VBAプロジェクトにアクセスできません"
        inv = _collect_book_inventory(xl, wb, include_vba=False)
    calls = _analyze_calls(inv)
    extra = _extra_code_scans(inv)
    refs = _vba_references(wb) if not vba_error else {'total': 0, 'broken': []}
    sheet_h = _sheet_health(wb)
    dead_links = _broken_links(wb)
    ref_names = [n['name'] for n in inv['names']
                 if n.get('refers_to') and '#REF!' in str(n.get('refers_to'))]

    # --- コードの機械検査（モジュール単位） ---
    # (比較キー, 表示文) のペア。キーは行番号を含めない＝行ずれで差分を出さない
    code_findings = []
    for m in inv['modules']:
        if not m['code']:
            continue
        norm = m['code'].replace('\r\n', '\n')
        dups = _find_duplicate_procedures(norm)
        for nm, lns in dups.items():
            code_findings.append((f"重複プロシージャ: [{m['name']}] {nm}",
                                  f"[{m['name']}] プロシージャ名の重複: {nm}"
                                  f"（行 {', '.join(map(str, lns))}）"))
        for ln, s in _find_consecutive_dup_lines(norm)[:10]:
            code_findings.append((f"連続重複行: [{m['name']}] 「{s[:40]}」",
                                  f"[{m['name']}] 連続する同一コード行: 行{ln} 「{s[:40]}」"))

    # --- フォームの検査（form_inspect の lint を統合） ---
    form_findings = {}
    forms_checked = 0
    try:
        if vba_error:
            raise ImportError     # ロック中はフォームにも触れない（未検査扱い）
        import form_inspect as _fi
        for comp in wb.VBProject.VBComponents:
            if int(comp.Type) != 3:
                continue
            forms_checked += 1
            try:
                info = _fi.get_form_info(comp)
                controls = _fi.collect_controls(comp)
                cm = comp.CodeModule
                code_txt = cm.Lines(1, cm.CountOfLines) if cm.CountOfLines > 0 else None
                f = _fi.lint_form(comp.Name, info, controls, code=code_txt)
                if f:
                    form_findings[comp.Name] = f
            except Exception as e:
                form_findings[comp.Name] = [f"(検査不可: {e})"]
    except ImportError:
        form_findings = None      # form_inspect.py が無い構成（マクロ管理のみ）

    # --- バックアップ状況 ---
    # 部分一致（base in name）だと「家計」が「家計簿2026」のバックアップまで
    # 数えてしまう。make_backup / make_module_backup の命名規則に揃えて数える
    base = os.path.splitext(inv['name'])[0]
    book_prefix = inv['name'] + ".backup_before_"   # ブック丸ごと
    mod_prefix = base + "_"                         # モジュール単位 (.bas)
    bk_count, bk_latest = 0, None
    if os.path.isdir(BACKUP_DIR):
        for name in os.listdir(BACKUP_DIR):
            if name.startswith(book_prefix) or name.startswith(mod_prefix):
                bk_count += 1
                t = os.path.getmtime(os.path.join(BACKUP_DIR, name))
                if bk_latest is None or t > bk_latest:
                    bk_latest = t

    # --- 確認済み（意図的）所見の除外 ---
    # コード検査・フォーム検査の所見だけが対象（存在しないマクロ呼び出し等の
    # 致命的な所見は対象外＝常に表示する）。ack_all 指定時はここで今回の
    # 所見を全て確認済みとして登録してから、その状態でフィルタする。
    ack_set = _load_checkup_ack(inv['name'])
    if ack_all:
        new_keys = {k for k, _ in code_findings}
        for fname, findings in (form_findings or {}).items():
            new_keys |= {f"フォーム {fname}: {s}" for s in findings}
        ack_new = new_keys - ack_set
        ack_set = ack_set | new_keys
        if ack_new:
            _save_checkup_ack(inv['name'], ack_set)
    else:
        ack_new = set()

    code_findings_active = [(k, t) for k, t in code_findings if k not in ack_set]
    code_findings_acked = [(k, t) for k, t in code_findings if k in ack_set]
    form_findings_active, form_findings_acked = {}, {}
    for fname, findings in (form_findings or {}).items():
        act, ackd = [], []
        for s in findings:
            (ackd if f"フォーム {fname}: {s}" in ack_set else act).append(s)
        if act:
            form_findings_active[fname] = act
        if ackd:
            form_findings_acked[fname] = ackd
    n_ack = len(code_findings_acked) + sum(len(v) for v in form_findings_acked.values())

    unresolved = calls['unresolved']
    n_form = sum(len(v) for v in form_findings_active.values())
    n_critical = (len(unresolved) + len(refs['broken'])
                  + len(dead_links) + len(ref_names))
    total_findings = n_critical + len(code_findings_active) + n_form
    rating = _checkup_rating(total_findings, n_critical)
    if vba_error:
        rating = "C（要確認）" if n_critical else "判定保留（VBA未検査）"
    total_procs = sum(len(m['procs']) for m in inv['modules'])
    total_lines = sum(m['total_lines'] for m in inv['modules'])
    n_err_cells = sum(n for _, n in sheet_h['error_cells'])

    # --- 定期健診: 前回スナップショットと比較して履歴に積む ---
    finding_keys = []
    for mod, proc, name, ln, text in unresolved:
        finding_keys.append(f"未解決Call: [{mod}] {proc or '(宣言部)'} → {name}")
    for name, path in refs['broken']:
        finding_keys.append(f"参照切れ: {name}")
    for s in dead_links:
        finding_keys.append(f"リンク切れ: {s}")
    for nm in ref_names:
        finding_keys.append(f"#REF!名前: {nm}")
    finding_keys += [k for k, _ in code_findings]
    for fname, findings in (form_findings or {}).items():
        finding_keys += [f"フォーム {fname}: {s}" for s in findings]

    # --- 確認済みリストの自動整理 ---
    # lintのロジック変更（閾値追加・チェック撤去等）で二度と出ない古いキーが
    # 確認済みリストに残り続けないよう、今回の生の所見に無いキーは毎回自動で
    # 除く。ただし検査が不完全な回は整理しない：
    #   - vba_error 時（縮退診断＝コード/フォームを見られていない）
    #   - form_findings is None 時（form_inspect.py が無い構成。フォームを
    #     検査できていないだけなのに、フォーム系の確認済みを「古いキー」と
    #     誤判定して全消しし、次にフォーム検査できる環境で全部復活してしまう）
    n_ack_pruned = 0
    if not vba_error and form_findings is not None:
        stale_ack = ack_set - set(finding_keys)
        if stale_ack:
            _save_checkup_ack(inv['name'], ack_set - stale_ack)
            n_ack_pruned = len(stale_ack)

    snap = {
        'time': time.strftime('%Y-%m-%d %H:%M'),
        'counts': {'unresolved': len(unresolved), 'code': len(code_findings_active),
                   'form': n_form, 'orphans': len(calls['orphans'])},
        'keys': sorted(finding_keys),
        'procs': {m['name']: [p['name'] for p in m['procs']] for m in inv['modules']},
        'total_procs': total_procs, 'total_lines': total_lines,
        'total_findings': total_findings, 'rating': rating, 'ack_count': n_ack,
        'sheets': [s['name'] for s in inv['sheets']],
        'forms': [f['name'] for f in inv['forms']],
    }
    if note:
        snap['note'] = note
    if vba_error:
        diff = None       # ロック中の縮退診断は履歴に混ぜない（解除後の差分が乱れる）
    else:
        hist = _load_checkup_history(inv['name'])
        diff = _checkup_diff(hist[-1] if hist else None, snap)
        try:
            _save_checkup_history(inv['name'], hist + [snap])
        except OSError as e:
            print(f"警告: 診断履歴を保存できませんでした（{e}）。比較機能は次回も初回扱いになります")

    json_payload = {
        "success": total_findings == 0, "file": inv['name'], "rating": rating,
        "total_findings": total_findings,
        "unresolved_calls": len(unresolved),
        "broken_references": [{"name": n, "path": p} for n, p in refs['broken']],
        "broken_links": dead_links, "ref_error_names": ref_names,
        "code_findings": [d for _, d in code_findings_active],
        "form_findings": form_findings_active, "orphans": len(calls['orphans']),
        "acknowledged_count": n_ack, "newly_acknowledged": len(ack_new),
        "backups": bk_count,
        "error_cells": [{"sheet": s, "count": n} for s, n in sheet_h['error_cells']],
        "auto_exec": len(extra['auto_exec']),
        "hardcoded_paths": len(extra['hardcoded_paths']),
        "no_option_explicit": extra['no_option_explicit'],
        "error_resume": len(extra['error_resume']),
        "long_procs": len(extra['long_procs']),
        "destructive": len(extra['destructive']),
        "no_restore": len(extra['no_restore']),
        "onaction": calls['onaction'],
        "dynamic_runs": len(calls['dynamic_runs']),
        "vba_error": vba_error,
        "diff": diff,
    }
    L = []
    L.append(f"# {inv['name']} 健康診断レポート")
    L.append("")
    L.append(f"- 診断日時: {time.strftime('%Y-%m-%d %H:%M')}（vba_manager checkup）")
    if vba_error:
        L.append(f"- 規模: シート {len(inv['sheets'])} / マクロ・フォームは VBA 未検査のため不明")
    else:
        L.append(f"- 規模: シート {len(inv['sheets'])} / マクロ {total_procs}（{total_lines}行） / "
                 f"フォーム {len(inv['forms'])}")
    L.append(f"- 総合判定: **{rating}**"
             "（A=所見なし / B=所見あり / C=壊れた参照・呼び出しあり。機械的な分類です）")
    if vba_error:
        L.append("- ⚠ **VBA は検査できませんでした**（プロジェクトのパスワード保護、"
                 "または VBOM 未信頼）。以下はシート側だけの縮退診断です")
    L.append("")

    L.append("## 所見サマリ")
    L.append("")
    mark = "✅" if total_findings == 0 else "⚠"
    L.append(f"{mark} **検出された所見: {total_findings}件**"
             + (f"（確認済み・意図的として除外: {n_ack}件 → 末尾参照）" if n_ack else ""))
    if ack_new:
        L.append(f"　今回 {len(ack_new)}件を新たに確認済み登録しました。")
    L.append("")
    L.append(f"| 診察項目 | 所見 |")
    L.append(f"|---|---|")
    if vba_error:
        L.append(f"| 外部ブックへのリンク切れ | **{len(dead_links)}件** |")
        L.append(f"| #REF! になった名前付き範囲 | **{len(ref_names)}件** |")
        L.append(f"| エラーセル（#REF!/#NAME?等） | {n_err_cells}個/"
                 f"{len(sheet_h['error_cells'])}シート（参考） |")
        L.append(f"| ボタン/図形に登録されたマクロ | {len(inv['onaction'])}箇所（参考） |")
        L.append("| VBAの検査（コード/フォーム/参照設定） | 実施できず（保護または未信頼） |")
    else:
        L.append(f"| 存在しないマクロへの呼び出し | **{len(unresolved)}件** |")
        L.append(f"| 参照設定の破損（MISSING） | **{len(refs['broken'])}件**（全{refs['total']}参照） |")
        L.append(f"| 外部ブックへのリンク切れ | **{len(dead_links)}件** |")
        L.append(f"| #REF! になった名前付き範囲 | **{len(ref_names)}件** |")
        L.append(f"| コードの機械検査（重複等） | {len(code_findings_active)}件 |")
        if form_findings is not None:
            L.append(f"| フォームの検査（{forms_checked}フォーム） | {n_form}件 |")
        if n_ack:
            L.append(f"| 確認済み（意図的・上記から除外済み） | {n_ack}件 |")
        L.append(f"| エラーセル（#REF!/#NAME?等） | {n_err_cells}個/"
                 f"{len(sheet_h['error_cells'])}シート（参考） |")
        L.append(f"| 破壊的な操作（Kill/シート削除等） | {len(extra['destructive'])}箇所（参考） |")
        L.append(f"| ScreenUpdating/Calculation/EnableEvents 戻し忘れ | {len(extra['no_restore'])}件（参考） |")
        L.append(f"| どこからも呼ばれていないマクロ | {len(calls['orphans'])}件（参考・"
                 "メニュー/直実行の可能性あり） |")
        L.append(f"| ボタン/図形に登録されたマクロ | {len(calls['onaction'])}件（参考） |")
        L.append(f"| 動的な Application.Run（実行時に呼び先決定） | "
                 f"{len(calls['dynamic_runs'])}箇所（参考） |")
        L.append(f"| 自動実行イベント（開く/変更等で起動） | {len(extra['auto_exec'])}件（参考） |")
        L.append(f"| ハードコードされたパス | {len(extra['hardcoded_paths'])}箇所（参考） |")
        L.append(f"| Option Explicit なしのモジュール | {len(extra['no_option_explicit'])}件（参考） |")
        L.append(f"| On Error Resume Next | {len(extra['error_resume'])}箇所（参考） |")
        L.append(f"| 150行を超えるプロシージャ | {len(extra['long_procs'])}件（参考） |")
    L.append(f"| バックアップ | {bk_count}件"
             + (f"（最新: {time.strftime('%Y-%m-%d %H:%M', time.localtime(bk_latest))}）"
                if bk_latest else "（なし）") + " |")
    L.append("")

    L.append("## 前回との比較（定期健診）")
    L.append("")
    if vba_error:
        L.append("VBA 未検査の縮退診断のため、履歴への記録と前回比較は行いません。")
    elif diff is None:
        L.append("初回の診断のため比較対象がありません（今回の結果を記録しました。"
                 "次回からここに前回との差分が出ます）。")
    elif not diff['changed']:
        L.append(f"前回（{diff['prev_time']}）から変化はありません。")
    else:
        L.append(f"前回の診断: {diff['prev_time']}")
        L.append("")
        for s in diff['new']:
            L.append(f"- ＋ 新しい所見: {s}")
        for s in diff['resolved']:
            L.append(f"- － 解消した所見: {s}")
        for s in diff['procs_added']:
            L.append(f"- ＋ マクロ追加: {s}")
        for s in diff['procs_removed']:
            L.append(f"- － マクロ削除: {s}")
        if diff['lines_delta']:
            sign = '+' if diff['lines_delta'] > 0 else ''
            L.append(f"- コード行数: {sign}{diff['lines_delta']}行（計{total_lines}行）")
        for s in diff['sheets_added']:
            L.append(f"- ＋ シート追加: {s}")
        for s in diff['sheets_removed']:
            L.append(f"- － シート削除: {s}")
        for s in diff['forms_added']:
            L.append(f"- ＋ フォーム追加: {s}")
        for s in diff['forms_removed']:
            L.append(f"- － フォーム削除: {s}")
    L.append("")

    if unresolved:
        L.append("## ⚠ 存在しないマクロへの呼び出し（最優先で確認）")
        L.append("")
        for mod, proc, name, ln, text in unresolved:
            L.append(f"- `[{mod}] {proc or '(宣言部)'}` の {ln} 行目 → **{name}**")
            L.append(f"  ```vba")
            L.append(f"  {text}")
            L.append(f"  ```")
        L.append("")

    if refs['broken'] or dead_links or ref_names:
        L.append("## ⚠ 壊れた参照（最優先で確認）")
        L.append("")
        for name, path in refs['broken']:
            L.append(f"- 参照設定の破損（MISSING）: **{name}**"
                     + (f" `{path}`" if path else ""))
        if refs['broken']:
            L.append("  （VBE の [ツール]→[参照設定] で「参照不可」になっている項目。"
                     "コンパイルエラーの典型原因）")
        for s in dead_links:
            L.append(f"- 外部ブックへのリンク切れ: `{s}`（参照先ファイルが存在しません）")
        for nm in ref_names:
            L.append(f"- #REF! になった名前付き範囲: **{nm}**")
        L.append("")

    if code_findings_active:
        L.append("## コードの機械検査")
        L.append("")
        for _, s in code_findings_active:
            L.append(f"- {s}")
        L.append("")

    if form_findings_active:
        L.append("## フォームの検査")
        L.append("")
        for fname, findings in form_findings_active.items():
            L.append(f"### {fname}")
            for s in findings:
                L.append(f"- {s}")
            L.append("")

    if n_ack:
        L.append("## 確認済み（意図的・件数から除外）")
        L.append("")
        L.append("目で見て意図的なデザイン等と判断し、`checkup --ack-all` で確認済み登録した"
                 "所見です。事実としては残しつつ、所見サマリ・総合判定からは除外しています。"
                 "判断を見直す場合は `checkup --unack \"文字列\"` で確認済みから外せます。")
        L.append("")
        for _, s in code_findings_acked:
            L.append(f"- {s}")
        for fname, findings in form_findings_acked.items():
            L.append(f"### {fname}")
            for s in findings:
                L.append(f"- {s}")
        L.append("")

    if sheet_h['error_cells']:
        L.append("## シートの検査（参考）")
        L.append("")
        for s, n in sheet_h['error_cells']:
            L.append(f"- {s}: エラーセル {n}個（#REF!/#NAME?/#VALUE! 等）")
        L.append("")

    if calls['orphans']:
        L.append("## どこからも呼ばれていないマクロ（参考）")
        L.append("")
        L.append("メニューやショートカットからの直実行用かもしれません（機械には判断できません）。")
        L.append("")
        for mod, name in calls['orphans'][:40]:
            sc = inv['shortcuts'].get(name)
            L.append(f"- [{mod}] {name}" + (f"（{sc}）" if sc else ""))
        if len(calls['orphans']) > 40:
            L.append(f"- … 他 {len(calls['orphans']) - 40}件")
        L.append("")

    if extra['auto_exec']:
        L.append("## 自動実行される処理（参考）")
        L.append("")
        L.append("ブックを開く・保存する・セルを変更する等で自動的に動くマクロです。"
                 "眠っていたブックを起こす前の問診に。")
        L.append("")
        for mod, name, lns in extra['auto_exec']:
            L.append(f"- [{mod}] {name}" + (f"（{lns}行）" if lns else ""))
        L.append("")

    if calls['onaction']:
        L.append("## ボタン・図形から実行されるマクロ（参考）")
        L.append("")
        L.append("シート上のボタン/図形に登録（OnAction）されているマクロです。"
                 "これらは「どこからも呼ばれていないマクロ」には数えていません。")
        L.append("")
        for name, places in sorted(calls['onaction'].items()):
            more = f" 他{len(places) - 5}箇所" if len(places) > 5 else ""
            L.append(f"- {name} ← {', '.join(places[:5])}{more}")
        L.append("")

    if calls['dynamic_runs']:
        L.append("## 動的な Application.Run（参考）")
        L.append("")
        L.append("呼び先が実行時に変数で決まる Run です（メニュー機構などの正常な作り）。"
                 "静的検査では実在確認ができないため、事実として所在だけ記します。")
        L.append("")
        for mod, proc, ln, text in calls['dynamic_runs'][:15]:
            L.append(f"- [{mod}] {proc}:{ln} `{text}`")
        if len(calls['dynamic_runs']) > 15:
            L.append(f"- … 他 {len(calls['dynamic_runs']) - 15}箇所")
        L.append("")

    if extra['hardcoded_paths']:
        L.append("## ハードコードされたパス（参考）")
        L.append("")
        L.append("フォルダ構成が変わると動かなくなる箇所の候補です（古いブックの復活時に特に確認）。")
        L.append("")
        for mod, proc, ln, path in extra['hardcoded_paths'][:20]:
            L.append(f"- [{mod}] {proc}:{ln} `{path}`")
        if len(extra['hardcoded_paths']) > 20:
            L.append(f"- … 他 {len(extra['hardcoded_paths']) - 20}箇所")
        L.append("")

    if extra['destructive']:
        L.append("## 破壊的な操作の所在（参考・問診）")
        L.append("")
        L.append("ファイル削除・シート削除などを行う箇所です（悪ではありません。"
                 "眠っていたブックを起こす前に「どこで何を消すか」を知っておくための一覧）。")
        L.append("")
        for mod, proc, ln, label, snippet in extra['destructive'][:20]:
            L.append(f"- [{mod}] {proc}:{ln} {label} `{snippet}`")
        if len(extra['destructive']) > 20:
            L.append(f"- … 他 {len(extra['destructive']) - 20}箇所")
        L.append("")

    if (extra['no_option_explicit'] or extra['error_resume'] or extra['long_procs']
            or extra['no_restore']):
        L.append("## その他の参考情報")
        L.append("")
        for mod, proc, desc in extra['no_restore']:
            L.append(f"- {desc}: [{mod}] {proc}")
        if extra['no_option_explicit']:
            L.append(f"- Option Explicit なしのモジュール: "
                     f"{', '.join(extra['no_option_explicit'])}")
        for mod, proc, ln in extra['error_resume'][:15]:
            L.append(f"- On Error Resume Next: [{mod}] {proc}:{ln}")
        if len(extra['error_resume']) > 15:
            L.append(f"- … On Error Resume Next 他 {len(extra['error_resume']) - 15}箇所")
        for mod, name, lns in extra['long_procs'][:10]:
            L.append(f"- 150行超のプロシージャ: [{mod}] {name}（{lns}行）")
        if len(extra['long_procs']) > 10:
            L.append(f"- … 150行超 他 {len(extra['long_procs']) - 10}件")
        L.append("")

    L.append("## 次の一手")
    L.append("")
    L.append("- 構成の全貌: `py vba_manager.py docs --preview 3`")
    L.append("- 呼び出し関係の図: `py vba_manager.py call-graph --mermaid`")
    L.append("- 修正前の影響確認: `py vba_manager.py impact <マクロ名>`（呼び元・呼び先を間接まで一覧）")
    L.append("- 経過観察: `py vba_manager.py 健康診断 --history`（診断履歴の一覧表）")
    if form_findings_active:
        L.append("- フォームの修正: `py form_tool.py`（move / align / tab-order …）")
    if unresolved:
        L.append("- 呼び先の検索: `py vba_manager.py grep \"<マクロ名>\"`")
    if total_findings and not vba_error:
        L.append("- 意図的な所見の除外: `py vba_manager.py checkup --ack-all`"
                 "（今回の所見を確認済みとして次回以降の件数・判定から除外）")
    L.append("")

    return {'name': inv['name'], 'rating': rating, 'total': total_findings,
            'counts': {'unresolved': len(unresolved), 'code': len(code_findings_active),
                       'form': n_form,
                       'broken': len(refs['broken']) + len(dead_links) + len(ref_names)},
            'ack_new': len(ack_new), 'ack_total': n_ack, 'ack_pruned': n_ack_pruned,
            'diff': diff, 'vba_error': vba_error, 'lines': L, 'json': json_payload}


def cmd_checkup(args):
    """ブックの健康診断レポート: checkup(健康診断) [excel_file] [--out f.md]

    構成（docs）＋呼び出し関係（call-graph）＋壊れた参照/リンク＋シート検査＋
    フォーム検査（lint）＋コードの機械検査＋バックアップ状況を 1枚の Markdown に
    束ね、総合判定（A/B/C）を付ける。所見は事実の列挙のみ（要不要の判断はしない）。
    診断のたびにブック別の履歴（_checkup_history/）へ結果を残し、
    次回の診断で「前回との比較（定期健診）」を自動で出す。
    --history=経過観察（履歴の一覧表・診断はしない）、--all=開いている全ブックを一括診断。
    --ack-all=今回の所見を確認済み（意図的）として登録し、以降の件数・判定から除外。
    --show-ack=確認済み一覧を表示、--unack 文字列=部分一致するものを確認済みから外す。
    """
    # --- 確認済み一覧の表示（診断はしない） ---
    if getattr(args, 'show_ack', False):
        if args.posargs:
            name = os.path.basename(args.posargs[0])
        else:
            _, wb = get_workbook(None, readonly=True)
            name = wb.Name
        ack = _load_checkup_ack(name)
        if not ack:
            print(f"{name} に確認済み（意図的）の所見はありません")
            return True
        print(f"===== {name} の確認済み（意図的・件数から除外）所見 {len(ack)}件 =====")
        for k in sorted(ack):
            print(f"  - {k}")
        return True

    # --- 確認済みの取り消し（部分一致で解除） ---
    if getattr(args, 'unack', None):
        if args.posargs:
            name = os.path.basename(args.posargs[0])
        else:
            _, wb = get_workbook(None, readonly=True)
            name = wb.Name
        ack = _load_checkup_ack(name)
        pat = args.unack
        matched = {k for k in ack if pat in k}
        if not matched:
            print(f"'{pat}' に一致する確認済み所見が見つかりません")
            return False
        _save_checkup_ack(name, ack - matched)
        print(f"{len(matched)}件を確認済みから外しました:")
        for k in sorted(matched):
            print(f"  - {k}")
        return True

    # --- 経過観察モード: 診断せず履歴を表で表示 ---
    if getattr(args, 'history', False):
        if args.posargs:
            name = os.path.basename(args.posargs[0])
        else:
            _, wb = get_workbook(None, readonly=True)
            name = wb.Name
        hist = _load_checkup_history(name)
        if not hist:
            print(f"{name} の診断履歴はまだありません（checkup を実行すると記録されます）")
            return False
        print(f"===== {name} の経過観察（診断履歴 {len(hist)}回） =====")
        print(f"{'日時':<18}{'判定':<12}{'所見':>4} {'未解決':>4} {'フォーム':>4} "
              f"{'マクロ':>4} {'行数':>7}  メモ")
        prev = None
        for s in hist:
            c = s.get('counts', {})
            tf = s.get('total_findings')
            if tf is None:      # 旧形式のスナップショット（判定・合計なし）
                tf = c.get('unresolved', 0) + c.get('code', 0) + c.get('form', 0)
            note = s.get('note', '')
            print(f"{s.get('time', ''):<18}{s.get('rating', '-'):<12}{tf:>4} "
                  f"{c.get('unresolved', 0):>6} {c.get('form', 0):>7} "
                  f"{s.get('total_procs', 0):>6} {s.get('total_lines', 0):>7}"
                  + (f"  {note}" if note else ""))
            if getattr(args, 'detail', False) and prev is not None:
                d = _checkup_diff(prev, s)
                if d and d['changed']:
                    for x in d['new']:
                        print(f"        ＋ {x}")
                    for x in d['resolved']:
                        print(f"        － {x}")
                    for x in d['procs_added']:
                        print(f"        ＋ マクロ追加 {x}")
                    for x in d['procs_removed']:
                        print(f"        － マクロ削除 {x}")
                    if d['lines_delta']:
                        sign = '+' if d['lines_delta'] > 0 else ''
                        print(f"        行数 {sign}{d['lines_delta']}")
            prev = s
        return True

    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, readonly=True)   # 健診モード（診断は読むだけ）
    if getattr(args, 'all_books', False):
        books = list(xl.Workbooks)
    else:
        books = [wb]

    results = []
    for b in books:
        try:
            results.append(_checkup_one(xl, b, note=getattr(args, 'note', None),
                                        ack_all=getattr(args, 'ack_all', False)))
        except Exception as e:
            print(f"警告: {getattr(b, 'Name', '?')} の診断に失敗しました: {e}")
    if not results:
        print("エラー: 診断できたブックがありません")
        return False

    if getattr(args, 'json', False):
        import json
        if len(results) == 1:
            payload = results[0]['json']
        else:
            payload = {"success": all(r['json']['success'] for r in results),
                       "books": [r['json'] for r in results]}
        print(json.dumps(payload, ensure_ascii=False), file=sys.stdout)
        if getattr(args, 'strict', False):
            return all(r['total'] == 0 for r in results)
        return True

    if len(results) > 1:
        lines = [f"# 開いている全ブックの健康診断（{len(results)}冊）", "",
                 f"- 診断日時: {time.strftime('%Y-%m-%d %H:%M')}（vba_manager checkup --all）",
                 "", "| ブック | 総合判定 | 所見 |", "|---|---|---|"]
        for r in results:
            lines.append(f"| {r['name']} | {r['rating']} | {r['total']}件 |")
        lines.append("")
        for r in results:
            lines += ["---", ""] + r['lines'] + [""]
    else:
        lines = results[0]['lines']

    out_path = getattr(args, 'out_opt', None)
    out_path = os.path.abspath(out_path) if out_path else os.path.join(SCRIPT_DIR, "_last_checkup.md")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print(f"健康診断レポートを生成: {out_path}")
    for r in results:
        c = r['counts']
        print(f"[{r['name']}] 総合判定 {r['rating']} / 所見 {r['total']}件"
              f"（未解決Call {c['unresolved']} / 参照・リンク切れ {c['broken']} / "
              f"コード {c['code']} / フォーム {c['form']}）")
        if r.get('ack_new'):
            print(f"  今回 {r['ack_new']}件を確認済み（意図的）として新規登録"
                  f"（確認済み合計 {r['ack_total']}件）")
        if r.get('ack_pruned'):
            print(f"  確認済みのうち {r['ack_pruned']}件は今回の所見に存在しないため自動整理しました"
                  "（lintロジック変更で二度と出ないキー等）")
        d = r['diff']
        if r.get('vba_error'):
            print("  VBA未検査（プロジェクト保護またはVBOM未信頼）＝シート側のみの縮退診断・履歴記録なし")
        elif d is None:
            print("  前回比: 初回診断（履歴を記録しました。次回から差分が出ます）")
        elif not d['changed']:
            print(f"  前回比: 変化なし（前回 {d['prev_time']}）")
        else:
            print(f"  前回比: 新規所見 {len(d['new'])} / 解消 {len(d['resolved'])} / "
                  f"マクロ +{len(d['procs_added'])}−{len(d['procs_removed'])}"
                  f"（前回 {d['prev_time']}）")
    # 所見があっても「診断の完了」は成功（終了コード0）。所見の有無で合否を
    # 判定したい自動化（CI/batch のゲート）だけ --strict で従来挙動にする
    if getattr(args, 'strict', False):
        return all(r['total'] == 0 for r in results)
    return True


def cmd_impact(args):
    """マクロ修正前の影響範囲予告: impact(影響範囲) [excel_file] <マクロ名>

    「このマクロに手を入れると、どこまで波及するか」を修正前に一覧する。
    - 呼び元（上流・間接含む）＝動作を変えたとき影響が及ぶ先
    - 呼び先（下流・間接含む）＝このマクロが依存している部品
    - 入口（ショートカット/自動実行イベント）も注記する
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: impact [excel_file] <マクロ名> [--json]")
        return False
    focus = rest[0]
    xl, wb = get_workbook(target_file, readonly=True)   # 健診モード（読むだけ）
    inv = _inventory_or_explain(xl, wb)
    if inv is None:
        return False
    res = _analyze_calls(inv)
    known, edges = res['known'], res['edges']
    hit = known.get(focus.lower())
    if not hit:
        print(f"エラー: マクロ '{focus}' が見つかりません")
        _suggest_similar(focus, [v[0] for v in known.values()])
        return False
    root = hit[0]
    mod_of = {v[0]: v[1] for v in known.values()}
    auto_names = {name for _, name, _ in _extra_code_scans(inv)['auto_exec']}
    oa = res.get('onaction', {})      # マクロ名 → ["シート/図形", ...]

    # 名前レベルの正方向・逆方向グラフ（call-graph と同じ粒度）
    fwd, rev = {}, {}
    for (mod, proc), callees in edges.items():
        fwd.setdefault(proc, set()).update(callees)
        for c in callees:
            rev.setdefault(c, set()).add(proc)

    def reach(graph, start):
        seen, stack = set(), [start]
        while stack:
            for nxt in graph.get(stack.pop(), ()):
                if nxt not in seen and nxt != start:
                    seen.add(nxt)
                    stack.append(nxt)
        return seen

    upstream, downstream = reach(rev, root), reach(fwd, root)

    def entry_note(name):
        notes = []
        sc = inv['shortcuts'].get(name)
        if sc:
            notes.append(sc)
        if name in auto_names:
            notes.append("自動実行イベント")
        if name in oa:
            notes.append("ボタン: " + ", ".join(oa[name][:3])
                         + (f" 他{len(oa[name]) - 3}" if len(oa[name]) > 3 else ""))
        return f"（{'/'.join(notes)}）" if notes else ""

    if getattr(args, 'json', False):
        import json
        print(json.dumps({
            "success": True, "file": inv['name'], "macro": root,
            "upstream": sorted(upstream), "downstream": sorted(downstream),
            "entries": {n: (inv['shortcuts'].get(n) or
                            ("ボタン: " + ", ".join(oa[n]) if n in oa else "自動実行イベント"))
                        for n in sorted(upstream | {root})
                        if n in inv['shortcuts'] or n in auto_names or n in oa},
        }, ensure_ascii=False), file=sys.stdout)
        return True

    def label(name):
        m = mod_of.get(name)
        return (f"[{m}] {name}" if m else name) + entry_note(name)

    print(f"===== 影響範囲の予告: {label(root)} =====")

    print(f"\n■ 呼び元（このマクロを直すと影響が及ぶ先・間接含む {len(upstream)}件）")
    if upstream:
        def walk_up(name, depth, seen):
            for caller in sorted(rev.get(name, ())):
                loop = "（循環）" if caller in seen else ""
                print("  " * depth + f"└ {label(caller)}{loop}")
                if not loop and depth < 6:
                    walk_up(caller, depth + 1, seen | {caller})
        walk_up(root, 1, {root})
    else:
        print("  (なし) — メニュー/ショートカット/ボタン/イベント直実行の可能性があります")

    print(f"\n■ 呼び先（このマクロが依存している部品・間接含む {len(downstream)}件）")
    if downstream:
        def walk_down(name, depth, seen):
            for callee in sorted(fwd.get(name, ())):
                loop = "（循環）" if callee in seen else ""
                print("  " * depth + f"└ {label(callee)}{loop}")
                if not loop and depth < 6:
                    walk_down(callee, depth + 1, seen | {callee})
        walk_down(root, 1, {root})
    else:
        print("  (なし) — 単体で完結しています")

    entries = [n for n in sorted(upstream | {root})
               if n in inv['shortcuts'] or n in auto_names or n in oa]
    if entries:
        print(f"\n■ 入口（ショートカット/ボタン/自動実行から届く経路）")
        for n in entries:
            print(f"  {label(n)}")
    return True


def cmd_grep(args):
    """全モジュール横断のVBAコード検索: grep [excel_file] <検索文字列>

    「どのマクロが ActiveSheet を使っているか」等を1回のCOM接続で調べる。
    出力: [モジュール] プロシージャ名:行番号: 該当行
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: grep [excel_file] <検索文字列> [--regex] [-i] [--module 名] [--max N] [--json]")
        return False
    needle = rest[0]
    if _reject_extra_args(rest, 1, '検索文字列は1つ。スペースを含むならクォートで囲む'):
        return False
    flags = re.IGNORECASE if getattr(args, 'ignore_case', False) else 0
    if getattr(args, 'regex', False):
        try:
            pat = re.compile(needle, flags)
        except re.error as e:
            print(f"エラー: 正規表現が不正です: {e}")
            return False
    else:
        pat = re.compile(re.escape(needle), flags)
    mod_filter = getattr(args, 'module_opt', None)
    max_hits = getattr(args, 'max_hits', None) or 200

    xl, wb = get_workbook(target_file)
    hits = []
    total = 0
    for comp in wb.VBProject.VBComponents:
        if mod_filter and comp.Name.lower() != mod_filter.lower():
            continue
        cm = comp.CodeModule
        n = cm.CountOfLines
        if n == 0:
            continue
        code = cm.Lines(1, n)
        for i, line in enumerate(code.split('\r\n'), 1):
            if pat.search(line):
                total += 1
                if len(hits) < max_hits:
                    try:
                        proc = cm.ProcOfLine(i, 0) or ''
                        # dynamic Dispatch は out引数付きメソッドをタプルで返すことがある
                        if isinstance(proc, tuple):
                            proc = proc[0] or ''
                    except Exception:
                        proc = ''
                    hits.append({'module': comp.Name, 'proc': proc,
                                 'line': i, 'text': line.rstrip()})

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": True, "file": wb.Name, "pattern": needle,
                          "total": total, "hits": hits}, ensure_ascii=False), file=sys.stdout)
        return True
    if total == 0:
        print(f"'{needle}' は見つかりませんでした。")
        return True
    for h in hits:
        proc_part = f" {h['proc']}" if h['proc'] else ""
        print(f"[{h['module']}]{proc_part}:{h['line']}: {h['text'].strip()}")
    if total > len(hits):
        print(f"…他 {total - len(hits)}件（--max で上限変更可）")
    print(f"--- {total}件 ヒット ---")
    return True


def cmd_code_replace(args):
    """全マクロ横断の一括置換: code-replace <検索> <置換>

    grep の対。差分プレビュー → 確認 → バックアップ → 変更行だけ ReplaceLine。
    行単位の置換のみ（複数行にまたがるパターンは対象外）。
    ReplaceLine 方式なので Attribute 行（ショートカット定義）は壊れない。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: code-replace [excel_file] <検索> <置換> [--regex] [--module 名] [-y]")
        return False
    needle, repl = rest[0], rest[1]
    if _reject_extra_args(rest, 2, 'スペースを含む場合はクォートで囲んでください'):
        return False
    use_regex = getattr(args, 'regex', False)
    if use_regex:
        try:
            pat = re.compile(needle)
        except re.error as e:
            print(f"エラー: 正規表現が不正です: {e}")
            return False
    else:
        pat = re.compile(re.escape(needle))
        repl_escaped = repl.replace('\\', '\\\\')   # 置換文字列の \ を文字通りに
    mod_filter = getattr(args, 'module_opt', None)

    xl, wb = get_workbook(target_file)

    # 変更計画の作成（この段階では何も書き換えない）
    plans = []       # (comp, [(行番号, 旧行, 新行), ...])
    total_lines = 0
    for comp in wb.VBProject.VBComponents:
        if mod_filter and comp.Name.lower() != mod_filter.lower():
            continue
        cm = comp.CodeModule
        n = cm.CountOfLines
        if n == 0:
            continue
        changes = []
        for i, line in enumerate(cm.Lines(1, n).split('\r\n'), 1):
            if not pat.search(line):
                continue
            try:
                new_line = pat.sub(repl if use_regex else repl_escaped, line)
            except re.error as e:
                # 置換文字列側の不正（存在しないグループ参照 \1 等）。計画段階なので無傷
                print(f"エラー: 置換文字列が不正です: {e}")
                return False
            if new_line != line:
                # ReplaceLine は行単位。置換結果に改行が入ると1行が複数行になり、
                # 同一モジュール内の後続の行番号が全部ずれて無関係な行を上書きする
                if '\r' in new_line or '\n' in new_line:
                    print("エラー: 置換結果に改行が含まれるため中止しました。")
                    print(f"  [{comp.Name}] {i}行目: {new_line.splitlines()[0]} …")
                    print("  （code-replace は行単位置換です。複数行への展開は replace-procedure を使ってください）")
                    return False
                changes.append((i, line, new_line))
        if changes:
            plans.append((comp, changes))
            total_lines += len(changes)

    if not plans:
        print(f"'{needle}' にマッチする行はありません（置換なし）")
        return True

    # 差分プレビュー
    print(f"--- 置換プレビュー: {len(plans)}モジュール / {total_lines}行 ---")
    for comp, changes in plans:
        print(f"[{comp.Name}] {len(changes)}行:")
        for i, old, new in changes[:20]:
            print(f"  {i}: - {old.strip()}")
            print(f"  {i}: + {new.strip()}")
        if len(changes) > 20:
            print(f"  … 他 {len(changes) - 20}行")
    print("-" * 40)

    if not getattr(args, 'yes', False):
        try:
            ans = input(f"{total_lines}行を置換しますか？ (y/N): ")
        except EOFError:
            # パイプ/MCP 等の非対話環境。トレースバックでなく正常なキャンセルにする
            print("非対話環境のため確認できません。-y を付けて実行してください。")
            return False
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False

    if make_backup(wb.FullName, "code_replace") is None and not getattr(args, 'force', False):
        print("エラー: バックアップが取れないため中止しました（--force で強行可）。")
        return False
    for comp, _ in plans:
        make_module_backup(wb, comp.Name)

    # 変更行だけを書き換える
    for comp, changes in plans:
        cm = comp.CodeModule
        for i, _, new_line in changes:
            cm.ReplaceLine(i, new_line)
        print(f"置換: [{comp.Name}] {len(changes)}行")
    wb.Save()
    print(f"完了: {len(plans)}モジュール / {total_lines}行 を置換して保存しました")
    return True


def _start_dialog_watcher(xl, mode=None):
    """マクロ発火中に出る MsgBox/InputBox(#32770) を検出して自動解除する監視スレッド。

    xl.Application.Run（や書き込みで走るイベントマクロ）は MsgBox が出ると閉じるまで
    ブロックし、別スレッドから閉じない限りコマンドが無言でハングする（2026-07-11 実害:
    write-range→Worksheet_Change→エラーMsgBox で操作が固まり、人が手で閉じるまで戻らず、
    改修が異常に長引いた）。そこで Excel が所有するモーダルダイアログにだけ WM_COMMAND を
    送って解除する。
    mode=None（既定・安全解除）: キャンセル→唯一ボタンの順で「閉じられるボタン」を押す。
      破壊的操作を確定させない方向に倒し、検出した事実と本文を .count / .last_text で残す。
    mode 明示: ok/enter->OK, cancel->キャンセル, yes->はい, no->いいえ。
    返り値の .stop() で終了。.count=解除した回数 / .last_text=最後に見た本文。
    """
    import threading
    try:
        import win32gui
        import win32process
        import win32con
    except Exception as ex:
        print(f"[WARNING] ダイアログ監視を開始できません（win32 不足）: {ex}", file=sys.stderr)

        class _Noop:
            def stop(self):
                pass
        return _Noop()

    excel_pid = None
    try:
        excel_pid = win32process.GetWindowThreadProcessId(int(xl.Hwnd))[1]
    except Exception:
        pass
    if excel_pid is None:
        # PID を特定できないまま監視すると、フィルタが外れて画面上の
        # 全アプリの #32770 ダイアログにボタンを送ってしまう。
        # フェイルセーフは「撃たない」側に倒す（自動応答は諦めて警告）
        print("[WARNING] Excel の PID を特定できないため、ダイアログ自動応答を無効化します"
              "（ダイアログが出た場合は手動で閉じてください）", file=sys.stderr)

        class _Noop:
            def stop(self):
                pass
        return _Noop()

    mode_l = (mode or 'safe').lower()
    # 標準ボタンID（Win32 DialogBox の既定値）。決め打ちの「望みのID」。
    _STD_ID = {'ok': 1, 'enter': 1, 'cancel': 2, 'yes': 6, 'no': 7,
               'abort': 3, 'retry': 4, 'ignore': 5, 'safe': 2}
    # テキストで拾う第2の網（OK専用MsgBoxのID化け対策・日英両対応）
    _TEXT_HINTS = {
        'ok':     ('ok', 'はい', '確定', '了解'),
        'enter':  ('ok', 'はい', '確定', '了解'),
        'cancel': ('cancel', 'キャンセル', '中止'),
        'yes':    ('yes', 'はい'),
        'no':     ('no', 'いいえ'),
    }
    want_id = _STD_ID.get(mode_l, 1)
    hints = _TEXT_HINTS.get(mode_l, ())

    def _dialog_buttons(hwnd):
        """ダイアログ内の Button コントロールを [(ctrl_id, text)] で返す。"""
        found = []

        def _child(ch, _):
            try:
                if win32gui.GetClassName(ch) == 'Button':
                    cid = win32gui.GetDlgCtrlID(ch)
                    txt = win32gui.GetWindowText(ch).replace('&', '').strip()
                    found.append((cid, txt))
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, _child, None)
        except Exception:
            pass
        return found

    def _dialog_text(hwnd):
        """ダイアログ本文（Static コントロールの文字）を採取して報告に使う。"""
        parts = []

        def _child(ch, _):
            try:
                if win32gui.GetClassName(ch) == 'Static':
                    t = win32gui.GetWindowText(ch).strip()
                    if t and t not in parts:
                        parts.append(t)
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, _child, None)
        except Exception:
            pass
        return ' / '.join(parts)

    def _resolve_button_id(hwnd):
        """このダイアログで実際に押すべきボタンIDを、実在ボタンから決める。

        OK のみの MsgBox は OK ボタンの ID が 2(IDCANCEL) になる Windows の仕様があり、
        標準IDを決め打ちで送ると閉じない（フェイブルが実弾で特定）。実在ボタンの
        ID とテキストを見て決めることで、OK専用MsgBox・Yes/No・InputBox すべてに効かせる。
        """
        btns = _dialog_buttons(hwnd)
        if not btns:
            return want_id                       # 取れなければ従来の決め打ち
        ids = [cid for cid, _ in btns]
        # 安全解除モード（既定）: 破壊確定を避けつつ「閉じられるボタン」を必ず1つ選ぶ
        if mode_l == 'safe':
            if 2 in ids:                         # キャンセル(IDCANCEL) があれば最優先
                return 2
            for cid, txt in btns:                # 文字でキャンセル/いいえ系
                tl = txt.lower()
                if any(h in tl for h in ('cancel', 'キャンセル', '中止', 'いいえ', 'no')):
                    return cid
            return btns[0][0]                    # OK専用等はその1つで閉じる
        # 1) 望む標準IDが実在すればそれ（通常の OK+キャンセル・Yes/No 等）
        if want_id in ids:
            return want_id
        # 2) テキストで一致するボタン（IDが化けていても文字で拾う）
        for cid, txt in btns:
            tl = txt.lower()
            if any(h in tl for h in hints):
                return cid
        # 3) OK系でボタンが1つだけ＝OK専用MsgBox（ID2化け）→ そのボタンを押す
        if mode_l in ('ok', 'enter') and len(btns) == 1:
            return btns[0][0]
        # 4) それでも決まらなければ決め打ちに戻す
        return want_id

    stop_evt = threading.Event()
    state = {'count': 0, 'last': ''}       # 解除した回数と最後に見た本文
    handled = set()                        # 同じダイアログを二重に数えない

    def _loop():
        while not stop_evt.is_set():
            targets = []

            def _cb(hwnd, _unused):
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return
                    if win32gui.GetClassName(hwnd) != '#32770':
                        return
                    if excel_pid is not None:
                        _, pid = win32process.GetWindowThreadProcessId(hwnd)
                        if pid != excel_pid:
                            return
                    targets.append(hwnd)
                except Exception:
                    pass

            try:
                win32gui.EnumWindows(_cb, None)
            except Exception:
                pass
            handled.intersection_update(targets)   # 消えた hwnd は忘れる
            for hwnd in targets:
                if hwnd in handled:                # 既に解除送信済みなら二重送信しない
                    continue
                try:
                    txt = _dialog_text(hwnd)
                    cid = _resolve_button_id(hwnd)
                    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, cid, 0)
                    handled.add(hwnd)
                    state['count'] += 1
                    if txt:
                        state['last'] = txt
                except Exception:
                    pass
            stop_evt.wait(0.15)

    th = threading.Thread(target=_loop, daemon=True)
    th.start()

    class _Watcher:
        def stop(self):
            stop_evt.set()

        @property
        def count(self):
            return state['count']

        @property
        def last_text(self):
            return state['last']

    return _Watcher()


def _dialog_watcher_note(watcher, mode):
    """ダイアログを自動解除したときの人向けの注記（無ければ空文字）。

    「黙って握りつぶした」にしないための報告。何が出て、どう閉じたかを一行で残す。
    """
    if watcher is None or not getattr(watcher, 'count', 0):
        return ""
    how = "指定ボタンで応答" if mode else "安全側（キャンセル優先）で自動解除"
    body = watcher.last_text or "(本文なし)"
    n = watcher.count
    return (f"⚠ 実行中にダイアログを{n}件検出し、{how}しました。"
            f"マクロがメッセージを出しています → 内容: {body}")


def cmd_run_macro(args):
    """Excelマクロを実行する"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: run-macro [excel_file] <macro_name>", file=sys.stderr)
        return False

    macro_name = rest[0]

    # 実行前にアドインや個人用マクロをロードする
    xl, wb = get_workbook(target_file, load_addins=True)

    # 警告を非表示にする
    try:
        xl.DisplayAlerts = False
    except Exception:
        pass

    full_macro_path = macro_name

    # マクロ名に "!" が含まれていない場合、どのブックにあるか検索する
    if "!" not in macro_name:
        found_wb = None
        # VBE.VBProjectsからマクロを検索
        try:
            pattern = re.compile(
                r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
                r'(?:Sub|Function)\s+' + re.escape(macro_name) + r'\b',
                re.IGNORECASE | re.MULTILINE
            )
            for p in xl.VBE.VBProjects:
                for comp in p.VBComponents:
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0:
                        code = cm.Lines(1, cm.CountOfLines)
                        if pattern.search(code):
                            try:
                                if p.Filename:
                                    found_wb = os.path.basename(p.Filename)
                            except Exception:
                                # 未保存ブックは Filename が取れない。プロジェクト名
                                # ("VBAProject") で修飾しても Run は解決しないため、
                                # 修飾なしの直接実行に任せる
                                pass
                            break
                if found_wb:
                    break
        except Exception as ex:
            print(f"[DEBUG] Failed to search macro location: {ex}", file=sys.stderr)

        if found_wb:
            # ブック名に空白等があると Application.Run はクォート必須
            # （cmd_test と同じ流儀。' 自体は Excel 規約どおり '' に重ねる）
            quoted_wb = found_wb.replace("'", "''")
            full_macro_path = f"'{quoted_wb}'!{macro_name}"
            print(f"[DEBUG] Macro found in: {found_wb}", file=sys.stderr)
        else:
            print(f"[WARNING] Macro '{macro_name}' not found in open projects. Trying direct run.", file=sys.stderr)

    # 引数（rest[1:]）。数値に見えるものは数値化して渡す（Excel MCP の run 相当）
    run_args = []
    for a in rest[1:]:
        v = _coerce_cell(a)
        run_args.append(a if v is None else v)
    if run_args:
        print(f"マクロ実行中: {full_macro_path}  引数: {run_args}", file=sys.stderr)
    else:
        print(f"マクロ実行中: {full_macro_path}", file=sys.stderr)

    # ダイアログ対策は既定で常設。--auto-dialog 明示時はそのボタンで応答、
    # 省略時は安全解除（キャンセル優先）で「無言ハング」を必ず断ち切る。
    _auto_dialog = getattr(args, 'auto_dialog', None)
    _dlg_watcher = _start_dialog_watcher(xl, _auto_dialog)

    try:
        # マクロ実行
        result = xl.Application.Run(full_macro_path, *run_args)

        _dlg_note = _dialog_watcher_note(_dlg_watcher, _auto_dialog)
        if getattr(args, 'json', False):
            import json
            out = {"success": True, "macro": full_macro_path, "result": str(result)}
            if _dlg_watcher.count:
                out["dialogs_dismissed"] = _dlg_watcher.count
                out["dialog_text"] = _dlg_watcher.last_text
            print(json.dumps(out, ensure_ascii=False), file=sys.stdout)
        else:
            print(f"マクロ実行成功。戻り値: {result}")
            if _dlg_note:
                print(_dlg_note, file=sys.stderr)
        return True
    except Exception as e:
        err_msg = str(e)
        if getattr(args, 'json', False):
            import json
            print(json.dumps({"success": False, "macro": full_macro_path, "error": err_msg}, ensure_ascii=False), file=sys.stdout)
        else:
            print(f"エラー: マクロの実行に失敗しました: {err_msg}", file=sys.stderr)
        return False
    finally:
        if _dlg_watcher is not None:
            _dlg_watcher.stop()
        # 戻さないとユーザーの Excel セッションに DisplayAlerts=False が残り、
        # 以後の手動操作で保存確認などの警告が出なくなる
        try:
            xl.DisplayAlerts = True
        except Exception:
            pass


def _com_error_text(e):
    """COM例外からVBAエラーの説明文を取り出す（取れなければ str(e)）"""
    try:
        info = getattr(e, 'excepinfo', None)
        if info and len(info) > 2 and info[2]:
            return str(info[2]).strip()
    except Exception:
        pass
    return str(e)


def cmd_test(args):
    """VBAテストランナー: test [excel_file] [絞り込み] [--module 名] [--auto-dialog ok] [--json]

    名前が「テスト」または「test」で始まる**引数なしの公開 Sub** を、
    開いたままのブックの中で1本ずつ実行し、成功/失敗を一覧で返す。
    テスト側の作法はただ一つ「失敗は Err.Raise で知らせる」
    （assert は VBA の  If 実際 <> 期待 Then Err.Raise 5, , "説明"  で書く）。
    補助モジュール不要＝テストも単体で他ブックに移植できる自立ユニット。
    実行はエラー捕捉ハーネス（一時モジュールを注入→Run→撤去）経由。
    テスト内の実行時エラーは VBA 側の On Error が受けるので、
    「Microsoft Visual Basic 実行時エラー」ダイアログは出ない。
    xlflow のテスト基盤とエラー割り込みの発想だけ移植し、ビルドせず相乗りのまま回す。
    全部成功なら終了コード0、1本でも失敗なら1（自動化ゲートにそのまま使える）。
    """
    import time as _time
    target_file, rest = parse_target_and_rest(args.posargs)
    keyword = rest[0] if rest else None
    if _reject_extra_args(rest, 1 if keyword else 0,
                          '使い方: test [excel_file] [絞り込み] [--module 名] [--auto-dialog ok] [--json]'):
        return False

    xl, wb = get_workbook(target_file)

    # 引数なしの公開 Sub だけを対象にする（Private/Friend/Function/引数つきは対象外）
    sub_pattern = re.compile(
        r'^\s*(?:Public\s+)?(?:Static\s+)?Sub\s+([^\s\(\)]+)\s*\(\s*\)',
        re.IGNORECASE | re.MULTILINE
    )

    module_filter = getattr(args, 'module', None)
    tests = []          # (モジュール名, Sub名)
    scanned_modules = 0
    try:
        for comp in wb.VBProject.VBComponents:
            if comp.Type != 1:          # 標準モジュールのみ（Runで呼べる場所）
                continue
            if module_filter and comp.Name.lower() != module_filter.lower():
                continue
            scanned_modules += 1
            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                continue
            code = cm.Lines(1, cm.CountOfLines)
            for m in sub_pattern.finditer(code):
                name = m.group(1)
                if not (name.startswith('テスト') or name.lower().startswith('test')):
                    continue
                if keyword and keyword.lower() not in name.lower():
                    continue
                tests.append((comp.Name, name))
    except Exception as ex:
        print(f"エラー: VBAプロジェクトの走査に失敗しました: {ex}", file=sys.stderr)
        return False

    if not tests:
        where = f"モジュール '{module_filter}'" if module_filter else f"標準モジュール {scanned_modules} 本"
        print(f"テストが見つかりません（{where} を走査）。")
        print("  名前が「テスト」または「test」で始まる引数なしの Sub がテストとして拾われます。")
        print("  例: Sub テスト加算()  /  失敗は  Err.Raise 5, , \"期待3 実際=\" & 結果  で知らせる。")
        return False

    print(f"テスト実行: {wb.Name}  （{len(tests)}本）")
    print("-" * 60)

    # ダイアログ自動応答（テスト中の MsgBox で止まらないように）
    _dlg_watcher = None
    _auto_dialog = getattr(args, 'auto_dialog', None)
    if _auto_dialog:
        _dlg_watcher = _start_dialog_watcher(xl, _auto_dialog)

    try:
        xl.DisplayAlerts = False
    except Exception:
        pass

    # エラー捕捉ハーネスを一時モジュールとして注入する。
    # 直接 Application.Run すると、テスト内の実行時エラー（Err.Raise 含む）は
    # COM例外にならず「Microsoft Visual Basic 実行時エラー」ダイアログで停止する
    # （実弾で確認済み・618秒ハングの正体）。さらに「VBA側で Application.Run を
    # 経由して呼ぶ」形も、Run の先のエラーが呼び元の On Error に届かず同じ結果に
    # なった（231秒・実測）。だからテストごとに**直接呼び出す**ラッパー関数を
    # 機械生成して注入する＝通常の呼び出しスタックなので On Error が確実に効き、
    # ダイアログは一切出ない。注入→Run→撤去の3ステップ、ブックは保存しない。
    _HARNESS = "VbaManagerTestHarness"
    lines = []
    for i, (mod_name, sub_name) in enumerate(tests, 1):
        lines += [
            f"Function VMT_{i}() As String",
            "    On Error GoTo eh",
            f"    {mod_name}.{sub_name}",
            f"    VMT_{i} = \"OK\"",
            "    Exit Function",
            "eh:",
            f"    VMT_{i} = \"ERR|\" & Err.Number & \"|\" & Err.Description",
            "End Function",
            "",
        ]
    harness_code = "\r\n".join(lines) + "\r\n"
    harness_comp = None
    try:
        # 前回の残骸があれば先に撤去
        for c in wb.VBProject.VBComponents:
            if c.Name == _HARNESS:
                wb.VBProject.VBComponents.Remove(c)
                break
        harness_comp = wb.VBProject.VBComponents.Add(1)
        harness_comp.Name = _HARNESS
        harness_comp.CodeModule.AddFromString(harness_code)
    except Exception as ex:
        print(f"エラー: テストハーネスの注入に失敗しました: {ex}", file=sys.stderr)
        # Add 成功後に Name 代入や AddFromString で失敗すると、既定名（Module1等）の
        # ゴミモジュールが残る。名前が _HARNESS でないと次回の残骸掃除にも拾われないため
        # ここで確実に撤去する
        if harness_comp is not None:
            try:
                wb.VBProject.VBComponents.Remove(harness_comp)
            except Exception:
                print("警告: 注入途中のモジュールを撤去できませんでした"
                      "（既定名のモジュールが残っていたら手で削除してください）", file=sys.stderr)
        if _dlg_watcher is not None:
            _dlg_watcher.stop()
        try:
            xl.DisplayAlerts = True
        except Exception:
            pass
        return False

    results = []
    # ブック名の ' は Excel 規約どおり '' に重ねる（cmd_run_macro と同じ流儀）
    quoted_wb = wb.Name.replace("'", "''")
    try:
        for i, (mod_name, sub_name) in enumerate(tests, 1):
            t0 = _time.time()
            try:
                ret = xl.Application.Run(f"'{quoted_wb}'!{_HARNESS}.VMT_{i}")
                sec = _time.time() - t0
                ret = str(ret) if ret is not None else ""
                if ret == "OK":
                    results.append({"module": mod_name, "name": sub_name,
                                    "ok": True, "seconds": round(sec, 2), "error": None})
                    print(f"○ {sub_name}  [{mod_name}]  ({sec:.2f}秒)")
                else:
                    parts = ret.split("|", 2)
                    err = (f"実行時エラー {parts[1]}: {parts[2]}"
                           if len(parts) == 3 else (ret or "不明なエラー"))
                    results.append({"module": mod_name, "name": sub_name,
                                    "ok": False, "seconds": round(sec, 2), "error": err})
                    print(f"✗ {sub_name}  [{mod_name}]  ({sec:.2f}秒)")
                    print(f"    {err}")
            except Exception as e:
                sec = _time.time() - t0
                err = _com_error_text(e)
                results.append({"module": mod_name, "name": sub_name,
                                "ok": False, "seconds": round(sec, 2), "error": err})
                print(f"✗ {sub_name}  [{mod_name}]  ({sec:.2f}秒)")
                print(f"    {err}")
    finally:
        if harness_comp is not None:
            try:
                wb.VBProject.VBComponents.Remove(harness_comp)
            except Exception:
                print("警告: テストハーネスの撤去に失敗しました（モジュール "
                      f"'{_HARNESS}' が残っていたら手で削除してください）", file=sys.stderr)
        if _dlg_watcher is not None:
            _dlg_watcher.stop()
        # 戻さないとユーザーの Excel セッションに DisplayAlerts=False が残る
        try:
            xl.DisplayAlerts = True
        except Exception:
            pass

    ok_count = sum(1 for r in results if r["ok"])
    print("-" * 60)
    print(f"結果: {ok_count}/{len(results)} 成功" + ("" if ok_count == len(results) else f"  （失敗 {len(results) - ok_count}）"))

    if getattr(args, 'json', False):
        import json
        print(json.dumps({"success": ok_count == len(results), "book": wb.Name,
                          "total": len(results), "passed": ok_count,
                          "tests": results}, ensure_ascii=False), file=sys.stdout)

    return ok_count == len(results)


# ================================================================
# 「目」コマンド (シート状態の読み取り)
# ================================================================

LAST_VIEW_FILE = os.path.join(SCRIPT_DIR, '_last_view.png')   # screenshot の出力先


def _col_letter(n):
    """列番号(1始まり)を A, B, ... Z, AA, ... に変換"""
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _resolve_range(xl, wb, spec, sheet_name=None):
    """
    範囲指定を (ws, rng) に解決する。
      'A1:D20'         → アクティブシートの範囲
      'Sheet1!A1:D20'  → シート指定の範囲
      'Sheet1'         → そのシートの UsedRange
      None / ''        → アクティブシートの UsedRange
    sheet_name（--sheet オプション）が来たら spec はアドレスのみとして扱う。
    「シート!範囲」一本槍だと、'!' を含むシート名（Excelでは合法）や
    記号入り日本語シート名のクォートで詰むための分離指定の口。
    """
    if sheet_name:
        ws = None
        for sh in wb.Sheets:
            if sh.Name == sheet_name:
                ws = sh
                break
        if ws is None:
            raise Exception(f"シート '{sheet_name}' が見つかりません")
        if not spec:
            return ws, ws.UsedRange
        return ws, ws.Range(spec)

    if not spec:
        ws = wb.ActiveSheet
        return ws, ws.UsedRange

    if '!' in spec:
        sheet_part, addr = spec.split('!', 1)
        # Excel の数式バー表記（'月次 集計'!A1）のクォートを剥がす（'' は ' に戻す）
        if len(sheet_part) >= 2 and sheet_part.startswith("'") and sheet_part.endswith("'"):
            sheet_part = sheet_part[1:-1].replace("''", "'")
        ws = wb.Sheets(sheet_part)
        if not addr:
            return ws, ws.UsedRange
        return ws, ws.Range(addr)

    # シート名そのものなら UsedRange
    for sh in wb.Sheets:
        if sh.Name == spec:
            return sh, sh.UsedRange

    # それ以外はアクティブシートのアドレスとして扱う
    ws = wb.ActiveSheet
    return ws, ws.Range(spec)


def _whole_sheet_spec(wb, spec, sheet_name=None):
    """spec がシート全域(UsedRange)に解決される形ならシート名を返す（破壊系コマンドのガード用）。

    「シート名だけ」「末尾!」「空」の spec は _resolve_range で UsedRange 全域になる。
    読み取り系では便利だが、clear/fill/sort/write 等の破壊系では
    範囲指定ミス1つで全域破壊になるため、明示指定(--whole-sheet)なしでは拒否する。
    """
    if sheet_name:
        return sheet_name if not spec else None
    if not spec:
        return wb.ActiveSheet.Name
    if '!' in spec:
        sheet_part, addr = spec.split('!', 1)
        return sheet_part if not addr else None
    for sh in wb.Sheets:
        if sh.Name == spec:
            return spec
    return None


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


def _disp_width(s):
    """全角文字を2幅として数えた表示幅"""
    w = 0
    for ch in s:
        w += 2 if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else 1
    return w


def _disp_truncate(s, width):
    """表示幅 width に収まるよう切り詰める。

    切れたことが分かるよう末尾に '…' を付ける（黙って切ると、欠けた値を
    全文と誤読して write で書き戻す事故の芽になる）。全文が要るときは
    read-range --width で広げるか --tsv で書き出す。
    """
    if _disp_width(s) <= width:
        return s
    lim = max(width - 2, 1)     # '…' は全角幅2として確保
    out = []
    w = 0
    for ch in s:
        cw = 2 if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else 1
        if w + cw > lim:
            break
        out.append(ch)
        w += cw
    return ''.join(out) + '…'


def _disp_pad(s, width, right=False):
    """表示幅基準で width までスペース埋め（right=Trueで右寄せ）"""
    pad = width - _disp_width(s)
    if pad <= 0:
        return s
    return (' ' * pad + s) if right else (s + ' ' * pad)


def _range_values_2d(rng, use_formula=False):
    """Range の値を 2次元リストへ正規化（--tsv / --json 用）"""
    raw = rng.Formula if use_formula else rng.Value
    if raw is None:
        return [['']]
    if not isinstance(raw, tuple):
        return [[raw]]
    return [list(r) if isinstance(r, tuple) else [r] for r in raw]


def _merged_areas_in_range(rng, cap=8000):
    """range 内の結合セル領域を "A3:I8" 形式のアドレス一覧で返す（重複なし）。

    戻り値は (areas, skipped_total)。
      ・結合が1つも無い範囲は rng.MergeCells が False を返すので即 ([], None)（最速パス）。
      ・cap を超えるセル数の範囲は走査せず (None, total) を返す（巨大UsedRangeの暴走防止）。
    MergeArea の矩形は Python 側の seen 集合に算術で畳み、余分な COM 呼び出しを避ける。
    神エクセルの二重構造（文字はA1・見た目はI列まで結合）を読み解くための素。
    """
    try:
        if rng.MergeCells is False:      # 範囲内に結合が皆無 → 走査不要
            return [], None
    except Exception:
        pass                              # None(混在)や例外は通常走査へ
    try:
        total = int(rng.Cells.Count)
    except Exception:
        return [], None
    if total > cap:
        return None, total
    ws = rng.Worksheet
    r0, c0 = rng.Row, rng.Column
    nr, nc = rng.Rows.Count, rng.Columns.Count
    seen = set()
    areas = []
    for i in range(nr):
        for j in range(nc):
            key = (r0 + i, c0 + j)
            if key in seen:
                continue
            try:
                cell = ws.Cells(r0 + i, c0 + j)
                if cell.MergeCells:
                    ma = cell.MergeArea
                    mr, mc = ma.Row, ma.Column
                    mrc, mcc = ma.Rows.Count, ma.Columns.Count
                    areas.append(
                        f"{_col_letter(mc)}{mr}:{_col_letter(mc + mcc - 1)}{mr + mrc - 1}")
                    for ii in range(mrc):        # 結合矩形を丸ごと走査済みにする
                        for jj in range(mcc):
                            seen.add((mr + ii, mc + jj))
                else:
                    seen.add(key)
            except Exception:
                seen.add(key)
    return areas, None


def _values_to_grid(rng, use_formula=False, max_col_width=40):
    """Range の値を、列文字＋行番号つきのテキスト格子にする

    use_formula=True のときは計算結果ではなく数式(.Formula)を表示する。
    数式のないセルは定数値がそのまま入る（write-range の .Value と同じ規約）。
    max_col_width を超える列は '…' 付きで切り詰める（--width で変更可）。
    """
    raw = rng.Formula if use_formula else rng.Value
    if raw is None:
        return "(空の範囲です)"

    # 単一セル
    if not isinstance(raw, tuple):
        a1 = f"{_col_letter(rng.Column)}{rng.Row}"
        return f"{a1}: {_cell_str(raw)}"

    # tuple-of-tuples へ正規化
    rows = []
    for row in raw:
        rows.append(list(row) if isinstance(row, tuple) else [row])
    if not rows:
        return "(空の範囲です)"

    start_row = rng.Row
    start_col = rng.Column
    ncols = max(len(r) for r in rows)

    str_rows = [[_cell_str(v) for v in row] + [''] * (ncols - len(row)) for row in rows]
    headers = [_col_letter(start_col + j) for j in range(ncols)]

    rownum_w = len(str(start_row + len(str_rows) - 1))
    col_w = []
    for j in range(ncols):
        w = _disp_width(headers[j])
        for i in range(len(str_rows)):
            w = max(w, _disp_width(str_rows[i][j]))
        col_w.append(min(w, max_col_width))

    def fmt_row(cells, label):
        parts = [_disp_pad(label, rownum_w, right=True)]
        for j, c in enumerate(cells):
            parts.append(_disp_pad(_disp_truncate(c, col_w[j]), col_w[j]))
        return ' | '.join(parts)

    out = [fmt_row(headers, ''),
           '-' * (rownum_w + sum(col_w) + 3 * ncols)]
    for i, row in enumerate(str_rows):
        out.append(fmt_row(row, str(start_row + i)))
    return '\n'.join(out)


def cmd_read_range(args):
    """シートのセル値をテキスト格子で読み取る（目・テキスト版）。

    複数範囲可（1回のCOM接続でまとめ読み）。--tsv で _last_values.tsv に
    書き出せば「読む→TSVを編集→write-range で書き戻す」の往復が
    get→_last_proc.vba→replace-procedure と同じ型になる。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    specs = rest if rest else [None]
    use_formula = getattr(args, 'formula', False)
    try:
        width = int(getattr(args, 'width', None) or 40)
    except (TypeError, ValueError):
        print("エラー: --width は数値で指定してください")
        return False
    tsv_out = getattr(args, 'tsv_out', None)
    if tsv_out is not None and len(specs) > 1:
        print("エラー: --tsv は範囲1つのときだけ使えます")
        return False

    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    blocks = [(_resolve_range(xl, wb, spec, sheet_opt)) for spec in specs]

    if getattr(args, 'json', False):
        import json
        out = []
        for ws, rng in blocks:
            rows = _range_values_2d(rng, use_formula)
            entry = {"sheet": ws.Name, "address": rng.Address,
                     "ref": f"{ws.Name}!{rng.Address}",
                     "rows": [[_cell_str(v) for v in r] for r in rows]}
            areas, skipped = _merged_areas_in_range(rng)
            if skipped:
                entry["merged_skipped_cells"] = skipped   # cap超で未走査
            elif areas:
                entry["merged"] = areas
            out.append(entry)
        print(json.dumps({"success": True, "file": wb.Name, "ranges": out},
                         ensure_ascii=False), file=sys.stdout)
        return True

    for ws, rng in blocks:
        mode = "（数式表示）" if use_formula else ""
        # 末尾の [シート名!番地] はそのまま次コマンドの range 引数に貼れる形
        print(f"シート: {ws.Name}   範囲: {rng.Address}{mode}   [{ws.Name}!{rng.Address}]")
        print("=" * 60)
        print(_values_to_grid(rng, use_formula, max_col_width=width))
        print("=" * 60)
        areas, skipped = _merged_areas_in_range(rng)
        if skipped:
            print(f"結合セル: 未走査（{skipped}セルはcap超・範囲を絞れば表示）")
        elif areas:
            shown = areas[:30]
            more = f"  …他{len(areas) - 30}件" if len(areas) > 30 else ""
            print(f"結合セル {len(areas)}件: " + ", ".join(shown) + more)

    if tsv_out is not None:
        path = _LAST_VALUES_FILE if tsv_out == '_DEFAULT_' else os.path.abspath(tsv_out)
        ws, rng = blocks[0]
        rows = _range_values_2d(rng, use_formula)
        # セル内改行(Alt+Enter)・タブは TSV の行/列区切りと衝突し、
        # そのまま write-range で書き戻すと格子がずれて無警告のデータ破壊になる。
        # 値は変えず、該当セルを名指しで警告する
        dirty = []
        for ri, r in enumerate(rows):
            for ci, v in enumerate(r):
                if isinstance(v, str) and ('\n' in v or '\r' in v or '\t' in v):
                    dirty.append(f"{_col_letter(rng.Column + ci)}{rng.Row + ri}")
        if dirty:
            shown = ", ".join(dirty[:8]) + ("" if len(dirty) <= 8 else f" …他{len(dirty) - 8}件")
            print(f"⚠ 警告: セル内に改行/タブを含むセルが {len(dirty)}件あります: {shown}")
            print("  このTSVをそのまま write-range で書き戻すと行・列がずれます。")
            print("  該当セルは手で編集するか、write-range の対象から外してください。")
        lines = ['\t'.join(_cell_str(v) for v in r) for r in rows]
        with open(path, 'w', encoding='utf-8-sig', newline='') as f:
            f.write('\n'.join(lines) + '\n')
        print(f"TSV書き出し: {path}  ({len(rows)}行 x {max(len(r) for r in rows)}列)")
        print(f"  編集後の書き戻し: py vba_manager.py write-range \"{ws.Name}!{_col_letter(rng.Column)}{rng.Row}\"")
        print(f"  （\"007\" 等の先頭ゼロを数値化させたくない場合は --raw を付ける）")
    return True


def cmd_read_selection(args):
    """ユーザーが今選択している範囲を読み取る"""
    target_file, _ = parse_target_and_rest(args.posargs)
    use_formula = getattr(args, 'formula', False)
    xl, wb = get_workbook(target_file)
    sel = xl.Selection
    if sel is None:
        print("選択範囲がありません。")
        return False
    try:
        wsname = sel.Worksheet.Name
    except Exception:
        wsname = "(不明)"
    try:
        addr = sel.Address
    except Exception:
        addr = "(範囲ではありません)"
    mode = "（数式表示）" if use_formula else ""
    print(f"シート: {wsname}   選択範囲: {addr}{mode}")
    print("=" * 60)
    try:
        print(_values_to_grid(sel, use_formula))
    except Exception as e:
        print(f"(値を読めませんでした: {e})")
    print("=" * 60)
    return True


def cmd_sheet_info(args):
    """ブックのシート構成・使用範囲を表示（見取り図）。

    --preview N で各シート使用範囲の先頭N行も格子表示（初見ブックの俯瞰が
    1コマンド1接続で済む。従来は sheet-info + シート毎の read-range で N+1 接続）。
    """
    target_file, _ = parse_target_and_rest(args.posargs)
    try:
        preview = int(getattr(args, 'preview', None) or 0)
    except (TypeError, ValueError):
        print("エラー: --preview は数値で指定してください")
        return False
    xl, wb = get_workbook(target_file)
    active = wb.ActiveSheet.Name
    print(f"ブック: {wb.Name}")
    print(f"シート数: {wb.Sheets.Count}   アクティブ: {active}")
    print("-" * 60)
    for sh in wb.Sheets:
        mark = '*' if sh.Name == active else ' '
        try:
            ur = sh.UsedRange
            dims = f"{ur.Rows.Count}行 x {ur.Columns.Count}列  ({ur.Address})"
        except Exception:
            ur = None
            dims = "(空)"
        vis = '' if sh.Visible == -1 else '  [非表示]'
        print(f"{mark} {sh.Name}: {dims}{vis}")
        if ur is not None:
            areas, skipped = _merged_areas_in_range(ur)
            if skipped:
                print(f"    結合: 未走査（{skipped}セル・大）")
            elif areas:
                shown = areas[:12]
                more = f"  …他{len(areas) - 12}件" if len(areas) > 12 else ""
                print(f"    結合 {len(areas)}件: " + ", ".join(shown) + more)
        if preview > 0 and ur is not None:
            try:
                nrows = min(preview, ur.Rows.Count)
                head = ur.Worksheet.Range(ur.Cells(1, 1), ur.Cells(nrows, ur.Columns.Count))
                grid = _values_to_grid(head)
                print('    ' + grid.replace('\n', '\n    '))
            except Exception as e:
                print(f"    (先頭行を読めませんでした: {e})")
    print("-" * 60)
    return True


def _shape_text(shp):
    """図形の表示文字を複数方式で拾う。

    Formsコントロール(ボタン等)は TextFrame.Characters().Text、
    AutoShape等は TextFrame2.TextRange.Text に文字が入る（相手で口が違う）。
    先に非空を返した方を採用。どちらも取れなければ None。
    """
    for getter in (
        lambda: shp.TextFrame.Characters().Text,
        lambda: shp.TextFrame2.TextRange.Text,
    ):
        try:
            t = getter()
            if t:
                return t
        except Exception:
            continue
    return None


def _collect_shapes(shapes, out, in_group=None):
    """Shapes を平坦なリストに集める。グループ(msoGroup=6)は中のボタンも展開して拾う。

    ファイル一覧等ではボタンがグループにまとめられており、展開しないと
    中の「抽出開始」「選択抽出」等が丸ごと落ちる（実測で判明）。
    """
    for shp in shapes:
        s = {}
        try:
            s["name"] = shp.Name
        except Exception:
            pass
        try:
            typ = int(shp.Type)
            s["type"] = typ
        except Exception:
            typ = None
        txt = _shape_text(shp)
        if txt:
            s["text"] = txt
        try:
            s["l"] = round(float(shp.Left))
            s["t"] = round(float(shp.Top))
        except Exception:
            pass
        try:
            oa = shp.OnAction
            if oa:
                s["onaction"] = oa.split('!')[-1].strip("'\" ")
                if '!' in oa:
                    # 「'ブック名'!マクロ名」のブック修飾。剥がして捨てると
                    # アドイン先の配線を wiring が「存在しない」と誤判定する
                    book = oa.rsplit('!', 1)[0].strip("'\" ")
                    if book:
                        s["onaction_book"] = book
        except Exception:
            pass
        if in_group:
            s["group"] = in_group
        out.append(s)
        if typ == 6:                       # msoGroup → 中身を1段展開
            try:
                _collect_shapes(shp.GroupItems, out, in_group=s.get("name"))
            except Exception:
                pass


def cmd_snapshot(args):
    """アクティブブック(または1シート)を意味構造JSONに畳む＝開いたままブックLMの下ごしらえ。

    セル(疎)＋結合(merged)＋図形/ボタン(text・座標・OnAction)＋テーブル(ListObject)を
    1ファイルに束ねる。晴美さんのExStruct extract を「開いてるブック相手・その場」で焼き直したもの。
    表かどうかの"推定"はしない（機械的事実だけ吐き、意味付けは読み手のAIがやる＝この子の設計思想）。
    """
    import json
    target_file, rest = parse_target_and_rest(args.posargs)
    only_sheet = rest[0] if rest else getattr(args, 'sheet_opt', None)
    out_path = os.path.abspath(getattr(args, 'out_opt', None) or _LAST_SNAPSHOT_FILE)
    try:
        max_rows = int(getattr(args, 'max_rows', None) or 5000)
    except (TypeError, ValueError):
        print("エラー: --max-rows は数値で指定してください")
        return False

    xl, wb = get_workbook(target_file)
    active = wb.ActiveSheet.Name

    if only_sheet:
        targets = [sh for sh in wb.Sheets if sh.Name == only_sheet]
        if not targets:
            print(f"シート '{only_sheet}' が見つかりません")
            return False
    else:
        targets = list(wb.Sheets)

    sheets = {}
    for sh in targets:
        info = {}
        try:
            ur = sh.UsedRange
            nr, nc = ur.Rows.Count, ur.Columns.Count
            info["dims"] = f"{nr}行 x {nc}列"
            info["used"] = ur.Address
        except Exception:
            ur, nr, nc = None, 0, 0
            info["dims"] = "(空)"
        info["visible"] = (sh.Visible == -1)

        # --- セル（疎：空セル・空行は落とす。読み込みも max_rows 行までに絞る）---
        cells = []
        if ur is not None and nr > 0:
            r0, c0 = ur.Row, ur.Column
            read_rows = min(nr, max_rows)
            try:
                sub = sh.Range(ur.Cells(1, 1), ur.Cells(read_rows, nc))
                raw = sub.Value
            except Exception:
                raw = None
            if raw is None:
                grid = []
            elif not isinstance(raw, tuple):
                grid = [(raw,)]
            else:
                grid = [r if isinstance(r, tuple) else (r,) for r in raw]
            for i, row in enumerate(grid):
                cmap = {}
                for j, v in enumerate(row):
                    if v is None or v == '':
                        continue
                    cmap[_col_letter(c0 + j)] = _cell_str(v)
                if cmap:
                    cells.append({"r": r0 + i, "c": cmap})
            if nr > max_rows:
                info["cells_truncated"] = {"read_rows": max_rows, "total_rows": nr}
        info["cells"] = cells

        # --- 結合（①の素を流用）---
        if ur is not None:
            areas, skipped = _merged_areas_in_range(ur)
            if skipped:
                info["merged_skipped_cells"] = skipped
            elif areas:
                info["merged"] = areas

        # --- 図形／ボタン（text・座標・実行マクロ。グループは1段展開）---
        shapes = []
        try:
            _collect_shapes(sh.Shapes, shapes)
        except Exception:
            pass
        if shapes:
            info["shapes"] = shapes

        # --- テーブル（正式な ListObject だけ・推定はしない）---
        tables = []
        try:
            for lo in sh.ListObjects:
                tables.append({"name": lo.Name, "address": lo.Range.Address})
        except Exception:
            pass
        if tables:
            info["tables"] = tables

        sheets[sh.Name] = info

    doc = {"success": True, "book": wb.Name, "active_sheet": active, "sheets": sheets}
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)

    # 端末にはサマリだけ（本体JSONは大きくなり得るのでファイルへ）
    print(f"スナップショット: {wb.Name}  → {out_path}")
    print("-" * 60)
    for name, info in sheets.items():
        parts = [info.get("dims", "")]
        parts.append(f"セル{len(info.get('cells', []))}行")
        if "cells_truncated" in info:
            parts.append(f"(全{info['cells_truncated']['total_rows']}行→{info['cells_truncated']['read_rows']}打切)")
        if "merged" in info:
            parts.append(f"結合{len(info['merged'])}")
        elif "merged_skipped_cells" in info:
            parts.append("結合未走査(大)")
        if "shapes" in info:
            parts.append(f"図形{len(info['shapes'])}")
        if "tables" in info:
            parts.append(f"表{len(info['tables'])}")
        mark = '*' if name == active else ' '
        print(f"{mark} {name}: " + "  ".join(p for p in parts if p))
    print("-" * 60)
    print("このJSONを read して質問すれば、開いたままブックLMになる。")
    return True


def cmd_snapshot_diff(args):
    """2つの snapshot JSON を機械的に比較: snapshot-diff <before.json> [after.json]

    checkup を前後に挟む型の「シート側」版＝マクロや手作業が実際に何を変えたかを
    セル・結合・図形・テーブル単位の事実で示す（COM 不要の純粋処理）。
    after 省略時は _last_snapshot.json（直近の snapshot）と比較する。
    評価はしない＝差分という事実だけ並べ、意味付けは読み手がやる（snapshot と同じ思想）。
    """
    import json
    if not args.posargs:
        print("使い方: snapshot-diff <before.json> [after.json]")
        print("  after 省略時は _last_snapshot.json（直近の snapshot）と比較します")
        return False
    old_path = os.path.abspath(args.posargs[0])
    new_path = (os.path.abspath(args.posargs[1]) if len(args.posargs) >= 2
                else _LAST_SNAPSHOT_FILE)
    try:
        max_show = int(getattr(args, 'max_opt', None) or 20)
    except (TypeError, ValueError):
        print("エラー: --max は数値で指定してください")
        return False

    docs = []
    for path in (old_path, new_path):
        if not os.path.exists(path):
            print(f"エラー: ファイルがありません: {path}")
            return False
        try:
            # utf-8-sig: PowerShell の Out-File 等が付ける BOM も受け入れる（BOM無しも可）
            with open(path, 'r', encoding='utf-8-sig') as f:
                d = json.load(f)
        except Exception as e:
            print(f"エラー: JSON を読めません: {path} ({e})")
            return False
        if not isinstance(d, dict) or 'sheets' not in d:
            print(f"エラー: snapshot 形式ではありません（'sheets' がない）: {path}")
            return False
        docs.append(d)
    old, new = docs

    def clip(v, n=40):
        s = str(v).replace('\r', ' ').replace('\n', ' ')
        return s if len(s) <= n else s[:n] + '…'

    def cell_map(info):
        out = {}
        for row in info.get('cells', ()):
            for col, v in row.get('c', {}).items():
                out[(row['r'], col)] = v
        return out

    def addr(key):
        return f"{key[1]}{key[0]}"

    def cell_order(key):
        return (key[0], len(key[1]), key[1])   # 行→列文字（桁→辞書順）で安定表示

    def show(label, items, fmt):
        print(f"  {label}: {len(items)}件")
        for it in items[:max_show]:
            print(f"    {fmt(it)}")
        if len(items) > max_show:
            print(f"    … 他 {len(items) - max_show}件（--max で表示数変更可）")

    print("===== スナップショット差分 =====")
    print(f"  旧: {old.get('book', '?')}  ({old_path})")
    print(f"  新: {new.get('book', '?')}  ({new_path})")
    if old.get('book') != new.get('book'):
        print("  ※ ブック名が異なります（別ブック同士の比較）")

    old_sheets, new_sheets = old['sheets'], new['sheets']
    diff_sheets = 0
    merged_note_hidden = []   # 結合未走査で比較できず、かつ他に差分がなく非表示のシート

    for name in [n for n in old_sheets if n not in new_sheets]:
        diff_sheets += 1
        print(f"\n- シート削除: {name}")
    for name in [n for n in new_sheets if n not in old_sheets]:
        diff_sheets += 1
        info = new_sheets[name]
        print(f"\n+ シート追加: {name}  ({info.get('dims', '?')})")

    for name in [n for n in old_sheets if n in new_sheets]:
        oi, ni = old_sheets[name], new_sheets[name]

        oc, nc_ = cell_map(oi), cell_map(ni)
        added = sorted([k for k in nc_ if k not in oc], key=cell_order)
        removed = sorted([k for k in oc if k not in nc_], key=cell_order)
        changed = sorted([k for k in oc if k in nc_ and oc[k] != nc_[k]],
                         key=cell_order)

        merged_unscanned = bool(oi.get('merged_skipped_cells')
                                or ni.get('merged_skipped_cells'))
        if merged_unscanned:
            # 片側でも結合未走査なら差分は出せない（空 vs 実データの疑似差分を出さない）
            m_add, m_del = [], []
        else:
            om = set(oi.get('merged') or [])
            nm = set(ni.get('merged') or [])
            m_add, m_del = sorted(nm - om), sorted(om - nm)

        def shape_map(info):
            out, dup = {}, set()
            for s in info.get('shapes', ()):
                nm2 = s.get('name', '(無名)')
                if nm2 in out:
                    dup.add(nm2)          # Excel は図形名の重複を許す＝黙って落とさず注記する
                out.setdefault(nm2, s)
            return out, dup
        os_, odup = shape_map(oi)
        ns_, ndup = shape_map(ni)
        shape_dup = odup | ndup
        s_add = sorted([n2 for n2 in ns_ if n2 not in os_])
        s_del = sorted([n2 for n2 in os_ if n2 not in ns_])
        s_chg = []
        for n2 in sorted(set(os_) & set(ns_)):
            a, b = os_[n2], ns_[n2]
            fields = []
            if a.get('text') != b.get('text'):
                fields.append(f"文字 '{clip(a.get('text'))}'→'{clip(b.get('text'))}'")
            if a.get('onaction') != b.get('onaction'):
                fields.append(f"OnAction {a.get('onaction')}→{b.get('onaction')}")
            if (a.get('l'), a.get('t')) != (b.get('l'), b.get('t')):
                fields.append(f"位置 ({a.get('l')},{a.get('t')})→({b.get('l')},{b.get('t')})")
            if fields:
                s_chg.append((n2, fields))

        def table_map(info):
            return {t['name']: t.get('address') for t in info.get('tables', ())}
        ot, nt = table_map(oi), table_map(ni)
        t_add = sorted([n2 for n2 in nt if n2 not in ot])
        t_del = sorted([n2 for n2 in ot if n2 not in nt])
        t_chg = sorted([n2 for n2 in ot if n2 in nt and ot[n2] != nt[n2]])

        has_diff = any((added, removed, changed, m_add, m_del,
                        s_add, s_del, s_chg, t_add, t_del, t_chg,
                        oi.get('dims') != ni.get('dims')))
        if not has_diff:
            if merged_unscanned:
                # シート自体を表示しない場合も、比較できなかった事実は落とさない
                merged_note_hidden.append(name)
            continue
        diff_sheets += 1
        print(f"\n* シート: {name}")
        if oi.get('dims') != ni.get('dims'):
            print(f"  使用範囲: {oi.get('dims')} → {ni.get('dims')}")
        trunc = oi.get('cells_truncated') or ni.get('cells_truncated')
        if trunc:
            print(f"  ※ セルは {trunc['read_rows']}行で打ち切られた snapshot ＝比較もその範囲まで")
        if merged_unscanned:
            print("  ※ 結合セル未走査の snapshot（範囲が大）＝結合の差分は比較していない")
        if shape_dup:
            print(f"  ※ 同名の図形が複数: {', '.join(sorted(shape_dup))}"
                  "（名前単位の比較のため2つ目以降は対象外）")
        if changed:
            show("セル変更", changed,
                 lambda k: f"{addr(k)}: '{clip(oc[k])}' → '{clip(nc_[k])}'")
        if added:
            show("セル追加", added, lambda k: f"{addr(k)}: '{clip(nc_[k])}'")
        if removed:
            show("セル削除", removed, lambda k: f"{addr(k)}: '{clip(oc[k])}'")
        if m_add:
            show("結合追加", m_add, lambda a: a)
        if m_del:
            show("結合解除", m_del, lambda a: a)
        if s_add:
            show("図形追加", s_add,
                 lambda n2: n2 + (f"「{clip(ns_[n2].get('text'))}」" if ns_[n2].get('text') else ""))
        if s_del:
            show("図形削除", s_del,
                 lambda n2: n2 + (f"「{clip(os_[n2].get('text'))}」" if os_[n2].get('text') else ""))
        if s_chg:
            show("図形変更", s_chg, lambda it: f"{it[0]}: " + " / ".join(it[1]))
        if t_add:
            show("テーブル追加", t_add, lambda n2: f"{n2} ({nt[n2]})")
        if t_del:
            show("テーブル削除", t_del, lambda n2: f"{n2} ({ot[n2]})")
        if t_chg:
            show("テーブル範囲変更", t_chg, lambda n2: f"{n2}: {ot[n2]} → {nt[n2]}")

    if merged_note_hidden:
        print(f"\n※ 結合セル未走査の snapshot のため、次のシートは結合の差分を比較できていません: "
              f"{', '.join(merged_note_hidden)}")
    print("\n" + "-" * 60)
    if diff_sheets:
        print(f"差分あり: シート{diff_sheets}枚に変更")
    else:
        print("差分なし（2つのスナップショットは一致）")
    return True


def cmd_wiring(args):
    """ボタン⇔マクロの配線図: wiring [excel_file] [--json]

    シート上の図形/ボタンに登録された OnAction を全部拾い（グループは1段展開）、
    ブック内のマクロ名簿と突き合わせる。行き先のないボタン＝壊れた配線の検出器。
    call-graph（孤立マクロ）・docs（マクロ→ボタン逆引き）と対になる「ボタン側から見た地図」。
    VBA に触れないブック（保護/VBOM未信頼）でも配線一覧だけは出す（実在確認のみ縮退）。
    別ブック修飾（'秀.xlam'!マクロ 等）の配線は名簿の外＝実在確認せず「外部ブック先」として
    別枠で表示し、壊れた配線には数えない（このブックの名簿で×を付けると誤検出になる）。
    終了コード: テキスト表示は壊れた配線ありで 1（lint 型）。--json は常に 0（broken が判定を運ぶ）。
    """
    import json
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file, readonly=True)   # 診断は読むだけ

    # マクロ名簿（Sub/Function 名 → 正式名）
    known = {}
    vba_error = None
    proc_pat = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE)
    try:
        for comp in wb.VBProject.VBComponents:
            if int(comp.Type) not in (1, 2, 3, 100):
                continue
            cm = comp.CodeModule
            n = cm.CountOfLines
            code = cm.Lines(1, n) if n else ""
            for m in proc_pat.finditer(code):
                known.setdefault(m.group(1).lower(), m.group(1))
    except Exception as e:
        vba_error = e

    rows = []
    for sh in wb.Worksheets:
        shapes = []
        try:
            _collect_shapes(sh.Shapes, shapes)
        except Exception:
            continue
        for s in shapes:
            macro = s.get('onaction')
            if not macro:
                continue
            # ブック修飾つき配線。自ブック名なら普通に照合、他ブック(アドイン等)なら
            # このブックの名簿では実在確認できない＝外部ブック先として別枠にする
            book = s.get('onaction_book')
            if book and book.lower() == wb.Name.lower():
                book = None
            resolved = None
            if book is None and vba_error is None:
                hit = known.get(macro.lower())
                if hit is None and '.' in macro:
                    # 「モジュール名.マクロ名」形式（同名マクロがあると Excel が
                    # 自動でこの形式にする）は末尾名で解決する
                    hit = known.get(macro.rsplit('.', 1)[-1].lower())
                resolved = hit is not None
            rows.append({'sheet': sh.Name, 'shape': s.get('name', '(図形)'),
                         'text': s.get('text'), 'group': s.get('group'),
                         'macro': macro, 'book': book, 'resolved': resolved})

    broken = [r for r in rows if r['resolved'] is False]
    external = [r for r in rows if r['book']]

    if getattr(args, 'json', False):
        print(json.dumps({"success": True, "book": wb.Name,
                          "vba_readable": vba_error is None,
                          "wires": rows, "broken": len(broken),
                          "external": len(external)},
                         ensure_ascii=False))
        # JSON では broken フィールドが判定を運ぶ。ここで exit 1 にすると
        # success:true の JSON に MCP が「失敗しました」を付けて食い違う
        return True

    def label(r):
        t = f"「{str(r['text'])[:30]}」" if r['text'] else ""
        g = f"（グループ {r['group']} 内）" if r['group'] else ""
        return f"{r['shape']}{t}{g}"

    print(f"===== ボタン⇔マクロ配線図: {wb.Name} =====")
    if vba_error is not None:
        print("※ VBA プロジェクトに触れないため実在確認は未実施（配線の一覧のみ）")
        print(f"   詳細: {vba_error}")
    if not rows:
        print("OnAction が登録された図形/ボタンはありません")
        return True

    cur = None
    for r in rows:
        if r['sheet'] != cur:
            cur = r['sheet']
            print(f"\nシート: {cur}")
        if r['book']:
            mark = "（外部ブック先・このブックの名簿では確認できない）"
            dest = f"{r['book']}!{r['macro']}"
        elif r['resolved'] is True:
            mark = "○"
            dest = r['macro']
        elif r['resolved'] is False:
            mark = "×（存在しない）"
            dest = r['macro']
        else:
            mark = "（未確認）"
            dest = r['macro']
        print(f"  {label(r)} → {dest}  {mark}")

    print("\n" + "-" * 60)
    if vba_error is None:
        line = (f"配線 {len(rows)}本 / 実在 {len(rows) - len(broken) - len(external)}"
                f" / 行き先なし {len(broken)}")
        if external:
            line += f" / 外部ブック先 {len(external)}（確認対象外）"
        print(line)
        if broken:
            print("\n⚠ 行き先のないボタン（OnAction 先のマクロが見つからない）:")
            for r in broken:
                print(f"  {r['sheet']} / {label(r)} → {r['macro']}")
    else:
        print(f"配線 {len(rows)}本（実在確認なし）")
    print("参考: マクロ側から見た地図は call-graph（孤立検出）/ docs（マクロ→ボタン逆引き）")
    return not broken


def cmd_screenshot(args):
    """範囲を画像(PNG)として書き出す（目・画像版）"""
    target_file, rest = parse_target_and_rest(args.posargs)
    spec = rest[0] if rest else None
    out_path = os.path.abspath(getattr(args, 'out_opt', None) or LAST_VIEW_FILE)

    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, spec)

    # 対象シートをアクティブにすると CopyPicture が安定する
    try:
        ws.Activate()
    except Exception:
        pass

    # Appearance: xlScreen=1 / xlPrinter=2,  Format: xlBitmap=2 / xlPicture(EMF)=-4147
    # ※ chart へ貼る前に cob.Activate() しないと Paste が無反応で白紙PNGになる（要・最重要）
    # ※ 成否は出力サイズではなく「貼り付け後の Shapes 数」で判定する（白紙でもファイルは生成されるため）
    attempts = [(1, 2), (1, -4147), (2, 2)]
    last_err = None
    for appearance, fmt in attempts:
        cob = None
        try:
            rng.CopyPicture(appearance, fmt)
            time.sleep(0.4)
            pythoncom.PumpWaitingMessages()

            cob = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
            cob.Activate()                      # ← これが無いと貼り付かない
            chart = cob.Chart
            time.sleep(0.3)
            chart.Paste()
            time.sleep(0.4)
            pythoncom.PumpWaitingMessages()

            pasted = chart.Shapes.Count          # 1 以上なら貼り付け成功
            if pasted >= 1:
                if os.path.exists(out_path):
                    os.remove(out_path)
                chart.Export(out_path, "PNG")
                cob.Delete()
                cob = None
                if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
                    _screenshot_cleanup(xl, ws)
                    print(f"シート: {ws.Name}   範囲: {rng.Address}")
                    print(f"画像保存: {out_path}")
                    print("（注: 一時グラフの作成を伴うため、ブックの Undo 履歴は消えています）")
                    return True
                last_err = "Export に失敗しました"
            else:
                last_err = "クリップボードから貼り付けできませんでした"
        except Exception as e:
            last_err = str(e)
        finally:
            if cob is not None:
                try:
                    cob.Delete()
                except Exception:
                    pass
        time.sleep(0.4)

    _screenshot_cleanup(xl, ws)
    print(f"エラー: スクリーンショットに失敗しました ({last_err})")
    return False


def _screenshot_cleanup(xl, ws):
    """screenshot の後始末（成功・失敗の両経路で共通）。

    ChartObject 操作で選択が動くため A1 に戻し、CopyPicture で
    クリップボードに残った画像もクリアする（copy-range と対称）。
    """
    try:
        ws.Range("A1").Select()
    except Exception:
        pass
    try:
        xl.CutCopyMode = False
    except Exception:
        pass


# ================================================================
# 「手」コマンド (シートの編集・整形・構造操作)
#   ※ アクティブ(開いたまま)のブックに COM で直接書き込む。
#      Excel MCP と違いブックを閉じる必要がない代わりに、
#      プログラム経由の変更は Excel の Undo 履歴を消す。
#   ※ 既定では保存しない。Excelで確認後に手動保存するか、
#      保存せず閉じれば変更を破棄できる（=Undo代わりの逃げ道）。
# ================================================================

_LAST_VALUES_FILE = os.path.join(SCRIPT_DIR, '_last_values.tsv')   # write-range のグリッド入力
_LAST_QUERY_FILE  = os.path.join(SCRIPT_DIR, '_last_query.m')       # powerquery add のM式入力
_LAST_DAX_FILE    = os.path.join(SCRIPT_DIR, '_last_dax.dax')       # datamodel measure add のDAX入力
_LAST_SNAPSHOT_FILE = os.path.join(SCRIPT_DIR, '_last_snapshot.json')  # snapshot の意味構造JSON出力先

# 配置・罫線の定数 (xl定数の実値)
_XL_ALIGN_H = {'left': -4131, 'center': -4108, 'right': -4152,
               'fill': 5, 'justify': -4130}
_XL_ALIGN_V = {'top': -4160, 'center': -4108, 'bottom': -4107}
_XL_BORDER_WEIGHT = {'hairline': 1, 'thin': 2, 'medium': -4138, 'thick': 4}


def _hex_to_excel_color(hexstr):
    """#RRGGBB / RRGGBB → Excel の BGR 整数"""
    s = hexstr.lstrip('#').strip()
    if len(s) != 6:
        raise ValueError(f"色は #RRGGBB 形式で指定してください: {hexstr}")
    r = int(s[0:2], 16); g = int(s[2:4], 16); b = int(s[4:6], 16)
    return r + g * 256 + b * 65536


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


def _read_tsv_grid(path, raw=False):
    """TSV(タブ区切り)をセル値の 2次元タプルに変換（行の長さは不揃いのまま返す）。

    以前は短い行を None で最大列数まで詰めていたが、None 代入は既存セルの
    クリアとして作用し「触れないつもりの右側セル」を消すため、詰め物はしない。
    矩形化の要否は書き込み側（cmd_write_range）が判断する。
    """
    with open(path, 'r', encoding='utf-8-sig') as f:
        text = f.read()
    text = text.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n')
    if text == '':
        return ()
    return tuple(tuple((c if raw else _coerce_cell(c)) for c in line.split('\t'))
                 for line in text.split('\n'))


def cmd_write_range(args):
    """セル範囲に値・数式を書き込む (read-range の対)"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: write-range [excel_file] <range> [値]")
        print("  単一値はインライン、グリッドは --tsv <file> か _last_values.tsv から読み込み")
        print("  '='始まりは数式として書き込み")
        return False
    spec = rest[0]
    inline_value = rest[1] if len(rest) >= 2 else None
    tsv_opt = getattr(args, 'tsv_opt', None)
    if _reject_extra_args(rest, 2, '使い方: write-range [excel_file] <range> [値]'):
        return False

    xl, wb = get_workbook(target_file)
    append = getattr(args, 'append', False)
    if append:
        # --append: spec は「シート名!列文字」または「列文字」。使用範囲の最終行の
        # 次の行から書く（自動の最終行判定を使うため、書き込み先番地を必ず明示表示する）
        if '!' in spec:
            sheet_part, col_part = spec.split('!', 1)
            ws = None
            for sh in wb.Worksheets:
                if sh.Name == sheet_part:
                    ws = sh
                    break
            if ws is None:
                print(f"エラー: シート '{sheet_part}' が見つかりません")
                return False
        else:
            ws, col_part = wb.ActiveSheet, spec
        if not re.fullmatch(r'[A-Za-z]{1,3}', col_part or ''):
            print("エラー: --append の range は「シート名!列文字」（例: ログ!A）で指定してください")
            return False
        ur = ws.UsedRange
        next_row = ur.Row + ur.Rows.Count if ur is not None else 1
        # 空シートでも UsedRange は $A$1 を返すため、そのままだと2行目から始まる
        try:
            if (ur is not None and ur.Rows.Count == 1 and ur.Columns.Count == 1
                    and ur.Row == 1 and ur.Column == 1
                    and ur.Cells(1, 1).Value is None):
                next_row = 1
        except Exception:
            pass
        rng = ws.Range(f"{col_part.upper()}{next_row}")
        print(f"追記位置: {ws.Name}!{rng.Address}（使用範囲の最終行の次）")
    else:
        sheet_opt = getattr(args, 'sheet_opt', None)
        whole = _whole_sheet_spec(wb, spec, sheet_opt)
        if whole is not None:
            print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
            print(f"  全セルが同じ値で上書きされる危険があるため、write-range では")
            print(f"  セル/範囲を明示してください（例: \"{whole}!A1\"）")
            return False
        ws, rng = _resolve_range(xl, wb, spec, sheet_opt)

    raw = getattr(args, 'raw', False)

    # 書き込みは Worksheet_Change 等のイベントマクロを同期発火させる。そのマクロが
    # MsgBox を出すと rng.Value 代入がそこでブロックし、閉じるまでコマンドが無言で
    # ハングする（2026-07-11 実害）。安全解除の監視を書き込みの間だけ常設する。
    _dlg_watcher = _start_dialog_watcher(xl)
    try:
        if inline_value is not None:
            # インライン単一値: 範囲全体に同じ値 (数式可)
            if raw:
                rng.NumberFormat = "@"   # 文字列として保持（"007" 等の先頭ゼロを守る）
                rng.Value = inline_value
            else:
                rng.Value = _coerce_cell(inline_value)
            print(f"書き込み: {ws.Name}!{rng.Address} ← {inline_value}")
        else:
            path = (smart_path_resolve(tsv_opt) if tsv_opt else _LAST_VALUES_FILE)
            if not path or not os.path.exists(path):
                print(f"エラー: TSVが見つかりません: {tsv_opt or _LAST_VALUES_FILE}")
                print("  単一値ならインラインで: write-range A1 \"値\"")
                return False
            grid = _read_tsv_grid(path, raw=raw)
            if not grid:
                print("エラー: TSVが空です")
                return False
            nrows = len(grid)
            lens = {len(r) for r in grid}
            top = ws.Cells(rng.Row, rng.Column)
            if nrows == 1 and len(grid[0]) == 1:
                if raw:
                    top.NumberFormat = "@"
                top.Value = grid[0][0]
                print(f"書き込み: {ws.Name}!{top.Address} ← {grid[0][0]}")
            elif len(lens) == 1:
                ncols = len(grid[0])
                target = ws.Range(top, ws.Cells(rng.Row + nrows - 1,
                                                rng.Column + ncols - 1))
                if raw:
                    target.NumberFormat = "@"
                target.Value = grid
                print(f"書き込み: {ws.Name}!{target.Address} ← TSV {nrows}行 x {ncols}列")
            else:
                # 行の長さが不揃い: 矩形化して None を書くと右側の既存セルが消えるため、
                # 行ごとに実際の長さぶんだけ書き込む
                print(f"⚠ TSVの行の長さが不揃いです（{min(lens)}〜{max(lens)}列）。"
                      "行ごとに書き込み、短い行の右側セルには触れません。")
                for i, row in enumerate(grid):
                    if not row:
                        continue
                    r_tgt = ws.Range(ws.Cells(rng.Row + i, rng.Column),
                                     ws.Cells(rng.Row + i, rng.Column + len(row) - 1))
                    if raw:
                        r_tgt.NumberFormat = "@"
                    r_tgt.Value = (row,)
                print(f"書き込み: {ws.Name}!{top.Address} 起点 ← TSV {nrows}行（不揃い）")
    finally:
        _dlg_watcher.stop()

    note = _dialog_watcher_note(_dlg_watcher, None)
    if note:
        print(note, file=sys.stderr)
    print("（保存はしていません。Excelで確認後に保存してください）")
    return True


def cmd_clear_range(args):
    """セル範囲をクリア (既定: すべて)"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: clear-range [excel_file] <range> [--contents|--formats|--all]")
        return False
    spec = rest[0]
    if _reject_extra_args(rest, 1, '範囲は「シート名!範囲」の単一引数で指定してください'
                                   '（例: clear-range "Sheet1!A1:B2" --contents）'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, spec, sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域を本当にクリアするなら --whole-sheet を付けてください。")
        print(f"  範囲を消すつもりなら「シート名!範囲」で指定してください（例: \"{whole}!A1:B2\"）")
        return False
    ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
    if getattr(args, 'contents', False):
        rng.ClearContents(); what = "値"
    elif getattr(args, 'formats', False):
        rng.ClearFormats(); what = "書式"
    else:
        rng.Clear(); what = "すべて"
    print(f"クリア({what}): {ws.Name}!{rng.Address}")
    print("（保存はしていません）")
    return True


def cmd_format_range(args):
    """セル範囲に書式・整形を適用"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: format-range [excel_file] <range> [オプション...]")
        print("  --font 名 --size N --bold --unbold --italic")
        print("  --color '#RRGGBB' --bg '#RRGGBB' --number-format 書式")
        print("  --align left|center|right --valign top|center|bottom --wrap")
        print("  --border thin|medium|thick|hairline|none")
        print("  --col-width N --row-height N --merge --unmerge --autofit")
        return False
    spec = rest[0]
    if _reject_extra_args(rest, 1, '使い方: format-range [excel_file] <range> [オプション...]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, spec, sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域が同じ書式で塗られ、--merge 併用なら全域結合で値が消える危険があります。")
        print(f"  範囲を明示するか、本当に全域なら --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
    applied = []

    if getattr(args, 'font', None):
        rng.Font.Name = args.font; applied.append(f"font={args.font}")
    if getattr(args, 'size', None):
        rng.Font.Size = float(args.size); applied.append(f"size={args.size}")
    if getattr(args, 'bold', False):
        rng.Font.Bold = True; applied.append("bold")
    if getattr(args, 'unbold', False):
        rng.Font.Bold = False; applied.append("unbold")
    if getattr(args, 'italic', False):
        rng.Font.Italic = True; applied.append("italic")
    if getattr(args, 'color', None):
        rng.Font.Color = _hex_to_excel_color(args.color); applied.append(f"color={args.color}")
    if getattr(args, 'bg', None):
        rng.Interior.Color = _hex_to_excel_color(args.bg); applied.append(f"bg={args.bg}")
    if getattr(args, 'number_format', None):
        rng.NumberFormatLocal = args.number_format; applied.append(f"numfmt={args.number_format}")
    if getattr(args, 'align', None):
        rng.HorizontalAlignment = _XL_ALIGN_H[args.align]; applied.append(f"align={args.align}")
    if getattr(args, 'valign', None):
        rng.VerticalAlignment = _XL_ALIGN_V[args.valign]; applied.append(f"valign={args.valign}")
    if getattr(args, 'wrap', False):
        rng.WrapText = True; applied.append("wrap")
    if getattr(args, 'border', None):
        if args.border == 'none':
            rng.Borders.LineStyle = -4142            # xlNone
            applied.append("border=none")
        else:
            rng.Borders.LineStyle = 1                # xlContinuous
            rng.Borders.Weight = _XL_BORDER_WEIGHT.get(args.border, 2)
            applied.append(f"border={args.border}")
    if getattr(args, 'col_width', None) is not None:
        rng.ColumnWidth = float(args.col_width); applied.append(f"col-width={args.col_width}")
    if getattr(args, 'row_height', None) is not None:
        rng.RowHeight = float(args.row_height); applied.append(f"row-height={args.row_height}")
    if getattr(args, 'merge', False):
        # 複数の値を含む範囲の Merge は Excel が確認ダイアログを出し、
        # CLI が無言で応答待ちブロックする。値の個数を先に数えて、
        # 消える場合は警告した上で DisplayAlerts を切って実行する（左上の値が残る）
        n_vals = 0
        try:
            n_vals = int(xl.WorksheetFunction.CountA(rng))
        except Exception:
            pass
        if n_vals > 1:
            print(f"⚠ 範囲に値が{n_vals}個あります。結合により左上以外の値は消えます。")
        xl.DisplayAlerts = False
        try:
            rng.Merge()
        finally:
            xl.DisplayAlerts = True
        applied.append("merge")
    if getattr(args, 'unmerge', False):
        rng.UnMerge(); applied.append("unmerge")
    if getattr(args, 'autofit', False):
        rng.Columns.AutoFit(); applied.append("autofit")
    if getattr(args, 'lock', False):
        rng.Locked = True; applied.append("lock（シート保護時に有効）")
    if getattr(args, 'unlock', False):
        rng.Locked = False; applied.append("unlock（シート保護時に有効）")

    if not applied:
        print("書式オプションが指定されていません。--bold --bg '#FFFF00' などを指定してください。")
        return False
    print(f"書式適用: {ws.Name}!{rng.Address}  [{', '.join(applied)}]")
    print("（保存はしていません）")
    return True


def cmd_sheet(args):
    """シート操作: add/delete/rename/copy/activate/show/hide"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: sheet [excel_file] <add|delete|rename|copy|activate|show|hide> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    def find_sheet(name):
        for sh in wb.Sheets:
            if sh.Name == name:
                return sh
        return None

    if action == 'add':
        new_name = rest[1] if len(rest) >= 2 else None
        before = getattr(args, 'before', None)
        after = getattr(args, 'after', None)
        # --before/--after の対象が実在しないと find_sheet が None になり、
        # 無言でアクティブシート手前に追加されてしまうため先に検証する
        if before and find_sheet(before) is None:
            print(f"エラー: --before のシート '{before}' が見つかりません"); return False
        if after and find_sheet(after) is None:
            print(f"エラー: --after のシート '{after}' が見つかりません"); return False
        # 名前重複は Add 後の rename で例外→既定名シートの残骸になるため先に検証
        if new_name and find_sheet(new_name) is not None:
            print(f"エラー: シート '{new_name}' は既に存在します"); return False
        if before:
            sh = wb.Sheets.Add(find_sheet(before))
        elif after:
            sh = wb.Sheets.Add(None, find_sheet(after))
        else:
            sh = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
        if new_name:
            sh.Name = new_name
        print(f"シート追加: {sh.Name}")
    elif action == 'delete':
        if len(rest) < 2:
            print("使い方: sheet delete <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        xl.DisplayAlerts = False
        try:
            sh.Delete()
        finally:
            xl.DisplayAlerts = True
        print(f"シート削除: {rest[1]}")
    elif action == 'rename':
        if len(rest) < 3:
            print("使い方: sheet rename <old> <new>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        sh.Name = rest[2]
        print(f"シート名変更: {rest[1]} → {rest[2]}")
    elif action == 'copy':
        if len(rest) < 2:
            print("使い方: sheet copy <name> [newname]"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        sh.Copy(None, sh)
        newsh = wb.ActiveSheet
        if len(rest) >= 3:
            newsh.Name = rest[2]
        print(f"シート複製: {rest[1]} → {newsh.Name}")
    elif action == 'activate':
        if len(rest) < 2:
            print("使い方: sheet activate <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        sh.Activate()
        print(f"アクティブ化: {rest[1]}")
    elif action in ('protect', 'unprotect'):
        if len(rest) < 2:
            print(f"使い方: sheet {action} <name> [--password パスワード]"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        pw = getattr(args, 'password', None)
        if action == 'protect':
            sh.Protect(Password=pw) if pw else sh.Protect()
            print(f"シート保護: {rest[1]}" + ("（パスワード付き）" if pw else ""))
            print("  ※ format-range --lock/--unlock で設定した Locked がここで効きます")
        else:
            try:
                sh.Unprotect(Password=pw) if pw else sh.Unprotect()
            except Exception as e:
                print(f"エラー: 保護解除に失敗しました（パスワード違いの可能性）: {e}")
                return False
            print(f"シート保護解除: {rest[1]}")
    elif action in ('show', 'hide', 'very-hide'):
        if len(rest) < 2:
            print(f"使い方: sheet {action} <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        vis = {'show': -1, 'hide': 0, 'very-hide': 2}[action]   # xlVisible/-Hidden/-VeryHidden
        sh.Visible = vis
        label = {'show': '表示', 'hide': '非表示', 'very-hide': '完全非表示(VBAのみ解除可)'}[action]
        print(f"シート{label}: {rest[1]}")
    elif action == 'visibility':
        if len(rest) < 2:
            print("使い方: sheet visibility <name>"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        label = {-1: 'visible（表示）', 0: 'hidden（非表示）',
                 2: 'veryhidden（完全非表示）'}.get(int(sh.Visible), str(sh.Visible))
        print(f"表示状態: {rest[1]} = {label}")
        return True
    elif action == 'tab-color':
        if len(rest) < 2:
            print("使い方: sheet tab-color <name> [#RRGGBB | R G B | --clear]"); return False
        sh = find_sheet(rest[1])
        if not sh:
            print(f"エラー: シート '{rest[1]}' が見つかりません"); return False
        if getattr(args, 'clear', False):
            sh.Tab.ColorIndex = -4142             # xlColorIndexNone
            print(f"タブ色クリア: {rest[1]}")
        elif len(rest) >= 5 and all(x.isdigit() for x in rest[2:5]):
            r, g, b = int(rest[2]), int(rest[3]), int(rest[4])
            sh.Tab.Color = r + g * 256 + b * 65536
            print(f"タブ色設定: {rest[1]} = RGB({r},{g},{b})")
        elif len(rest) >= 3:
            sh.Tab.Color = _hex_to_excel_color(rest[2])
            print(f"タブ色設定: {rest[1]} = {rest[2]}")
        else:
            if int(sh.Tab.ColorIndex) == -4142:
                print(f"タブ色: {rest[1]} = （未設定）")
            else:
                c = int(sh.Tab.Color)
                r = c & 255; g = (c >> 8) & 255; b = (c >> 16) & 255
                print(f"タブ色: {rest[1]} = #{r:02X}{g:02X}{b:02X} (R={r},G={g},B={b})")
            return True
    else:
        print(f"未知のアクション: {action}")
        return False

    print("（保存はしていません）")
    return True


def cmd_table(args):
    """テーブル(ListObject)操作: create/list/delete"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: table [excel_file] <create|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
            for lo in sh.ListObjects:
                cnt += 1
                print(f"[{sh.Name}] {lo.Name}  範囲={lo.Range.Address}")
        if cnt == 0:
            print("テーブルはありません。")
        return True

    if action == 'create':
        if len(rest) < 2:
            print("使い方: table create <range> [name] [--no-headers]"); return False
        ws, rng = _resolve_range(xl, wb, rest[1])
        has_headers = 2 if getattr(args, 'no_headers', False) else 1  # xlNo=2 / xlYes=1
        lo = ws.ListObjects.Add(1, rng, None, has_headers)             # xlSrcRange=1
        if len(rest) >= 3:
            lo.Name = rest[2]
        print(f"テーブル作成: [{ws.Name}] {lo.Name}  範囲={lo.Range.Address}")
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: table delete <name>"); return False
        name = rest[1]
        for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
            for lo in sh.ListObjects:
                if lo.Name == name:
                    lo.Unlist()                    # テーブル解除 (データは残す)
                    print(f"テーブル解除: [{sh.Name}] {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: テーブル '{name}' が見つかりません")
        return False

    # ---- 以降は <table名> を rest[1] に取る列・フィルタ・ソート操作 ----
    def _find_lo(name):
        for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
            for lo in sh.ListObjects:
                if lo.Name == name:
                    return sh, lo
        return None, None

    def _col_field(lo, col_name):
        """テーブル内の列番号(1始まり)を名前から得る。無ければ None。"""
        for i in range(1, lo.ListColumns.Count + 1):
            if lo.ListColumns.Item(i).Name == col_name:
                return i
        return None

    if action == 'column':
        # table column <add|remove|rename|format> <table> ...
        sub = rest[1].lower() if len(rest) >= 2 else ''
        tname = rest[2] if len(rest) >= 3 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません（table list で確認）。"); return False
        if sub == 'add':
            col_name = rest[3] if len(rest) >= 4 else None
            pos = getattr(args, 'at', None)
            lc = lo.ListColumns.Add(int(pos)) if pos else lo.ListColumns.Add()
            if col_name:
                lc.Name = col_name
            print(f"列追加: {tname}[{lc.Name}]（位置 {lc.Index}）")
            print("（保存はしていません）"); return True
        if sub == 'remove':
            col_name = rest[3] if len(rest) >= 4 else None
            if _col_field(lo, col_name) is None:
                print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
            lo.ListColumns.Item(col_name).Delete()
            print(f"列削除: {tname}[{col_name}]"); print("（保存はしていません）"); return True
        if sub == 'rename':
            if len(rest) < 5:
                print("使い方: table column rename <table> <旧列> <新列>"); return False
            old, new = rest[3], rest[4]
            if _col_field(lo, old) is None:
                print(f"エラー: 列 '{old}' が見つかりません。"); return False
            lo.ListColumns.Item(old).Name = new
            print(f"列名変更: {tname}[{old}] → [{new}]"); print("（保存はしていません）"); return True
        if sub == 'format':
            col_name = rest[3] if len(rest) >= 4 else None
            if _col_field(lo, col_name) is None:
                print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
            lc = lo.ListColumns.Item(col_name)
            if len(rest) >= 5:                       # set
                lc.DataBodyRange.NumberFormat = rest[4]
                print(f"列書式設定: {tname}[{col_name}] = {rest[4]}")
                print("（保存はしていません）"); return True
            else:                                    # get
                try:
                    fmt = lc.DataBodyRange.Cells(1, 1).NumberFormat
                except Exception:
                    fmt = '(取得不可)'
                print(f"列書式: {tname}[{col_name}] = {fmt}"); return True
        print("使い方: table column <add|remove|rename|format> <table> ...")
        return False

    if action in ('filter', 'filter-values', 'filter-clear', 'filters'):
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。"); return False
        if action == 'filter-clear':
            try:
                if lo.AutoFilter is not None:
                    lo.AutoFilter.ShowAllData()
                print(f"フィルタ解除: {tname}")
            except Exception:
                print(f"フィルタは設定されていません: {tname}")
            print("（保存はしていません）"); return True
        if action == 'filters':
            af = lo.AutoFilter
            if af is None:
                print(f"フィルタなし: {tname}"); return True
            print(f"--- {tname} のフィルタ ---")
            any_on = False
            for i in range(1, lo.ListColumns.Count + 1):
                fl = af.Filters.Item(i)
                try:
                    on = fl.On
                except Exception:
                    on = False
                if on:
                    any_on = True
                    try:
                        c1 = fl.Criteria1
                    except Exception:
                        c1 = '(?)'
                    print(f"  {lo.ListColumns.Item(i).Name}: {c1}")
            if not any_on:
                print("  (フィルタ条件なし)")
            return True
        # filter / filter-values は列指定が必要
        col_name = rest[2] if len(rest) >= 3 else None
        field = _col_field(lo, col_name)
        if field is None:
            print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
        if action == 'filter':
            crit = rest[3] if len(rest) >= 4 else None
            if not crit:
                print("使い方: table filter <table> <列> <条件>（例: \">100\" \"=Active\"）"); return False
            lo.Range.AutoFilter(Field=field, Criteria1=crit)
            print(f"フィルタ適用: {tname}[{col_name}] {crit}")
            print("（保存はしていません）"); return True
        if action == 'filter-values':
            vals = rest[3:]
            if not vals:
                print("使い方: table filter-values <table> <列> 値1 値2 ..."); return False
            lo.Range.AutoFilter(Field=field, Criteria1=list(vals), Operator=7)   # xlFilterValues
            print(f"フィルタ適用(値): {tname}[{col_name}] {vals}")
            print("（保存はしていません）"); return True

    if action in ('sort', 'sort-multi'):
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。"); return False
        so = lo.Sort
        so.SortFields.Clear()
        if action == 'sort':
            col_name = rest[2] if len(rest) >= 3 else None
            if _col_field(lo, col_name) is None:
                print(f"エラー: 列 '{col_name}' が見つかりません。"); return False
            order = 2 if getattr(args, 'desc', False) else 1   # xlDescending/xlAscending
            so.SortFields.Add(lo.ListColumns.Item(col_name).Range, 0, order)
            so.Apply()
            print(f"ソート: {tname} {col_name} {'降順' if order == 2 else '昇順'}")
            print("（保存はしていません）"); return True
        else:  # sort-multi  col:asc col:desc ...
            specs = rest[2:]
            if not specs:
                print("使い方: table sort-multi <table> 列:asc 列:desc ..."); return False
            applied = []
            for spec in specs:
                if ':' in spec:
                    cn, od = spec.rsplit(':', 1)
                else:
                    cn, od = spec, 'asc'
                if _col_field(lo, cn) is None:
                    print(f"エラー: 列 '{cn}' が見つかりません。"); return False
                order = 2 if od.lower().startswith('d') else 1
                so.SortFields.Add(lo.ListColumns.Item(cn).Range, 0, order)
                applied.append(f"{cn}{'↓' if order == 2 else '↑'}")
            so.Apply()
            print(f"複数ソート: {tname} {' / '.join(applied)}")
            print("（保存はしていません）"); return True

    if action == 'read':
        # テーブル名で直接読む（従来は table list で番地を得て read-range する2段＝2接続だった）
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。（table list で確認）"); return False
        rng = lo.Range
        print(f"テーブル: {tname}   [{sh.Name}!{rng.Address}]")
        print("=" * 60)
        print(_values_to_grid(rng))
        print("=" * 60)
        tsv_out = getattr(args, 'tsv_out', None)
        if tsv_out is not None:
            path = _LAST_VALUES_FILE if tsv_out == '_DEFAULT_' else os.path.abspath(tsv_out)
            rows = _range_values_2d(rng)
            with open(path, 'w', encoding='utf-8-sig', newline='') as f:
                f.write('\n'.join('\t'.join(_cell_str(v) for v in r) for r in rows) + '\n')
            print(f"TSV書き出し: {path}  ({len(rows)}行)")
            print(f"  編集後の書き戻し: py vba_manager.py write-range \"{sh.Name}!{_col_letter(rng.Column)}{rng.Row}\"")
            print(f"  （\"007\" 等の先頭ゼロを数値化させたくない場合は --raw を付ける）")
        return True

    if action == 'ref':
        tname = rest[1] if len(rest) >= 2 else None
        sh, lo = _find_lo(tname) if tname else (None, None)
        if not lo:
            print(f"エラー: テーブル '{tname}' が見つかりません。"); return False
        col_name = rest[2] if len(rest) >= 3 else None
        if col_name:
            print(f"構造化参照: {tname}[{col_name}]")
        else:
            cols = [lo.ListColumns.Item(i).Name for i in range(1, lo.ListColumns.Count + 1)]
            print(f"構造化参照: {tname}[#All] / 列: {', '.join(cols)}")
        return True

    print(f"未知のアクション: {action}")
    return False


def cmd_name(args):
    """名前付き範囲操作: add/list/delete"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: name [excel_file] <add|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for nm in wb.Names:
            cnt += 1
            try:
                refers = nm.RefersTo
            except Exception:
                refers = '(?)'
            print(f"{nm.Name}  →  {refers}")
        if cnt == 0:
            print("名前付き範囲はありません。")
        return True

    if action == 'add':
        if len(rest) < 3:
            print("使い方: name add <name> <range>"); return False
        nm_name = rest[1]
        ws, rng = _resolve_range(xl, wb, rest[2])
        # rng.Address は既定で絶対参照 ($A$2)。pywin32 ではプロパティなので引数なしで使う
        refers = "='" + ws.Name.replace("'", "''") + "'!" + rng.Address
        wb.Names.Add(nm_name, refers)
        print(f"名前付き範囲を追加: {nm_name} → {refers}")
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: name delete <name>"); return False
        nm_name = rest[1]
        # 完全一致を優先。シートスコープ名（'Sheet1'!名前）の末尾一致は
        # 同名が複数シートにあると最初の1個を消す取り違えになるため、
        # 複数一致ならエラーで止めて候補を出す。
        exact = [nm for nm in wb.Names if nm.Name == nm_name]
        if exact:
            exact[0].Delete()
            print(f"名前付き範囲を削除: {nm_name}")
            print("（保存はしていません）")
            return True
        suffix = [nm for nm in wb.Names if nm.Name.split('!')[-1] == nm_name]
        if len(suffix) == 1:
            actual = suffix[0].Name
            suffix[0].Delete()
            print(f"名前付き範囲を削除: {actual}")
            print("（保存はしていません）")
            return True
        if len(suffix) > 1:
            print(f"エラー: 名前 '{nm_name}' はシート違いで複数あります。完全名で指定してください:")
            for nm in suffix:
                print(f"  {nm.Name}")
            return False
        print(f"エラー: 名前 '{nm_name}' が見つかりません")
        return False

    print(f"未知のアクション: {action}")
    return False


# ================================================================
# 「手」コマンド 第2弾 (編集の足回り / 検索置換 / 保存印刷 / 仕上げ)
#   ※ いずれもアクティブ(開いたまま)のブックに COM で直接作用。
#      save 系を除き既定では保存しない。
# ================================================================

def _col_num(s):
    """列文字(A,B,..,AA) を列番号(1始まり)に変換"""
    n = 0
    for ch in s.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


# ---- a. 編集の足回り ----

def _sheet_or_active(wb, sheet_name):
    """--sheet 指定があればそのシート、なければアクティブシートを返す。
    指定シートが見つからなければエラー表示して None。"""
    if not sheet_name:
        return wb.ActiveSheet
    for sh in wb.Worksheets:
        if sh.Name == sheet_name:
            return sh
    print(f"エラー: シート '{sheet_name}' が見つかりません")
    return None


def cmd_row(args):
    """行の挿入・削除: row <insert|delete> <行番号> [本数]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: row <insert|delete> <行番号> [本数]")
        return False
    action = rest[0].lower()
    start = int(rest[1])
    count = int(rest[2]) if len(rest) >= 3 else 1
    xl, wb = get_workbook(target_file)
    ws = _sheet_or_active(wb, getattr(args, 'sheet', None))
    if ws is None:
        return False
    if action == 'delete':
        # 破壊操作は実行前に対象を明示する（対象取り違え事故の防止）
        print(f"対象シート: {ws.Name}（{wb.Name}）")
    rng = ws.Rows(f"{start}:{start + count - 1}")
    if action == 'insert':
        rng.Insert()
        print(f"行挿入: {ws.Name} {start}行目に {count}行")
    elif action == 'delete':
        rng.Delete()
        print(f"行削除: {ws.Name} {start}〜{start + count - 1}行")
    else:
        print(f"未知のアクション: {action}（insert|delete）"); return False
    print("（保存はしていません）")
    return True


def cmd_col(args):
    """列の挿入・削除: col <insert|delete> <列文字> [本数]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: col <insert|delete> <列文字> [本数]")
        return False
    action = rest[0].lower()
    start = rest[1]
    count = int(rest[2]) if len(rest) >= 3 else 1
    end = _col_letter(_col_num(start) + count - 1)
    xl, wb = get_workbook(target_file)
    ws = _sheet_or_active(wb, getattr(args, 'sheet', None))
    if ws is None:
        return False
    if action == 'delete':
        # 破壊操作は実行前に対象を明示する（対象取り違え事故の防止）
        print(f"対象シート: {ws.Name}（{wb.Name}）")
    rng = ws.Columns(f"{start}:{end}")
    if action == 'insert':
        rng.Insert()
        print(f"列挿入: {ws.Name} {start}列に {count}列")
    elif action == 'delete':
        rng.Delete()
        print(f"列削除: {ws.Name} {start}〜{end}列")
    else:
        print(f"未知のアクション: {action}（insert|delete）"); return False
    print("（保存はしていません）")
    return True


def cmd_copy_range(args):
    """範囲コピー: copy-range <src> <dst> [--values]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: copy-range <src> <dst> [--values]")
        return False
    if _reject_extra_args(rest, 2, '使い方: copy-range <src> <dst> [--values]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, rest[0], sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: コピー元 '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域が複写先を上書きする危険があるため、範囲を明示するか --whole-sheet を付けてください。")
        return False
    # コピー先も同じガード。シート名だけ渡すと UsedRange 全域が貼り付け先になり、
    # コピー元が単一セル等だと全域タイル上書きになる（2026-07-09 再点検で発見）
    whole_d = _whole_sheet_spec(wb, rest[1])
    if whole_d is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: コピー先 '{rest[1]}' はシート '{whole_d}' の使用範囲全域を指します。")
        print(f"  全域がタイル状に上書きされる危険があるため、貼り付け先の左上セルを明示するか --whole-sheet を付けてください。")
        return False
    ws_s, rng_s = _resolve_range(xl, wb, rest[0], sheet_opt)
    ws_d, rng_d = _resolve_range(xl, wb, rest[1])
    if getattr(args, 'values', False):
        try:
            rng_s.Copy()
            rng_d.PasteSpecial(-4163)          # xlPasteValues
        finally:
            # 失敗しても Excel にコピーの点線（CutCopyMode）を残さない
            try:
                xl.CutCopyMode = False
            except Exception:
                pass
        print(f"コピー(値のみ): {ws_s.Name}!{rng_s.Address} → {ws_d.Name}!{rng_d.Address}")
    else:
        rng_s.Copy(rng_d)
        print(f"コピー(書式・式込): {ws_s.Name}!{rng_s.Address} → {ws_d.Name}!{rng_d.Address}")
    print("（保存はしていません）")
    return True


def cmd_fill(args):
    """オートフィル: fill <range> [--right]（既定は下方向）"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: fill <range> [--right]   範囲の先頭セルを残りに複写")
        return False
    if _reject_extra_args(rest, 1, '使い方: fill <range> [--right]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, rest[0], sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  先頭行/列が全域に複写される危険があるため、範囲を明示するか --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0], sheet_opt)
    if getattr(args, 'right', False):
        rng.FillRight(); direction = "右"
    else:
        rng.FillDown(); direction = "下"
    print(f"フィル({direction}): {ws.Name}!{rng.Address}")
    print("（保存はしていません）")
    return True


def cmd_sort(args):
    """並べ替え: sort <range> [--key 列文字] [--desc] [--header|--no-header]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: sort <range> [--key 列文字] [--desc] [--header|--no-header]")
        return False
    if _reject_extra_args(rest, 1, '使い方: sort <range> [--key 列文字] [--desc] [--header|--no-header]'):
        return False
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    whole = _whole_sheet_spec(wb, rest[0], sheet_opt)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  全域を並べ替えるなら --whole-sheet を付けてください（--header の明示も推奨）。")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0], sheet_opt)
    keycol = getattr(args, 'key', None)
    if keycol and not keycol.isalpha():
        # _col_num は英字以外も黙って数値化してしまい、"A1" が K列扱いになる
        print(f"エラー: --key は列文字（A〜XFD）で指定してください: '{keycol}'")
        return False
    key_idx = _col_num(keycol) if keycol else rng.Column
    keycell = ws.Cells(rng.Row, key_idx)
    order = 2 if getattr(args, 'desc', False) else 1       # xlDescending=2 / xlAscending=1
    if getattr(args, 'header', False):
        header = 1                                          # xlYes
    elif getattr(args, 'no_header', False):
        header = 2                                          # xlNo
    else:
        header = 0                                          # xlGuess
    # Orientation / MatchCase は Excel が「前回の並べ替え設定」をシートに保存して
    # 引き継ぐ仕様。未指定だと手動の列単位ソート等が引き継がれ、行方向のつもりが
    # 列方向に並べ替わる事故になるため必ず明示する（xlTopToBottom=1）
    rng.Sort(Key1=keycell, Order1=order, Header=header,
             Orientation=1, MatchCase=False)
    print(f"並べ替え: {ws.Name}!{rng.Address}  キー列={keycol or _col_letter(rng.Column)}  "
          f"{'降順' if order == 2 else '昇順'}")
    print("（保存はしていません）")
    return True


def cmd_autofilter(args):
    """オートフィルタ: autofilter [range] [--off]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if _reject_extra_args(rest, 1, '使い方: autofilter [range] [--off]'):
        return False
    xl, wb = get_workbook(target_file)
    if getattr(args, 'off', False):
        # 位置引数（範囲/シート名）があればそのシートを解除対象にする。
        # 黙って ActiveSheet に落とすと、指定したつもりの別シートでなく
        # アクティブなシートのフィルタ（絞り込み条件ごと）が消える
        if rest:
            ws, _ = _resolve_range(xl, wb, rest[0])
        else:
            ws = wb.ActiveSheet
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
            print(f"オートフィルタ解除: {ws.Name}")
            print("（保存はしていません）")
        else:
            print(f"オートフィルタは設定されていません: {ws.Name}")
        return True
    spec = rest[0] if rest else None
    ws, rng = _resolve_range(xl, wb, spec)
    if ws.AutoFilterMode:
        print(f"既にオートフィルタが設定されています: {ws.Name}")
    else:
        # win32com の遅延バインディングでは引数なし rng.AutoFilter() が
        # 「AutoFilter メソッドが失敗しました」で落ちる（全省略可能引数を省くとCOMが弾く）。
        # Field:=1 だけ渡すと Criteria なし＝どの列も絞り込まずに範囲全体へフィルタUIを付ける
        # ＝引数なしと同じ「オートフィルタON」になる（実弾で確認済み）。
        rng.AutoFilter(1)
        print(f"オートフィルタ設定: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
    return True


# ---- b. 検索・置換 ----

def cmd_find(args):
    """セル検索: find <文字> [--book] [--whole] [--formula]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: find <文字> [--book(全シート)] [--whole(完全一致)] [--formula(式も検索)]")
        return False
    needle = rest[0]
    xl, wb = get_workbook(target_file)
    # グラフシートには UsedRange が無く例外で検索全体が落ちるため Worksheets 限定
    sheets = list(wb.Worksheets) if getattr(args, 'book', False) else [wb.ActiveSheet]
    look_in = -4123 if getattr(args, 'formula', False) else -4163  # xlFormulas / xlValues
    look_at = 1 if getattr(args, 'whole', False) else 2            # xlWhole / xlPart
    total = 0
    max_hits = getattr(args, 'max_hits', None) or 200
    for ws in sheets:
        try:
            rng = ws.UsedRange
        except Exception:
            # アクティブがグラフシート等だと UsedRange 自体が例外になる
            print(f"（'{ws.Name}' はワークシートではないためスキップ）")
            continue
        try:
            cell = rng.Find(What=needle, LookIn=look_in, LookAt=look_at, MatchCase=False)
        except Exception:
            cell = None
        first = None
        while cell is not None:
            addr = cell.Address
            if first is None:
                first = addr
            elif addr == first:
                break
            total += 1
            if total <= max_hits:
                # 「シート名!$A$1」はそのまま write-range 等の range 引数に貼れる形
                print(f"{ws.Name}!{addr}: {cell.Value}")
            cell = rng.FindNext(cell)
    if total == 0:
        print(f"'{needle}' は見つかりませんでした。")
    else:
        if total > max_hits:
            print(f"…他 {total - max_hits}件（--max で上限変更可）")
        print(f"--- {total}件 ヒット ---")
    return True


def cmd_find_replace(args):
    """一括置換: find-replace <検索> <置換> [range] [--whole] [--wildcard]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: find-replace <検索> <置換> [range] [--whole] [--wildcard]")
        return False
    needle, repl = rest[0], rest[1]
    spec = rest[2] if len(rest) >= 3 else None
    if _reject_extra_args(rest, 3, '使い方: find-replace <検索> <置換> [range] [--whole] [--match-case] [--wildcard]'):
        return False
    # Excel の Find/Replace は * ? を常にワイルドカード解釈する。
    # 素通しすると `find-replace "*" "×"` が全非空セルの丸ごと置換になるため、
    # 既定は ~ エスケープで文字どおりに扱い、--wildcard で明示オプトインする
    if not getattr(args, 'wildcard', False):
        escaped = re.sub(r'([~*?])', r'~\1', needle)
        if escaped != needle:
            print("（* ? ~ は文字どおりに置換します。パターンとして使うなら --wildcard）")
        needle = escaped
    xl, wb = get_workbook(target_file)
    sheet_opt = getattr(args, 'sheet_opt', None)
    if spec or sheet_opt:
        ws, rng = _resolve_range(xl, wb, spec, sheet_opt)
    else:
        ws = wb.ActiveSheet
        try:
            rng = ws.UsedRange
        except Exception:
            print(f"エラー: アクティブシート '{ws.Name}' は置換できる"
                  "ワークシートではありません（グラフシート等）。")
            return False
    look_at = 1 if getattr(args, 'whole', False) else 2
    match_case = getattr(args, 'match_case', False)
    # Range.Replace は置換件数を返さないため、置換前にヒットセル数を数える
    # （LookIn は Replace と同じ数式(-4123)で揃える）
    count = 0
    first = None
    cell = rng.Find(What=needle, LookAt=look_at, LookIn=-4123, MatchCase=match_case)
    while cell is not None:
        addr = cell.Address
        if first is None:
            first = addr
        elif addr == first:
            break
        count += 1
        cell = rng.FindNext(cell)
    if count == 0:
        print(f"'{needle}' は {ws.Name}!{rng.Address} に見つかりませんでした（置換なし）")
        return True
    rng.Replace(What=needle, Replacement=repl, LookAt=look_at, MatchCase=match_case)
    print(f"置換: {ws.Name}!{rng.Address}  '{needle}' → '{repl}'  （{count}セルにヒット）")
    print("（保存はしていません）")
    return True


# ---- c. 保存・印刷まわり ----

def cmd_save(args):
    """上書き保存: save [excel_file]"""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    xl.DisplayAlerts = False
    try:
        wb.Save()
    finally:
        xl.DisplayAlerts = True
    print(f"保存しました: {wb.FullName}")
    return True


def cmd_save_as(args):
    """別名保存: save-as [excel_file] <path>（省略時はアクティブブックを対象）"""
    # 他コマンドと同じ流儀: 引数2つなら第1引数を対象ブック、第2引数を出力パスとする。
    # （以前は rest[0] を無条件に出力パスにしていたため、
    #   `save-as 既存ブック.xlsx 新名.xlsx` で既存ブックが無言上書きされる罠があった）
    rest = list(args.posargs)
    if not rest:
        print("使い方: save-as [excel_file] <出力path> [--overwrite]")
        return False
    if len(rest) >= 3:
        print("エラー: 引数が多すぎます。使い方: save-as [excel_file] <出力path> [--overwrite]")
        return False
    if len(rest) == 2:
        target_file, out_arg = rest[0], rest[1]
    else:
        target_file, out_arg = None, rest[0]
    out = os.path.abspath(out_arg)

    FMT = {'.xlsx': 51, '.xlsm': 52, '.xlsb': 50, '.xls': 56,
           '.csv': 6, '.txt': -4158}
    ext = os.path.splitext(out)[1].lower()
    if ext not in FMT:
        # 未知拡張子を黙って xlsx にフォールバックすると「中身xlsxの .pdf」等の壊れファイルになる
        print(f"エラー: 対応していない拡張子です: '{ext or '(なし)'}'")
        print(f"  対応: {' '.join(sorted(FMT))}")
        return False
    fmt = FMT[ext]

    if os.path.exists(out) and not getattr(args, 'overwrite', False):
        print(f"エラー: 出力先が既に存在します: {out}")
        print("  上書きするなら --overwrite を付けてください。")
        return False

    xl, wb = get_workbook(target_file)
    try:
        old_full = wb.FullName
    except Exception:
        old_full = None
    src_ext = os.path.splitext(wb.Name)[1].lower()
    if src_ext in ('.xlsm', '.xlsb', '.xls') and ext == '.xlsx':
        # DisplayAlerts=False で Excel の警告が出ないため、こちらで明示する
        print("⚠ 注意: マクロ付きブックを .xlsx で保存するため、VBAマクロは保存されません。")
    if ext in ('.csv', '.txt'):
        try:
            n_sheets = wb.Sheets.Count
        except Exception:
            n_sheets = 1
        if n_sheets > 1:
            # これも DisplayAlerts=False で Excel 側の警告が抑止されるため明示する
            print(f"⚠ 注意: {ext} はアクティブシート1枚しか保存されません"
                  f"（このブックは {n_sheets} シート）。")
    xl.DisplayAlerts = False
    try:
        wb.SaveAs(out, FileFormat=fmt)
    finally:
        xl.DisplayAlerts = True
    # batch/shell/MCP の1接続セッションでは接続キャッシュが旧パスキーのまま残り、
    # 旧名を指定した後続コマンドが改名後のブックに当たる（対象取り違え）。
    # SaveAs 成功時にキャッシュのキーを新パスへ付け替える
    if old_full:
        old_key = old_full.lower()
        if old_key != out.lower() and old_key in _wb_cache:
            _wb_cache[out.lower()] = _wb_cache.pop(old_key)
    print(f"別名保存しました: {out}")
    print("  （以後、開いているブックの保存先はこの新パスになります）")
    return True


def cmd_export_pdf(args):
    """PDF出力: export-pdf [excel_file] <出力.pdf> [--sheet 名 | --range "シート!範囲"]

    ExportAsFixedFormat による出力。既定はブック全体、--sheet で1シート、
    --range で範囲のみ。ブック自体は変更しない（保存フラグも汚さない）。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: export-pdf [excel_file] <出力.pdf> [--sheet 名 | --range \"シート!A1:H50\"] [--overwrite]")
        return False
    out = os.path.abspath(rest[0])
    if _reject_extra_args(rest, 1, '出力パスは1つだけ指定してください'):
        return False
    if not out.lower().endswith('.pdf'):
        print(f"エラー: 出力は .pdf で指定してください: {out}")
        return False
    if os.path.exists(out) and not getattr(args, 'overwrite', False):
        print(f"エラー: 出力先が既に存在します: {out}")
        print("  上書きするなら --overwrite を付けてください。")
        return False
    sheet_opt = getattr(args, 'sheet_opt', None)
    range_opt = getattr(args, 'range_opt', None)
    if sheet_opt and range_opt:
        print("エラー: --sheet と --range は同時に指定できません")
        return False

    xl, wb = get_workbook(target_file)
    if range_opt:
        ws, rng = _resolve_range(xl, wb, range_opt)
        rng.ExportAsFixedFormat(0, out)               # 0 = xlTypePDF
        scope = f"範囲 {ws.Name}!{rng.Address}"
    elif sheet_opt:
        ws = None
        for sh in wb.Worksheets:
            if sh.Name == sheet_opt:
                ws = sh
                break
        if ws is None:
            print(f"エラー: シート '{sheet_opt}' が見つかりません")
            return False
        ws.ExportAsFixedFormat(0, out)
        scope = f"シート '{ws.Name}'"
    else:
        wb.ExportAsFixedFormat(0, out)
        scope = "ブック全体"
    print(f"PDF出力: {scope} → {out}")
    return True


def cmd_print_setup(args):
    """印刷設定: print-setup [--area R] [--title-rows 1:3] [--title-cols A:B] ..."""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    ws = wb.ActiveSheet
    ps = ws.PageSetup
    applied = []

    if getattr(args, 'area', None):
        ps.PrintArea = ws.Range(args.area).Address
        applied.append(f"area={args.area}")
    if getattr(args, 'title_rows', None):
        a, b = (args.title_rows.split(':') + [args.title_rows])[:2]
        ps.PrintTitleRows = f"${a}:${b}"
        applied.append(f"title-rows={args.title_rows}")
    if getattr(args, 'title_cols', None):
        a, b = (args.title_cols.split(':') + [args.title_cols])[:2]
        ps.PrintTitleColumns = f"${a}:${b}"
        applied.append(f"title-cols={args.title_cols}")
    if getattr(args, 'landscape', False):
        ps.Orientation = 2; applied.append("landscape")
    if getattr(args, 'portrait', False):
        ps.Orientation = 1; applied.append("portrait")
    if getattr(args, 'fit_wide', None) is not None:
        ps.Zoom = False; ps.FitToPagesWide = int(args.fit_wide)
        applied.append(f"fit-wide={args.fit_wide}")
    if getattr(args, 'fit_tall', None) is not None:
        ps.Zoom = False; ps.FitToPagesTall = int(args.fit_tall)
        applied.append(f"fit-tall={args.fit_tall}")
    if getattr(args, 'zoom', None) is not None:
        ps.Zoom = int(args.zoom); applied.append(f"zoom={args.zoom}")
    if getattr(args, 'center_h', False):
        ps.CenterHorizontally = True; applied.append("center-h")
    if getattr(args, 'center_v', False):
        ps.CenterVertically = True; applied.append("center-v")

    if not applied:
        print("オプションが指定されていません。--area / --title-rows / --landscape など。")
        return False
    print(f"印刷設定: {ws.Name}  [{', '.join(applied)}]")
    print("（保存はしていません）")
    return True


def cmd_printer_list(args):
    """プリンター一覧およびアクティブプリンターを取得"""
    target_file, _ = parse_target_and_rest(args.posargs)
    
    # Excelのアクティブプリンターを取得
    active_printer = None
    try:
        xl, wb = get_workbook(target_file)
        active_printer = xl.ActivePrinter
        print(f"現在のアクティブプリンター: {active_printer}")
    except Exception as e:
        print(f"警告: Excelからアクティブプリンターを取得できませんでした ({e})")

    # WMI経由でOSにインストールされているプリンター一覧を取得
    import win32com.client
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        printers = wmi.InstancesOf("Win32_Printer")
        print("\nインストールされているプリンター一覧:")
        for printer in printers:
            name = printer.Name
            status = " (選択中)" if active_printer and name in active_printer else ""
            print(f"  - {name}{status}")
    except Exception as e:
        print(f"エラー: インストールされているプリンター一覧を取得できませんでした ({e})")
        return False
    return True


def cmd_printer_setup(args):
    """プリンターの詳細設定（両面印刷・カラー等）を変更・表示"""
    # 対象プリンター名の決定
    printer_name = getattr(args, 'printer', None)
    if not printer_name:
        # Excelが起動していればそのアクティブプリンター名を使用、さもなくばデフォルトプリンター
        try:
            target_file, _ = parse_target_and_rest(args.posargs)
            xl, wb = get_workbook(target_file)
            raw_printer = xl.ActivePrinter
            if " on " in raw_printer:
                printer_name = raw_printer.split(" on ")[0]
            else:
                printer_name = raw_printer
        except Exception:
            import win32print
            printer_name = win32print.GetDefaultPrinter()

    if not printer_name:
        print("エラー: 対象プリンターが特定できません。--printer で指定してください。")
        return False

    print(f"対象プリンター: {printer_name}")

    import win32print
    try:
        # 設定変更に必要なアクセス権を指定 (PRINTER_ACCESS_ADMINISTER=4, PRINTER_ACCESS_USE=8)
        access = 4 | 8
        handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": access})
    except Exception:
        try:
            handle = win32print.OpenPrinter(printer_name)
        except Exception as e:
            print(f"エラー: プリンターを開けませんでした ({e})")
            return False

    try:
        info = win32print.GetPrinter(handle, 2)
        devmode = info["pDevMode"]
        if devmode is None:
            print("エラー: プリンターの構成情報 (DevMode) を取得できませんでした。")
            return False

        applied = []
        
        # 1. 両面印刷 (Duplex)
        duplex_opt = getattr(args, 'duplex', None)
        if duplex_opt:
            val = {'simplex': 1, 'vertical': 2, 'horizontal': 3}.get(duplex_opt.lower())
            if val:
                devmode.Duplex = val
                applied.append(f"duplex={duplex_opt}")
                
        # 2. カラー (Color)
        color_opt = getattr(args, 'color', None)
        if color_opt:
            val = {'mono': 1, 'color': 2}.get(color_opt.lower())
            if val:
                devmode.Color = val
                applied.append(f"color={color_opt}")

        # 3. 用紙向き (Orientation)
        orient_opt = getattr(args, 'orientation', None)
        if orient_opt:
            val = {'portrait': 1, 'landscape': 2}.get(orient_opt.lower())
            if val:
                devmode.Orientation = val
                applied.append(f"orientation={orient_opt}")

        if applied:
            win32print.SetPrinter(handle, 2, info, 0)
            print(f"プリンター設定更新完了: [{', '.join(applied)}]")
        else:
            # 現在の設定を表示
            duplex_names = {1: '片面 (simplex)', 2: '両面/長辺綴じ (vertical)', 3: '両面/短辺綴じ (horizontal)'}
            color_names = {1: 'モノクロ (mono)', 2: 'カラー (color)'}
            orient_names = {1: '縦向き (portrait)', 2: '横向き (landscape)'}
            
            d_val = duplex_names.get(devmode.Duplex, f"不明({devmode.Duplex})")
            c_val = color_names.get(devmode.Color, f"不明({devmode.Color})")
            o_val = orient_names.get(devmode.Orientation, f"不明({devmode.Orientation})")
            
            print(f"現在のプリンター構成:")
            print(f"  - 両面印刷: {d_val}")
            print(f"  - カラー　: {c_val}")
            print(f"  - 用紙向き: {o_val}")
            
    except Exception as e:
        print(f"エラー: プリンター設定の変更に失敗しました ({e})")
        return False
    finally:
        win32print.ClosePrinter(handle)
    return True


# ---- d. 仕上げ・見た目 ----

_XL_COND_OP = {'gt': 5, 'lt': 6, 'eq': 3, 'ne': 4, 'ge': 7, 'le': 8, 'between': 1}


def cmd_cond_format(args):
    """条件付き書式(セルの値): cond-format <range> --gt 100 --bg '#FFC7CE'"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: cond-format <range> [--gt|--lt|--ge|--le|--eq|--ne 値 | --between v1 v2]")
        print("         [--bg '#RRGGBB'] [--color '#RRGGBB'] [--bold] [--clear]")
        return False
    spec = rest[0]
    if _reject_extra_args(rest, 1, '使い方: cond-format [excel_file] <range> [--gt 値 ...|--clear]'):
        return False
    xl, wb = get_workbook(target_file)
    whole = _whole_sheet_spec(wb, spec)
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{spec}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  範囲を明示するか、本当に全域なら --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, spec)

    if getattr(args, 'clear', False):
        rng.FormatConditions.Delete()
        print(f"条件付き書式を全削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True

    formula = getattr(args, 'formula_opt', None)
    if formula:
        # 数式ベースのルール（xlExpression=2）。'=' 始まりの数式が TRUE のセルに書式。
        # 相対参照は範囲の左上セル基準（Excel の条件付き書式と同じ規約）
        if not formula.startswith('='):
            formula = '=' + formula
        # xlExpression は名前付き引数だと DISP_E_PARAMNOTOPTIONAL になるため位置渡し
        # （Operator は式タイプでは無視されるがスロットとして必要。1=xlBetween をダミーに）
        fc = rng.FormatConditions.Add(2, 1, formula)
    else:
        op = None; f1 = None; f2 = None
        for name in ('gt', 'lt', 'eq', 'ne', 'ge', 'le'):
            v = getattr(args, name, None)
            if v is not None:
                op = _XL_COND_OP[name]; f1 = str(v); break
        if op is None and getattr(args, 'between', None):
            op = _XL_COND_OP['between']; f1, f2 = args.between[0], args.between[1]
        if op is None:
            print("比較条件がありません。--gt 100 か --formula \"=数式\" を指定してください。")
            return False

        if f2 is not None:
            fc = rng.FormatConditions.Add(Type=1, Operator=op, Formula1=f1, Formula2=f2)
        else:
            fc = rng.FormatConditions.Add(Type=1, Operator=op, Formula1=f1)

    if getattr(args, 'bg', None):
        fc.Interior.Color = _hex_to_excel_color(args.bg)
    if getattr(args, 'color', None):
        fc.Font.Color = _hex_to_excel_color(args.color)
    if getattr(args, 'bold', False):
        fc.Font.Bold = True
    print(f"条件付き書式を追加: {ws.Name}!{rng.Address}")
    print("（保存はしていません）")
    return True


def cmd_hyperlink(args):
    """ハイパーリンク: hyperlink <cell> <url> [--text 表示文字] / --remove / --list

    url を省略して単セルを指定すると、そのセルのリンクを表示する（取得）。
    --list でシート内の全ハイパーリンクを一覧する。
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)

    list_opt = getattr(args, 'list_links', None)
    if list_opt is not None:
        # シート内の全リンク。シート名は --list シート名 か 位置引数、無ければアクティブ
        want = None
        if list_opt != "__ACTIVE__":
            want = list_opt
        elif rest:
            want = rest[0]
        if want:
            ws = None
            for sh in wb.Worksheets:
                if sh.Name == want:
                    ws = sh
                    break
            if ws is None:
                # 黙ってアクティブシートに落とすと「指定シートは0件」と誤読される
                print(f"エラー: シート '{want}' が見つかりません")
                return False
        else:
            ws = wb.ActiveSheet
        cnt = ws.Hyperlinks.Count
        print(f"--- {ws.Name} のハイパーリンク（{cnt}件） ---")
        for i in range(1, cnt + 1):
            hl = ws.Hyperlinks.Item(i)
            try:
                addr = hl.Range.Address
            except Exception:
                addr = '(図形)'
            sub = f"#{hl.SubAddress}" if getattr(hl, 'SubAddress', '') else ''
            print(f"  {ws.Name}!{addr}: {hl.Address or ''}{sub}")
        return True

    if not rest:
        print("使い方: hyperlink <cell> <url> [--text 表示文字]")
        print("       hyperlink <cell>            # そのセルのリンクを表示")
        print("       hyperlink <cell> --remove   # 削除")
        print("       hyperlink --list [シート名]  # シート内の全リンク一覧")
        return False
    if getattr(args, 'remove', False):
        # シート名だけの指定は使用範囲全域のリンク一括削除（書式も戻る）になるためガード
        whole = _whole_sheet_spec(wb, rest[0])
        if whole is not None and not getattr(args, 'whole_sheet', False):
            print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
            print(f"  セル/範囲を明示するか、全リンクを消すなら --whole-sheet を付けてください。")
            return False
    ws, rng = _resolve_range(xl, wb, rest[0])
    if getattr(args, 'remove', False):
        rng.Hyperlinks.Delete()
        print(f"ハイパーリンク削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True
    if len(rest) < 2:
        # 取得モード: そのセルのリンクを表示
        cell = rng.Cells(1, 1)
        if cell.Hyperlinks.Count == 0:
            print(f"{ws.Name}!{cell.Address}: ハイパーリンクはありません")
        else:
            hl = cell.Hyperlinks.Item(1)
            sub = f"#{hl.SubAddress}" if getattr(hl, 'SubAddress', '') else ''
            print(f"{ws.Name}!{cell.Address}: {hl.Address or ''}{sub}"
                  f"  表示=「{cell.Value}」")
        return True
    url = rest[1]
    cell = rng.Cells(1, 1)
    ws.Hyperlinks.Add(Anchor=cell, Address=url)
    # TextToDisplay は環境により効かないので、表示文字は明示的にセル値で上書き
    if getattr(args, 'text', None):
        cell.Value = args.text
    print(f"ハイパーリンク追加: {ws.Name}!{cell.Address} → {url}")
    print("（保存はしていません）")
    return True


def cmd_validation(args):
    """入力規則(ドロップダウン): validation <range> --list 'A,B,C' / --clear"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: validation <range> --list 'A,B,C'  /  validation <range> --clear")
        return False
    if _reject_extra_args(rest, 1, "使い方: validation <range> --list 'A,B,C'  /  validation <range> --clear"):
        return False
    xl, wb = get_workbook(target_file)
    # --clear も設定パス（先に既存規則を Delete する）も破壊的なので、
    # シート名だけの指定＝使用範囲全域は明示なしでは拒否する
    whole = _whole_sheet_spec(wb, rest[0])
    if whole is not None and not getattr(args, 'whole_sheet', False):
        print(f"エラー: '{rest[0]}' はシート '{whole}' の使用範囲全域を指します。")
        print(f"  範囲を明示するか、本当に全域なら --whole-sheet を付けてください。")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0])
    if getattr(args, 'clear', False):
        rng.Validation.Delete()
        print(f"入力規則を削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True
    lst = getattr(args, 'list', None)
    if not lst:
        print("--list 'A,B,C' を指定してください。")
        return False
    rng.Validation.Delete()
    rng.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=lst)  # xlValidateList=3
    rng.Validation.InCellDropdown = True
    print(f"入力規則(リスト)を設定: {ws.Name}!{rng.Address}  [{lst}]")
    print("（保存はしていません）")
    return True


def cmd_freeze(args):
    """ウィンドウ枠固定: freeze <cell> / freeze off"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: freeze <cell>（そのセルの左上で固定）  /  freeze off")
        return False
    xl, wb = get_workbook(target_file)
    ws = wb.ActiveSheet
    ws.Activate()
    if rest[0].lower() == 'off':
        xl.ActiveWindow.FreezePanes = False
        print(f"枠固定を解除: {ws.Name}")
    else:
        ws.Range(rest[0]).Select()
        xl.ActiveWindow.FreezePanes = True
        print(f"枠固定: {ws.Name} {rest[0]} の左上で固定")
    print("（保存はしていません）")
    return True


def cmd_comment(args):
    """セルコメント: comment <cell> <text> / comment <cell> --remove"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: comment <cell> <text>  /  comment <cell> --remove")
        return False
    xl, wb = get_workbook(target_file)
    # シート名だけの指定は使用範囲の左上（A1とは限らない）に黙って命中するため拒否
    whole = _whole_sheet_spec(wb, rest[0])
    if whole is not None:
        print(f"エラー: '{rest[0]}' はセルではなくシート '{whole}' を指します。"
              f"セルを明示してください（例: {whole}!A1）")
        return False
    ws, rng = _resolve_range(xl, wb, rest[0])
    cell = rng.Cells(1, 1)
    if getattr(args, 'remove', False):
        cell.ClearComments()
        print(f"コメント削除: {ws.Name}!{cell.Address}")
        print("（保存はしていません）")
        return True
    if len(rest) < 2:
        # 引数不足のときに既存コメントを消さないよう、ClearComments はチェックの後
        print("使い方: comment <cell> <text>")
        return False
    cell.ClearComments()
    cell.AddComment(rest[1])
    print(f"コメント追加: {ws.Name}!{cell.Address}")
    print("（保存はしていません）")
    return True


# ================================================================
# 重量級コマンド (1) チャート
# ================================================================

_XL_CHART_TYPE = {
    'column':  51,     # xlColumnClustered
    'bar':     57,     # xlBarClustered
    'line':    4,      # xlLine
    'pie':     5,      # xlPie
    'scatter': -4169,  # xlXYScatter
    'area':    1,      # xlArea
    'doughnut': -4120, # xlDoughnut
}
_XL_CHART_TYPE_NAME = {v: k for k, v in _XL_CHART_TYPE.items()}


def cmd_chart(args):
    """グラフ操作: chart <create|list|delete> ...

      chart create <data_range> [--type column|bar|line|pie|scatter|area]
                   [--title "見出し"] [--at セル] [--name 名] [--width N --height N]
      chart list
      chart delete <name>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: chart <create|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sh in wb.Worksheets:   # グラフシートは ChartObjects を持たないため除外
            for co in sh.ChartObjects():
                cnt += 1
                try:
                    t = _XL_CHART_TYPE_NAME.get(int(co.Chart.ChartType), co.Chart.ChartType)
                except Exception:
                    t = '?'
                print(f"[{sh.Name}] {co.Name}  type={t}")
        if cnt == 0:
            print("グラフはありません。")
        return True

    if action == 'create':
        if len(rest) < 2:
            print("使い方: chart create <data_range> [--type ...] [--title ...] [--at セル] [--name 名]")
            return False
        ws, rng = _resolve_range(xl, wb, rest[1])

        # 配置: --at 指定があればそのセルの左上、なければデータ範囲の右隣
        at = getattr(args, 'at', None)
        if at:
            anchor = ws.Range(at)
            left, top = anchor.Left, anchor.Top
        else:
            left, top = rng.Left + rng.Width + 10, rng.Top
        width = float(getattr(args, 'width', None) or 360)
        height = float(getattr(args, 'height', None) or 216)

        co = ws.ChartObjects().Add(left, top, width, height)
        ch = co.Chart
        ch.SetSourceData(rng)
        ctype = (getattr(args, 'type', None) or 'column').lower()
        if ctype not in _XL_CHART_TYPE:
            print(f"未知のグラフ種別: {ctype}（{'/'.join(_XL_CHART_TYPE)}）")
            co.Delete()
            return False
        ch.ChartType = _XL_CHART_TYPE[ctype]
        if getattr(args, 'title', None):
            ch.HasTitle = True
            ch.ChartTitle.Text = args.title
        if getattr(args, 'name', None):
            co.Name = args.name
        print(f"グラフ作成: [{ws.Name}] {co.Name}  種別={ctype}  データ={rng.Address}")
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: chart delete <name>"); return False
        name = rest[1]
        for sh in wb.Worksheets:   # グラフシートは ChartObjects を持たないため除外
            for co in sh.ChartObjects():
                if co.Name == name:
                    co.Delete()
                    print(f"グラフ削除: [{sh.Name}] {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: グラフ '{name}' が見つかりません")
        return False

    print(f"未知のアクション: {action}（create|list|delete）")
    return False


_XL_AXIS = {'category': 1, 'value': 2, 'series': 3, 'secondary': 2}   # xlCategory/xlValue/xlSeriesAxis
_XL_LEGEND_POS = {'bottom': -4107, 'corner': 2, 'top': -4160, 'right': -4152, 'left': -4131}
_XL_TRENDLINE = {'linear': -4132, 'exponential': 5, 'logarithmic': -4133,
                 'movingaverage': 6, 'polynomial': 3, 'power': 4}


def cmd_chart_config(args):
    """グラフ詳細設定: chart-config <action> <chart名> ...

      set-source <chart> <range>                          データ範囲を再設定
      set-type <chart> <type>                             種別変更(column/bar/line/pie/...)
      set-title <chart> <text>                            グラフタイトル
      set-axis-title <chart> <category|value|secondary> <text>   軸タイトル
      axis-format <chart> <axis> [format]                 軸の表示形式 get/set
      axis-scale <chart> <axis> [--min N --max N --major N --minor N]  軸目盛
      gridlines <chart> <axis> [--major on|off --minor on|off]        目盛線
      legend <chart> <bottom|top|right|left|corner|off>   凡例
      style <chart> <1-48>                                組込スタイル
      placement <chart> <1|2|3>                           1=移動+サイズ/2=移動のみ/3=自由
      data-labels <chart> [--value --percent --category --series --position 位置]
      add-series <chart> <values_range> [--series-name 名] [--category-range 範囲]
      remove-series <chart> <index>
      series-format <chart> <index> [--marker-style N --marker-size N --marker-fg #.. --marker-bg #.. --invert]
      trendline list <chart> <series_index>
      trendline add  <chart> <series_index> <type> [--name 名]
      trendline delete <chart> <series_index> <trendline_index>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: chart-config <action> <chart名> ...（詳細は --help）")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    def find_chart(name):
        for sh in wb.Worksheets:   # グラフシートは ChartObjects を持たないため除外
            for co in sh.ChartObjects():
                if co.Name == name:
                    return sh, co, co.Chart
        return None, None, None

    # trendline はサブ動詞が rest[1] に来る特例
    if action == 'trendline':
        sub = rest[1].lower() if len(rest) >= 2 else ''
        cname = rest[2] if len(rest) >= 3 else None
        sh, co, ch = find_chart(cname) if cname else (None, None, None)
        if not ch:
            print(f"エラー: グラフ '{cname}' が見つかりません（chart list で確認）。"); return False
        try:
            sidx = int(rest[3]) if len(rest) >= 4 else 1
        except ValueError:
            print("series_index は数値で指定してください。"); return False
        s = ch.SeriesCollection(sidx)
        if sub == 'list':
            tls = s.Trendlines()
            print(f"--- {cname} 系列{sidx} の近似曲線 ({tls.Count}) ---")
            for i in range(1, tls.Count + 1):
                tl = tls.Item(i)
                try:
                    nm = tl.Name
                except Exception:
                    nm = f"#{i}"
                print(f"  [{i}] {nm}")
            if tls.Count == 0:
                print("  (なし)")
            return True
        if sub == 'add':
            ttype = (rest[4] if len(rest) >= 5 else 'linear').lower()
            if ttype not in _XL_TRENDLINE:
                print(f"未知の種別: {ttype}（{'/'.join(_XL_TRENDLINE)}）"); return False
            tl = s.Trendlines().Add(_XL_TRENDLINE[ttype])
            if getattr(args, 'name', None):
                tl.Name = args.name
            print(f"近似曲線追加: {cname} 系列{sidx} {ttype}")
            print("（保存はしていません）"); return True
        if sub == 'delete':
            # 削除系のインデックス省略は「黙って#1が消える」事故のもと。明示必須。
            if len(rest) < 5:
                print("使い方: chart-config trendline delete <chart> <series_index> <trendline_index>")
                print("  （削除対象の trendline_index は省略できません。trendline list で確認）")
                return False
            try:
                tidx = int(rest[4])
            except ValueError:
                print("trendline_index は数値で指定してください。"); return False
            s.Trendlines().Item(tidx).Delete()
            print(f"近似曲線削除: {cname} 系列{sidx} #{tidx}")
            print("（保存はしていません）"); return True
        print("使い方: chart-config trendline <list|add|delete> <chart> <series_index> ...")
        return False

    # それ以外は rest[1] が chart 名
    cname = rest[1] if len(rest) >= 2 else None
    sh, co, ch = find_chart(cname) if cname else (None, None, None)
    if not ch:
        print(f"エラー: グラフ '{cname}' が見つかりません（chart list で確認）。"); return False

    def get_axis(axname):
        a = (axname or 'value').lower()
        if a not in _XL_AXIS:
            return None
        if a == 'secondary':
            return ch.Axes(2, 2)            # xlValue, xlSecondary
        return ch.Axes(_XL_AXIS[a])

    if action == 'set-source':
        if len(rest) < 3:
            print("使い方: chart-config set-source <chart> <range>"); return False
        ws, rng = _resolve_range(xl, wb, rest[2])
        ch.SetSourceData(rng)
        print(f"データ範囲再設定: {cname} ← {rng.Address}")
        print("（保存はしていません）"); return True

    if action == 'set-type':
        t = (rest[2] if len(rest) >= 3 else '').lower()
        if t not in _XL_CHART_TYPE:
            print(f"未知の種別: {t}（{'/'.join(_XL_CHART_TYPE)}）"); return False
        ch.ChartType = _XL_CHART_TYPE[t]
        print(f"種別変更: {cname} → {t}"); print("（保存はしていません）"); return True

    if action == 'set-title':
        txt = rest[2] if len(rest) >= 3 else ''
        ch.HasTitle = True
        ch.ChartTitle.Text = txt
        print(f"タイトル設定: {cname} = {txt}"); print("（保存はしていません）"); return True

    if action == 'set-axis-title':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        txt = rest[3] if len(rest) >= 4 else ''
        ax.HasTitle = True
        ax.AxisTitle.Text = txt
        print(f"軸タイトル設定: {cname} {rest[2]} = {txt}"); print("（保存はしていません）"); return True

    if action == 'axis-format':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        if len(rest) >= 4:
            ax.TickLabels.NumberFormat = rest[3]
            print(f"軸表示形式設定: {cname} {rest[2]} = {rest[3]}")
            print("（保存はしていません）"); return True
        else:
            print(f"軸表示形式: {cname} {rest[2]} = {ax.TickLabels.NumberFormat}"); return True

    if action == 'axis-scale':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        changed = []
        for opt, prop in (('min', 'MinimumScale'), ('max', 'MaximumScale'),
                          ('major', 'MajorUnit'), ('minor', 'MinorUnit')):
            v = getattr(args, opt, None)
            if v is not None:
                setattr(ax, prop, float(v)); changed.append(f"{opt}={v}")
        if changed:
            print(f"軸目盛設定: {cname} {rest[2]} [{', '.join(changed)}]")
            print("（保存はしていません）"); return True
        else:
            print(f"軸目盛: {cname} {rest[2]} min={ax.MinimumScale} max={ax.MaximumScale} "
                  f"major={ax.MajorUnit} minor={ax.MinorUnit}"); return True

    if action == 'gridlines':
        ax = get_axis(rest[2] if len(rest) >= 3 else None)
        if ax is None:
            print("軸は category|value|secondary で指定してください。"); return False
        mj = getattr(args, 'major', None); mn = getattr(args, 'minor', None)
        if mj is None and mn is None:
            print(f"目盛線: {cname} {rest[2]} major={ax.HasMajorGridlines} minor={ax.HasMinorGridlines}")
            return True
        if mj is not None:
            ax.HasMajorGridlines = (mj.lower() == 'on')
        if mn is not None:
            ax.HasMinorGridlines = (mn.lower() == 'on')
        print(f"目盛線設定: {cname} {rest[2]} major={getattr(args,'major',None)} minor={getattr(args,'minor',None)}")
        print("（保存はしていません）"); return True

    if action == 'legend':
        pos = (rest[2] if len(rest) >= 3 else 'bottom').lower()
        if pos == 'off':
            ch.HasLegend = False
            print(f"凡例: {cname} = 非表示")
        else:
            if pos not in _XL_LEGEND_POS:
                print(f"位置は {'/'.join(_XL_LEGEND_POS)}|off で指定してください。"); return False
            ch.HasLegend = True
            ch.Legend.Position = _XL_LEGEND_POS[pos]
            print(f"凡例: {cname} = {pos}")
        print("（保存はしていません）"); return True

    if action == 'style':
        try:
            sid = int(rest[2]) if len(rest) >= 3 else 1
        except ValueError:
            print("使い方: chart-config style <chart> <1-48の数値>"); return False
        ch.ChartStyle = sid
        print(f"スタイル設定: {cname} = {sid}"); print("（保存はしていません）"); return True

    if action == 'placement':
        try:
            pl = int(rest[2]) if len(rest) >= 3 else 1
        except ValueError:
            print("使い方: chart-config placement <chart> <1|2|3>"); return False
        co.Placement = pl
        names = {1: '移動+サイズ', 2: '移動のみ', 3: '自由配置'}
        print(f"配置方法: {cname} = {pl}（{names.get(pl, pl)}）"); print("（保存はしていません）"); return True

    if action == 'data-labels':
        ch.ApplyDataLabels(
            ShowValue=bool(getattr(args, 'value', False)),
            ShowPercentage=bool(getattr(args, 'percent', False)),
            ShowCategoryName=bool(getattr(args, 'category', False)),
            ShowSeriesName=bool(getattr(args, 'series', False)))
        pos = getattr(args, 'position', None)
        if pos:
            posmap = {'center': -4108, 'insideend': 3, 'outsideend': 2, 'bestfit': 5, 'insidebase': 4}
            if pos.lower() not in posmap:
                print(f"⚠ 未知の位置: {pos}（{'/'.join(posmap)}）位置指定はスキップしました。")
            else:
                # 全系列に適用（以前は系列1のみで、複数系列だと部分適用のまま成功表示だった）
                pos_failed = []
                for si in range(1, ch.SeriesCollection().Count + 1):
                    try:
                        ch.SeriesCollection(si).DataLabels().Position = posmap[pos.lower()]
                    except Exception:
                        pos_failed.append(si)
                if pos_failed:
                    print(f"⚠ 位置指定が適用できなかった系列: {pos_failed}"
                          "（グラフ種別によって位置指定不可の場合があります）")
        print(f"データラベル設定: {cname}"); print("（保存はしていません）"); return True

    if action == 'add-series':
        ws, vrng = _resolve_range(xl, wb, rest[2])
        s = ch.SeriesCollection().NewSeries()
        s.Values = vrng
        if getattr(args, 'series_name', None):
            s.Name = args.series_name
        if getattr(args, 'category_range', None):
            _, crng = _resolve_range(xl, wb, args.category_range)
            s.XValues = crng
        print(f"系列追加: {cname} ← {vrng.Address}"); print("（保存はしていません）"); return True

    if action == 'remove-series':
        # 削除系のインデックス省略は「黙って系列1が消える」事故のもと。明示必須。
        if len(rest) < 3:
            print("使い方: chart-config remove-series <chart> <series_index>")
            print("  （削除対象の series_index は省略できません）")
            return False
        try:
            idx = int(rest[2])
        except ValueError:
            print("series_index は数値で指定してください。"); return False
        ch.SeriesCollection(idx).Delete()
        print(f"系列削除: {cname} #{idx}"); print("（保存はしていません）"); return True

    if action == 'series-format':
        idx = int(rest[2]) if len(rest) >= 3 else 1
        s = ch.SeriesCollection(idx)
        ch_list = []
        if getattr(args, 'marker_style', None) is not None:
            s.MarkerStyle = int(args.marker_style); ch_list.append('style')
        if getattr(args, 'marker_size', None) is not None:
            s.MarkerSize = int(args.marker_size); ch_list.append('size')
        if getattr(args, 'marker_fg', None):
            s.MarkerForegroundColor = _hex_to_excel_color(args.marker_fg); ch_list.append('fg')
        if getattr(args, 'marker_bg', None):
            s.MarkerBackgroundColor = _hex_to_excel_color(args.marker_bg); ch_list.append('bg')
        if getattr(args, 'invert', False):
            s.InvertIfNegative = True; ch_list.append('invert')
        print(f"系列書式: {cname} #{idx} [{', '.join(ch_list)}]"); print("（保存はしていません）"); return True

    print(f"未知のアクション: {action}")
    return False


# ================================================================
# 重量級コマンド (2) ピボットテーブル
# ================================================================

_XL_PIVOT_FUNC = {'sum': -4157, 'count': -4112, 'average': -4106,
                  'max': -4136, 'min': -4139}


def _unique_sheet_name(wb, base):
    """重複しないシート名を返す"""
    existing = {sh.Name for sh in wb.Sheets}
    if base not in existing:
        return base
    i = 2
    while f"{base}{i}" in existing:
        i += 1
    return f"{base}{i}"


def cmd_pivot(args):
    """ピボット操作: pivot <create|list|delete> ...

      pivot create <data_range> [--rows F1,F2] [--cols F1] [--values F1,F2]
                   [--func sum|count|average|max|min] [--sheet 出力シート | --at セル] [--name 名]
      pivot list
      pivot delete <name>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: pivot <create|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
            for pt in sh.PivotTables():
                cnt += 1
                print(f"[{sh.Name}] {pt.Name}")
        if cnt == 0:
            print("ピボットテーブルはありません。")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: pivot delete <name>"); return False
        name = rest[1]
        for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
            for pt in sh.PivotTables():
                if pt.Name == name:
                    pt.TableRange2.Clear()
                    print(f"ピボット削除: [{sh.Name}] {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: ピボット '{name}' が見つかりません")
        return False

    if action == 'create':
        if len(rest) < 2:
            print("使い方: pivot create <data_range> [--rows ..][--cols ..][--values ..][--func ..]")
            return False
        ws_s, rng = _resolve_range(xl, wb, rest[1])

        # 出力先の決定: --sheet > --at > 新規シート
        sheet_opt = getattr(args, 'sheet', None)
        at = getattr(args, 'at', None)
        created_sheet = None      # このコマンドが新規に作ったシート（失敗時の後始末用）
        if sheet_opt:
            dws = None
            for sh in wb.Sheets:
                # Excel のシート名重複判定は大文字小文字を区別しない
                if sh.Name.lower() == sheet_opt.lower():
                    dws = sh; break
            if dws is None:
                dws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
                created_sheet = dws
                try:
                    dws.Name = sheet_opt
                except Exception as ex:
                    # 禁止文字(/ 等)・31文字超で失敗すると無名シートが残骸になる
                    print(f"エラー: シート名 '{sheet_opt}' を設定できません: {ex}")
                    xl.DisplayAlerts = False
                    try:
                        dws.Delete()
                    finally:
                        xl.DisplayAlerts = True
                    return False
            dest = dws.Range("A3")
        elif at:
            dest = ws_s.Range(at)
        else:
            dws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
            dws.Name = _unique_sheet_name(wb, "ピボット")
            created_sheet = dws
            dest = dws.Range("A3")

        pt = None
        funcname = (getattr(args, 'func', None) or 'sum').lower()
        values = getattr(args, 'values', None)
        try:
            pc = wb.PivotCaches().Create(1, rng)        # xlDatabase=1
            name = getattr(args, 'name', None)
            if name:
                pt = pc.CreatePivotTable(dest, name)
            else:
                pt = pc.CreatePivotTable(dest)

            def set_fields(spec, orient):
                if not spec:
                    return
                for f in spec.split(','):
                    f = f.strip()
                    if f:
                        pt.PivotFields(f).Orientation = orient

            set_fields(getattr(args, 'rows', None), 1)   # xlRowField
            set_fields(getattr(args, 'cols', None), 2)   # xlColumnField

            func = _XL_PIVOT_FUNC.get(funcname, -4157)
            if values:
                for f in values.split(','):
                    f = f.strip()
                    if f:
                        df = pt.AddDataField(pt.PivotFields(f), f"{funcname}/{f}", func)
        except Exception as ex:
            # 存在しないフィールド名等で途中失敗すると、追加した新規シートと
            # 空ピボットが残骸として残り、再実行のたび「ピボット2/3…」と増殖する。
            # このコマンドが作ったものだけ片づける（既存シートは消さない）
            print(f"エラー: ピボット作成に失敗しました: {ex}")
            print("  --rows/--cols/--values のフィールド名がデータ範囲の見出しと一致しているか確認してください。")
            try:
                if created_sheet is not None:
                    xl.DisplayAlerts = False
                    try:
                        created_sheet.Delete()
                    finally:
                        xl.DisplayAlerts = True
                    print(f"  作成途中のシートは片づけました。")
                elif pt is not None:
                    pt.TableRange2.Clear()
                    print(f"  作成途中の空ピボットは片づけました。")
            except Exception:
                pass
            return False

        print(f"ピボット作成: [{pt.Parent.Name}] {pt.Name}  ソース={ws_s.Name}!{rng.Address}")
        print(f"  行={getattr(args,'rows',None) or '-'}  列={getattr(args,'cols',None) or '-'}  "
              f"値={values or '-'}({funcname})")
        print("（保存はしていません）")
        return True

    print(f"未知のアクション: {action}（create|list|delete）")
    return False


def _find_pivot(wb, name):
    for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
        for pt in sh.PivotTables():
            if pt.Name == name:
                return sh, pt
    return None, None


_XL_PIVOT_ORIENT = {'row': 1, 'col': 2, 'column': 2, 'filter': 3, 'page': 3, 'value': 4, 'data': 4, 'hidden': 0}


def cmd_pivot_field(args):
    """ピボットのフィールド管理: pivot-field <action> <pivot名> <フィールド> ...

      list <pivot>
      add-row|add-col|add-filter <pivot> <field>
      add-value <pivot> <field> [--func sum|count|average|max|min] [--name 表示名]
      remove <pivot> <field>
      set-func <pivot> <field> <func>            データフィールドの集計関数
      set-name <pivot> <field> <表示名>          データフィールドの表示名
      set-format <pivot> <field> <書式コード>    数値書式
      set-filter <pivot> <field> 値1 値2 ...     表示する値を限定
      sort <pivot> <field> <asc|desc>
      group-date <pivot> <field> <days|months|quarters|years>
      group-numeric <pivot> <field> <start> <end> <interval>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: pivot-field <action> <pivot名> <field> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)
    pname = rest[1] if len(rest) >= 2 else None
    sh, pt = _find_pivot(wb, pname) if pname else (None, None)
    if not pt:
        print(f"エラー: ピボット '{pname}' が見つかりません（pivot list で確認）。"); return False

    def pf(field):
        try:
            return pt.PivotFields(field)
        except Exception:
            return None

    if action == 'list':
        print(f"--- {pname} のフィールド ---")
        orient_name = {1: '行', 2: '列', 3: 'フィルタ', 4: '値', 0: '未配置'}
        for i in range(1, pt.PivotFields().Count + 1):
            f = pt.PivotFields().Item(i)
            try:
                o = int(f.Orientation)
            except Exception:
                o = 0
            print(f"  {f.Name}: {orient_name.get(o, o)}")
        return True

    field = rest[2] if len(rest) >= 3 else None
    if action in ('add-row', 'add-col', 'add-filter'):
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Orientation = {'add-row': 1, 'add-col': 2, 'add-filter': 3}[action]
        print(f"フィールド配置: {pname}[{field}] = {action[4:]}"); print("（保存はしていません）"); return True

    if action == 'add-value':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        fn = (getattr(args, 'func', None) or 'sum').lower()
        func = _XL_PIVOT_FUNC.get(fn, -4157)
        cname = getattr(args, 'name', None) or f"{fn}/{field}"
        pt.AddDataField(p, cname, func)
        print(f"値フィールド追加: {pname}[{field}] ({fn}) 表示名={cname}"); print("（保存はしていません）"); return True

    if action == 'remove':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Orientation = 0     # xlHidden
        print(f"フィールド除外: {pname}[{field}]"); print("（保存はしていません）"); return True

    if action == 'set-func':
        fn = (rest[3] if len(rest) >= 4 else 'sum').lower()
        if fn not in _XL_PIVOT_FUNC:
            print(f"未知の関数: {fn}（{'/'.join(_XL_PIVOT_FUNC)}）"); return False
        # データフィールドは表示名で参照されるため DataFields を走査
        target = None
        for i in range(1, pt.DataFields.Count + 1):
            d = pt.DataFields.Item(i)
            if d.Name == field or d.SourceName == field:
                target = d; break
        if target is None:
            print(f"エラー: 値フィールド '{field}' が見つかりません。"); return False
        target.Function = _XL_PIVOT_FUNC[fn]
        print(f"集計関数変更: {pname}[{field}] = {fn}"); print("（保存はしていません）"); return True

    if action == 'set-name':
        newname = rest[3] if len(rest) >= 4 else None
        if not newname:
            print("使い方: pivot-field set-name <pivot> <field> <新しい表示名>")
            return False
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Caption = newname
        print(f"表示名変更: {pname}[{field}] → {newname}"); print("（保存はしていません）"); return True

    if action == 'set-format':
        code = rest[3] if len(rest) >= 4 else None
        if not code:
            print('使い方: pivot-field set-format <pivot> <field> <書式コード（例: "#,##0"）>')
            return False
        for i in range(1, pt.DataFields.Count + 1):
            d = pt.DataFields.Item(i)
            if d.Name == field or d.SourceName == field:
                d.NumberFormat = code
                print(f"値フィールド書式: {pname}[{field}] = {code}"); print("（保存はしていません）"); return True
        print(f"エラー: 値フィールド '{field}' が見つかりません。"); return False

    if action == 'set-filter':
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        wanted = set(rest[3:])
        if not wanted:
            print("表示する値を1つ以上指定してください。"); return False
        # 指定値が実在するか先に照合（全部タイポだと Excel が「全項目非表示」を拒否し、
        # 実状態と成功メッセージが食い違うため）
        item_names = []
        for i in range(1, p.PivotItems().Count + 1):
            item_names.append(p.PivotItems().Item(i).Name)
        missing = wanted - set(item_names)
        if missing:
            print(f"エラー: 存在しない値が指定されています: {sorted(missing)}")
            print(f"  このフィールドの値: {item_names}")
            return False
        failed = []
        for i in range(1, p.PivotItems().Count + 1):
            it = p.PivotItems().Item(i)
            try:
                it.Visible = (it.Name in wanted)
            except Exception:
                failed.append(it.Name)
        if failed:
            print(f"⚠ 一部の項目の表示切替に失敗しました: {failed}")
            print("  （Excel の制約: 全項目非表示は不可、など。実際の表示状態を確認してください）")
        print(f"値フィルタ: {pname}[{field}] = {sorted(wanted)}"); print("（保存はしていません）"); return True

    if action == 'sort':
        order = (rest[3] if len(rest) >= 4 else 'asc').lower()
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.AutoSort(2 if order.startswith('d') else 1, p.Name)
        print(f"並べ替え: {pname}[{field}] = {'降順' if order.startswith('d') else '昇順'}")
        print("（保存はしていません）"); return True

    if action == 'group-date':
        interval = (rest[3] if len(rest) >= 4 else 'months').lower()
        # Periods: [秒,分,時,日,月,四半期,年]
        flags = {'days': 3, 'months': 4, 'quarters': 5, 'years': 6}
        if interval not in flags:
            print("interval は days|months|quarters|years"); return False
        periods = [False] * 7
        periods[flags[interval]] = True
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.DataRange.Cells(1, 1).Group(Periods=periods)
        print(f"日付グループ化: {pname}[{field}] = {interval}"); print("（保存はしていません）"); return True

    if action == 'group-numeric':
        if len(rest) < 6:
            print("使い方: pivot-field group-numeric <pivot> <field> <start> <end> <interval>"); return False
        start, end, step = float(rest[3]), float(rest[4]), float(rest[5])
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.DataRange.Cells(1, 1).Group(Start=start, End=end, By=step)
        print(f"数値グループ化: {pname}[{field}] = {start}〜{end} 刻み{step}"); print("（保存はしていません）"); return True

    print(f"未知のアクション: {action}")
    return False


def cmd_pivot_calc(args):
    """ピボットの計算フィールド・レイアウト: pivot-calc <action> <pivot名> ...

      get-data <pivot>                                出力範囲の値を表示
      calc-field create <pivot> <名前> <数式>         計算フィールド作成（=Revenue-Cost 等）
      calc-field list <pivot>
      calc-field delete <pivot> <名前>
      layout <pivot> <compact|tabular|outline>        レポートレイアウト
      subtotals <pivot> <field> <on|off>              小計の表示
      grand-totals <pivot> <rows|cols|both> <on|off>  総計の表示
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: pivot-calc <action> <pivot名> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'calc-field':
        sub = rest[1].lower() if len(rest) >= 2 else ''
        pname = rest[2] if len(rest) >= 3 else None
        sh, pt = _find_pivot(wb, pname) if pname else (None, None)
        if not pt:
            print(f"エラー: ピボット '{pname}' が見つかりません。"); return False
        if sub == 'create':
            if len(rest) < 5:
                print("使い方: pivot-calc calc-field create <pivot> <名前> <数式>"); return False
            cf_name, formula = rest[3], rest[4]
            pt.CalculatedFields().Add(cf_name, formula)
            print(f"計算フィールド作成: {pname}[{cf_name}] = {formula}")
            print("  （値に表示するには pivot-field add-value で追加）")
            print("（保存はしていません）"); return True
        if sub == 'list':
            cfs = pt.CalculatedFields()
            print(f"--- {pname} の計算フィールド ({cfs.Count}) ---")
            for i in range(1, cfs.Count + 1):
                f = cfs.Item(i)
                try:
                    formula = f.Formula
                except Exception:
                    formula = ''
                print(f"  {f.Name} = {formula}")
            if cfs.Count == 0:
                print("  (なし)")
            return True
        if sub == 'delete':
            cf_name = rest[3] if len(rest) >= 4 else None
            if not cf_name:
                print("使い方: pivot-calc calc-field delete <pivot> <計算フィールド名>")
                return False
            try:
                pt.PivotFields(cf_name).Delete()
            except Exception:
                names = []
                try:
                    cfs = pt.CalculatedFields()
                    names = [cfs.Item(i).Name for i in range(1, cfs.Count + 1)]
                except Exception:
                    pass
                print(f"エラー: 計算フィールド '{cf_name}' を削除できません。")
                print("  存在する計算フィールド: " + (", ".join(names) or "(なし)"))
                return False
            print(f"計算フィールド削除: {pname}[{cf_name}]"); print("（保存はしていません）"); return True
        print("使い方: pivot-calc calc-field <create|list|delete> <pivot> ...")
        return False

    pname = rest[1] if len(rest) >= 2 else None
    sh, pt = _find_pivot(wb, pname) if pname else (None, None)
    if not pt:
        print(f"エラー: ピボット '{pname}' が見つかりません。"); return False

    if action == 'get-data':
        rng = pt.TableRange2
        print(f"ピボット出力範囲: {pt.Parent.Name}!{rng.Address}")
        data = rng.Value
        if data is not None:
            for row in data:
                cells = [('' if c is None else str(c)) for c in (row if isinstance(row, tuple) else [row])]
                print("  " + " | ".join(cells))
        return True

    if action == 'layout':
        lay = (rest[2] if len(rest) >= 3 else 'compact').lower()
        laymap = {'compact': 0, 'tabular': 1, 'outline': 2}   # xlCompactRow/xlTabularRow/xlOutlineRow
        if lay not in laymap:
            print("layout は compact|tabular|outline"); return False
        pt.RowAxisLayout(laymap[lay])
        print(f"レイアウト: {pname} = {lay}"); print("（保存はしていません）"); return True

    if action == 'subtotals':
        field = rest[2] if len(rest) >= 3 else None
        onoff = (rest[3] if len(rest) >= 4 else 'on').lower()
        try:
            p = pt.PivotFields(field)
        except Exception:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Subtotals = tuple([onoff == 'on'] + [False] * 11)   # 先頭=自動小計
        print(f"小計: {pname}[{field}] = {onoff}"); print("（保存はしていません）"); return True

    if action == 'grand-totals':
        which = (rest[2] if len(rest) >= 3 else 'both').lower()
        onoff = (rest[3] if len(rest) >= 4 else 'on').lower()
        val = (onoff == 'on')
        if which in ('rows', 'both'):
            pt.RowGrand = val
        if which in ('cols', 'both'):
            pt.ColumnGrand = val
        print(f"総計: {pname} {which} = {onoff}"); print("（保存はしていません）"); return True

    print(f"未知のアクション: {action}")
    return False


# ================================================================
# 重量級コマンド (3) スライサー
# ================================================================

def _find_pivot_or_table(wb, name):
    """名前からピボット or テーブル(ListObject)を探す。戻り値 (obj, kind, sheet) """
    for sh in wb.Worksheets:   # グラフシートは PivotTables を持たないため除外
        for pt in sh.PivotTables():
            if pt.Name == name:
                return pt, 'pivot', sh
    for sh in wb.Worksheets:   # グラフシートは ListObjects を持たないため除外
        for lo in sh.ListObjects:
            if lo.Name == name:
                return lo, 'table', sh
    return None, None, None


def cmd_slicer(args):
    """スライサー操作: slicer <add|list|delete> ...

      slicer add <pivot名 or テーブル名> <フィールド> [--at セル] [--name 名]
      slicer list
      slicer delete <name>
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: slicer <add|list|delete> ...")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        cnt = 0
        for sc in wb.SlicerCaches:
            for sl in sc.Slicers:
                cnt += 1
                # Slicer.Parent は SlicerCache を返す実装があるため、シート名は Shape 経由で取る
                try:
                    sheet_name = sl.Shape.Parent.Name
                except Exception:
                    try:
                        sheet_name = sl.Parent.Name
                    except Exception:
                        sheet_name = '?'
                print(f"{sl.Name}  (フィールド={sc.SourceName}, シート={sheet_name})")
        if cnt == 0:
            print("スライサーはありません。")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: slicer delete <name>"); return False
        name = rest[1]
        for sc in wb.SlicerCaches:
            for sl in sc.Slicers:
                if sl.Name == name:
                    sl.Delete()
                    print(f"スライサー削除: {name}")
                    print("（保存はしていません）")
                    return True
        print(f"エラー: スライサー '{name}' が見つかりません")
        return False

    if action == 'add':
        if len(rest) < 3:
            print("使い方: slicer add <pivot名 or テーブル名> <フィールド> [--at セル] [--name 名]")
            return False
        src_name, field = rest[1], rest[2]
        src, kind, sh = _find_pivot_or_table(wb, src_name)
        if src is None:
            print(f"エラー: ピボット/テーブル '{src_name}' が見つかりません")
            return False

        # 配置先の座標を先に解決する（Add2 の後に --at が不正で例外になると、
        # スライサー本体のない SlicerCache だけが孤児としてブックに残るため）
        at = getattr(args, 'at', None)
        dws = sh
        if at:
            anchor = dws.Range(at)
            top, left = anchor.Top, anchor.Left
        else:
            top, left = 10.0, 400.0

        sc = wb.SlicerCaches.Add2(src, field)
        sl = sc.Slicers.Add(SlicerDestination=dws, Caption=field,
                            Top=top, Left=left, Width=144.0, Height=180.0)
        # Slicers.Add の Name 引数は効かないことがあるので作成後に明示セット
        req_name = getattr(args, 'name', None)
        if req_name:
            eff_name = req_name.replace(' ', '')
            if eff_name != req_name:
                print(f"⚠ スライサー名のスペースは使えないため除去しました: '{req_name}' → '{eff_name}'")
            try:
                sl.Name = eff_name
            except Exception as ex:
                print(f"⚠ 名前 '{eff_name}' を設定できませんでした（{ex}）。自動名のままです。")
        print(f"スライサー追加: {sl.Name}  ソース={src_name}({kind})  フィールド={field}  シート={dws.Name}")
        print("（保存はしていません）")
        return True

    print(f"未知のアクション: {action}（add|list|delete）")
    return False


# ================================================================
# 計算モード (大量書き込みの高速化)
# ================================================================

def cmd_calc_mode(args):
    """計算モードの確認・切替・再計算

      calc-mode                 現在のモードを表示
      calc-mode manual          手動計算に（大量書込の前に）
      calc-mode auto            自動計算に戻す
      calc-mode recalc          今すぐ再計算（手動中の一括計算）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)

    names = {-4105: '自動 (automatic)', -4135: '手動 (manual)', 2: '半自動 (semiautomatic)'}

    if not rest:
        m = xl.Calculation
        print(f"現在の計算モード: {names.get(m, m)}")
        return True

    sub = rest[0].lower()
    if sub in ('auto', 'automatic'):
        xl.Calculation = -4105
        print("計算モード → 自動")
    elif sub == 'manual':
        xl.Calculation = -4135
        print("計算モード → 手動（書込後は calc-mode recalc / auto で再計算）")
    elif sub in ('recalc', 'now', 'calculate'):
        xl.Calculate()
        print("再計算しました")
    else:
        print(f"未知の指定: {sub}（manual|auto|recalc）")
        return False
    return True


# ================================================================
# 重量級コマンド (4) PowerQuery （一覧・更新・作成・M式書換・削除・読み込み配線）
# ================================================================

def cmd_powerquery(args):
    """PowerQuery: powerquery <list|refresh> ...

      powerquery list                 クエリと接続の一覧
      powerquery refresh              全クエリ/接続を更新 (RefreshAll)
      powerquery refresh <name>       指定クエリ/接続を更新
      powerquery add <name>           M式から新規クエリ作成（接続のみ）
                                      M式は --m-file / _last_query.m / --m "..."
      powerquery delete <name>        クエリを削除
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: powerquery <list|refresh [name]|add <name>|delete <name>>")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        # クエリ一覧
        try:
            qs = wb.Queries
            qcount = qs.Count
        except Exception:
            qs = None
            qcount = 0
        if qs and qcount > 0:
            print(f"--- PowerQuery クエリ ({qcount}) ---")
            for i in range(1, qcount + 1):
                q = qs.Item(i)
                try:
                    desc = q.Description or ''
                except Exception:
                    desc = ''
                # M式の行数を補助表示
                try:
                    nlines = len(str(q.Formula).replace('\r\n', '\n').split('\n'))
                except Exception:
                    nlines = '?'
                print(f"  {q.Name}  (M式 {nlines}行)" + (f"  - {desc}" if desc else ""))
        else:
            print("PowerQuery クエリはありません。")
        # 接続一覧（更新対象の確認用）
        try:
            conns = wb.Connections
            ccount = conns.Count
        except Exception:
            ccount = 0
        if ccount > 0:
            print(f"--- 接続 ({ccount}) ---")
            for cn in conns:
                print(f"  {cn.Name}")
        return True

    if action == 'refresh':
        if len(rest) >= 2:
            name = rest[1]
            target_conn = None
            for cn in wb.Connections:
                if cn.Name == name or cn.Name == f"Query - {name}":
                    target_conn = cn
                    break
            if target_conn:
                target_conn.Refresh()
                print(f"更新しました: {target_conn.Name}")
                return True
            # 接続が無い（読み込みなしクエリ等）
            print(f"接続 '{name}' が見つかりません。")
            print("  （読み込みなしクエリは更新対象がありません。powerquery list で名前を確認）")
            return False
        else:
            wb.RefreshAll()
            print("全クエリ/接続を更新しました (RefreshAll)")
            print("  ※ バックグラウンド更新の場合、完了まで数秒かかることがあります。")
            return True

    if action == 'add':
        if len(rest) < 2:
            print("使い方: powerquery add <name> [--m-file f | --m \"M式\"]")
            return False
        name = rest[1]
        # M式の取得: --m インライン > --m-file > _last_query.m
        m_inline = getattr(args, 'm_opt', None)
        if m_inline:
            formula = m_inline
        else:
            mf = getattr(args, 'm_file', None)
            path = smart_path_resolve(mf) if mf else _LAST_QUERY_FILE
            if not path or not os.path.exists(path):
                print(f"エラー: M式ファイルが見つかりません: {mf or _LAST_QUERY_FILE}")
                print("  _last_query.m にM式を書くか、--m-file / --m を指定してください。")
                return False
            formula = read_code_file(path)
        if not formula or not formula.strip():
            print("エラー: M式が空です。")
            return False
        # 重複チェック
        try:
            existing = [wb.Queries.Item(i).Name for i in range(1, wb.Queries.Count + 1)]
        except Exception:
            existing = []
        if name in existing:
            print(f"エラー: クエリ '{name}' は既に存在します（delete してから add）。")
            return False
        desc = getattr(args, 'desc', None) or ''
        wb.Queries.Add(name, formula, desc)
        print(f"クエリ作成: {name}（接続のみ。シート/モデルへの読み込みは別途）")
        print("（保存はしていません）")
        return True

    if action == 'edit':
        if len(rest) < 2:
            print("使い方: powerquery edit <name> [--m-file f | --m \"M式\"]")
            return False
        name = rest[1]
        # M式の取得（add と同じ）: --m > --m-file > _last_query.m
        m_inline = getattr(args, 'm_opt', None)
        if m_inline:
            formula = m_inline
        else:
            mf = getattr(args, 'm_file', None)
            path = smart_path_resolve(mf) if mf else _LAST_QUERY_FILE
            if not path or not os.path.exists(path):
                print(f"エラー: M式ファイルが見つかりません: {mf or _LAST_QUERY_FILE}")
                return False
            formula = read_code_file(path)
        if not formula or not formula.strip():
            print("エラー: M式が空です。")
            return False
        try:
            cnt = wb.Queries.Count
        except Exception:
            cnt = 0
        for i in range(1, cnt + 1):
            q = wb.Queries.Item(i)
            if q.Name == name:
                q.Formula = formula                # WorkbookQuery.Formula は書込可（検証済）
                print(f"クエリ書き換え: {name}")
                print("（保存はしていません。反映には powerquery refresh が必要なことがあります）")
                return True
        print(f"エラー: クエリ '{name}' が見つかりません")
        return False

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: powerquery delete <name>"); return False
        name = rest[1]
        try:
            cnt = wb.Queries.Count
        except Exception:
            cnt = 0
        for i in range(1, cnt + 1):
            if wb.Queries.Item(i).Name == name:
                wb.Queries.Item(i).Delete()
                print(f"クエリ削除: {name}")
                print("（保存はしていません）")
                return True
        print(f"エラー: クエリ '{name}' が見つかりません")
        return False

    if action == 'load':
        # 接続のみクエリを「シートのテーブル」または「データモデル」に読み込む配線。
        #   powerquery load <name> --to sheet  [--sheet S] [--at A1]
        #   powerquery load <name> --to model
        if len(rest) < 2:
            print('使い方: powerquery load <name> --to sheet|model [--sheet S] [--at A1]')
            return False
        name = rest[1]
        to = (getattr(args, 'to', None) or 'sheet').lower()
        # クエリ存在チェック
        try:
            existing = [wb.Queries.Item(i).Name for i in range(1, wb.Queries.Count + 1)]
        except Exception:
            existing = []
        if name not in existing:
            print(f"エラー: クエリ '{name}' が見つかりません（powerquery list で確認）。")
            return False

        # Power Query (Mashup) の OLEDB 接続文字列 — 記録マクロが生成する形に合わせる
        conn_str = ("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
                    f'Location={name};Extended Properties=""')
        cmd_text = f"SELECT * FROM [{name}]"

        if to == 'sheet':
            # 出力先シート（--sheet 省略時はアクティブシート）
            sheet_name = getattr(args, 'sheet', None)
            ws = None
            created_ws = None
            if sheet_name:
                for i in range(1, wb.Worksheets.Count + 1):
                    # Excel のシート名重複判定は大文字小文字を区別しない
                    if wb.Worksheets.Item(i).Name.lower() == sheet_name.lower():
                        ws = wb.Worksheets.Item(i); break
                if ws is None:
                    ws = wb.Worksheets.Add()
                    created_ws = ws
                    try:
                        ws.Name = sheet_name
                    except Exception as ex:
                        # 禁止文字(/ 等)・31文字超で失敗すると無名シートが残骸になる
                        print(f"エラー: シート名 '{sheet_name}' を設定できません: {ex}")
                        xl.DisplayAlerts = False
                        try:
                            ws.Delete()
                        finally:
                            xl.DisplayAlerts = True
                        return False
            else:
                ws = wb.ActiveSheet
            at = getattr(args, 'at', None) or 'A1'
            dest = ws.Range(at)
            # 0 = xlSrcExternal。Source に Mashup の OLEDB 文字列を渡す
            lo = ws.ListObjects.Add(0, conn_str, None, True, dest)
            qt = lo.QueryTable
            qt.CommandType = 2                 # xlCmdSql
            qt.CommandText = cmd_text
            qt.RowNumbers = False
            qt.FillAdjacentFormulas = False
            qt.PreserveFormatting = True
            qt.RefreshOnFileOpen = False
            qt.BackgroundQuery = False
            qt.AdjustColumnWidth = True
            try:
                qt.Refresh(False)              # BackgroundQuery:=False
            except Exception as ex:
                # M式の実行時エラー等で失敗すると、追加済みの ListObject と
                # 自動生成の接続が孤児として残り、再実行のたび「テーブル1/2…」と
                # 増殖する。このコマンドが作ったものだけ片づける
                print(f"エラー: クエリの読み込みに失敗しました: {ex}")
                print("  M式の実行時エラーの可能性があります（powerquery get で確認）。")
                try:
                    wbconn = qt.WorkbookConnection
                except Exception:
                    wbconn = None
                try:
                    lo.Delete()
                except Exception:
                    pass
                try:
                    if wbconn is not None:
                        wbconn.Delete()
                except Exception:
                    pass
                if created_ws is not None:
                    xl.DisplayAlerts = False
                    try:
                        created_ws.Delete()
                    except Exception:
                        pass
                    finally:
                        xl.DisplayAlerts = True
                print("  追加途中のテーブル・接続は片づけました。")
                return False
            try:
                lo.Name = name
            except Exception:
                pass
            # 既定では「接続」等の汎用名が付く。refresh <name> で引けるよう
            # Excel 標準の "Query - <name>" に揃える。
            try:
                wbconn = qt.WorkbookConnection
                if wbconn is not None:
                    wbconn.Name = f"Query - {name}"
            except Exception:
                pass
            print(f"シートに読み込みました: {name} → {ws.Name}!{at}（テーブル: {lo.Name}）")
            print("（保存はしていません）")
            return True

        if to == 'model':
            # データモデル（Power Pivot）へ。Queries.Add が作る "Query - name"
            # 接続が残っていると衝突するので、あれば作り直す。
            # ただしその接続がシートのテーブル（--to sheet の読み込み）に使われている
            # 場合、削除するとシート側の更新配線が壊れるため停止する。
            cn_name = f"Query - {name}"
            for cn in list(wb.Connections):
                if cn.Name == cn_name:
                    used_by = None
                    try:
                        for ws_chk in wb.Worksheets:
                            for lo_chk in ws_chk.ListObjects:
                                try:
                                    if lo_chk.QueryTable.WorkbookConnection.Name == cn_name:
                                        used_by = f"{ws_chk.Name}!{lo_chk.Name}"
                                        break
                                except Exception:
                                    continue
                            if used_by:
                                break
                    except Exception:
                        pass
                    if used_by:
                        print(f"エラー: 接続 '{cn_name}' はシートのテーブル {used_by} が使用中です。")
                        print("  削除するとテーブルの更新ができなくなるため中止しました。")
                        print("  モデルにも読み込みたい場合は、シート読み込みを解除してから実行してください。")
                        return False
                    try:
                        cn.Delete()
                    except Exception:
                        pass
            # Connections.Add2(Name, Description, ConnectionString, CommandText,
            #                  lCmdtype, CreateModelConnection, ImportRelationships)
            # モデル読込は記録マクロ形式に合わせる: CommandText=クエリ名,
            # lCmdtype=6 (xlCmdTableCollection)。これでモデルテーブル名が
            # クエリ名になる（SQL/SELECT形式だと "クエリ" の汎用名になる）。
            wb.Connections.Add2(cn_name, "", conn_str, name, 6, True, False)
            print(f"データモデルに読み込みました: {name}")
            print("（保存はしていません。datamodel list で確認できます）")
            return True

        print(f"未知の読み込み先: {to}（sheet|model）")
        return False

    print(f"未知のアクション: {action}（list|refresh|add|edit|delete|load）")
    return False


# ================================================================
# 重量級コマンド (5) コネクション / データモデル （管理・読み取り）
# ================================================================

# XlConnectionType: xlConnectionTypeOLEDB=1, ODBC=2, XMLMAP=3, TEXT=4, WEB=5,
#                   DATAFEED=6, MODEL=7, WORKSHEET=8, NOSOURCE=9
_XL_CONN_TYPE = {1: 'OLEDB', 2: 'ODBC', 3: 'XMLMAP', 4: 'TEXT',
                 5: 'WEB', 6: 'DATAFEED', 7: 'MODEL', 8: 'WORKSHEET', 9: 'NOSOURCE'}


def cmd_connection(args):
    """ブック接続の管理: connection <list|refresh|delete> [name]

      connection list                クエリ/外部データ接続の一覧（種別・接続文字列）
      connection refresh [name]      接続を更新（name 省略で全件 RefreshAll）
      connection delete <name>       接続を削除
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: connection <list|refresh|delete> [name]")
        return False
    action = rest[0].lower()
    xl, wb = get_workbook(target_file)

    if action == 'list':
        conns = wb.Connections
        n = conns.Count
        if n == 0:
            print("接続はありません。")
            return True
        print(f"--- ブック接続 ({n}) ---")
        for cn in conns:
            try:
                t = _XL_CONN_TYPE.get(int(cn.Type), cn.Type)
            except Exception:
                t = '?'
            print(f"  {cn.Name}  [{t}]")
            try:
                if cn.Description:
                    print(f"      説明: {cn.Description}")
            except Exception:
                pass
            # 接続文字列・コマンド（OLEDB/ODBC）
            try:
                sub = None
                if int(cn.Type) == 1:
                    sub = cn.OLEDBConnection
                elif int(cn.Type) == 2:
                    sub = cn.ODBCConnection
                if sub is not None:
                    cs = str(sub.Connection)
                    print(f"      接続: {cs[:100]}{'…' if len(cs) > 100 else ''}")
            except Exception:
                pass
        return True

    if action == 'refresh':
        if len(rest) >= 2:
            name = rest[1]
            for cn in wb.Connections:
                if cn.Name == name or cn.Name == f"Query - {name}":
                    cn.Refresh()
                    print(f"更新しました: {cn.Name}")
                    return True
            print(f"エラー: 接続 '{name}' が見つかりません")
            return False
        wb.RefreshAll()
        print("全接続を更新しました (RefreshAll)")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: connection delete <name>"); return False
        name = rest[1]
        for cn in wb.Connections:
            if cn.Name == name or cn.Name == f"Query - {name}":
                actual = cn.Name           # Delete 後は参照不可になるので退避
                cn.Delete()
                print(f"接続を削除: {actual}")
                print("（保存はしていません）")
                return True
        print(f"エラー: 接続 '{name}' が見つかりません")
        return False

    print(f"未知のアクション: {action}（list|refresh|delete）")
    return False


def cmd_datamodel(args):
    """データモデル: datamodel <list|relation|measure>

      datamodel list   モデルのテーブル・リレーションシップ・メジャーを一覧
      datamodel relation add    <FKテーブル> <FK列> <PKテーブル> <PK列>   リレーション作成
      datamodel relation delete <FKテーブル> <FK列> <PKテーブル> <PK列>   リレーション削除
      datamodel measure add <テーブル> <メジャー名> --dax "式" [--format general|whole|decimal|currency|percent|scientific]
                                                                 [--decimals N] [--thousands] [--symbol JPY]   メジャー(DAX)作成
      datamodel measure delete <メジャー名>                              メジャー削除
      （※ テーブルの追加は powerquery load --to model）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    action = rest[0].lower() if rest else 'list'
    xl, wb = get_workbook(target_file)

    try:
        model = wb.Model
    except Exception:
        print("このブックはデータモデルに対応していません。")
        # list 系は「無い」を正常報告でよいが、追加/削除の要求は実行されて
        # いないので失敗として返す（batch やスクリプト連携で握りつぶさない）
        return action in ('list', 'tables', 'measures', 'relations')

    # --- リレーションシップの作成・削除 ---
    #   datamodel relation add    <FKテーブル> <FK列> <PKテーブル> <PK列>
    #   datamodel relation delete <FKテーブル> <FK列> <PKテーブル> <PK列>
    if action in ('relation', 'rel', 'relationship'):
        sub = rest[1].lower() if len(rest) >= 2 else ''
        if sub in ('add', 'delete'):
            if len(rest) < 6:
                print(f'使い方: datamodel relation {sub} <FKテーブル> <FK列> <PKテーブル> <PK列>')
                print('  FK=多側(参照する側) / PK=一側(参照される側)')
                return False
            fkt_name, fkc_name, pkt_name, pkc_name = rest[2], rest[3], rest[4], rest[5]
        if sub == 'add':
            try:
                fkt = model.ModelTables.Item(fkt_name)
                pkt = model.ModelTables.Item(pkt_name)
            except Exception:
                print(f"エラー: テーブルが見つかりません（{fkt_name} / {pkt_name}）。datamodel list で確認。")
                return False
            try:
                fkc = fkt.ModelTableColumns.Item(fkc_name)
                pkc = pkt.ModelTableColumns.Item(pkc_name)
            except Exception:
                print(f"エラー: 列が見つかりません（{fkt_name}[{fkc_name}] / {pkt_name}[{pkc_name}]）。")
                return False
            model.ModelRelationships.Add(fkc, pkc)
            print(f"リレーション作成: {fkt_name}[{fkc_name}] → {pkt_name}[{pkc_name}]")
            print("（保存はしていません）")
            return True
        if sub == 'delete':
            rels = model.ModelRelationships
            for i in range(1, rels.Count + 1):
                r = rels.Item(i)
                try:
                    if (r.ForeignKeyTable.Name == fkt_name and r.ForeignKeyColumn.Name == fkc_name
                            and r.PrimaryKeyTable.Name == pkt_name and r.PrimaryKeyColumn.Name == pkc_name):
                        r.Delete()
                        print(f"リレーション削除: {fkt_name}[{fkc_name}] → {pkt_name}[{pkc_name}]")
                        print("（保存はしていません）")
                        return True
                except Exception:
                    continue
            print("エラー: 該当するリレーションが見つかりません（datamodel list で確認）。")
            return False
        print('使い方: datamodel relation <add|delete> <FKテーブル> <FK列> <PKテーブル> <PK列>')
        return False

    # --- メジャー(DAX)の作成・削除 ---
    #   datamodel measure add <テーブル> <メジャー名> [--dax "式" | --dax-file f | _last_dax.dax]
    #   datamodel measure delete <メジャー名>
    if action in ('measure', 'measures'):
        sub = rest[1].lower() if len(rest) >= 2 else ''
        if sub == 'add':
            if len(rest) < 4:
                print('使い方: datamodel measure add <テーブル> <メジャー名> --dax "DAX式"')
                print('  DAX は --dax / --dax-file / _last_dax.dax(UTF-8) から取得。')
                print('  ※ 先頭の = は不要。日本語テーブル名は DAX 内でシングルクォート: SUM(\'売上\'[数量])')
                return False
            tbl_name, measure_name = rest[2], rest[3]
            # DAX の取得: --dax インライン > --dax-file > _last_dax.dax
            dax = getattr(args, 'dax', None)
            if not dax:
                df = getattr(args, 'dax_file', None)
                path = smart_path_resolve(df) if df else _LAST_DAX_FILE
                if not path or not os.path.exists(path):
                    print(f"エラー: DAXファイルが見つかりません: {df or _LAST_DAX_FILE}")
                    print("  _last_dax.dax に式を書くか、--dax / --dax-file を指定してください。")
                    return False
                dax = read_code_file(path)
            if not dax or not dax.strip():
                print("エラー: DAX式が空です。")
                return False
            dax = dax.strip()
            if dax.startswith('='):            # Excel数式の癖で = を付けても通るように
                dax = dax[1:].strip()
            try:
                tbl = model.ModelTables.Item(tbl_name)
            except Exception:
                print(f"エラー: テーブル '{tbl_name}' が見つかりません。datamodel list で確認。")
                return False
            # 数値書式（既定 general）。引数付き書式は GetModelFormat* メソッドで取得
            #   （ModelFormat* プロパティは既定値専用で引数を渡せないため）。
            fmt_name = (getattr(args, 'format', None) or 'general').lower()
            dec_arg = getattr(args, 'decimals', None)
            try:
                decimals = int(dec_arg) if dec_arg is not None else 2
            except (TypeError, ValueError):
                print(f"エラー: --decimals は数値で指定してください: '{dec_arg}'")
                return False
            thousands = bool(getattr(args, 'thousands', False))
            symbol = getattr(args, 'symbol', None) or ''
            try:
                if fmt_name == 'general':
                    fmt = model.ModelFormatGeneral
                elif fmt_name in ('whole', 'wholenumber'):
                    fmt = model.GetModelFormatWholeNumber(thousands)
                elif fmt_name in ('decimal', 'decimalnumber'):
                    fmt = model.GetModelFormatDecimalNumber(thousands, decimals)
                elif fmt_name == 'currency':
                    # Symbol は通貨コード（USD/JPY/EUR 等）。グリフ（$ 等）は無効だが
                    # GetModelFormatCurrency では落ちず Add 時に例外になるため、
                    # フォールバックは Add 側で行う。
                    fmt = model.GetModelFormatCurrency(symbol, decimals)
                elif fmt_name in ('percent', 'percentage'):
                    fmt = model.GetModelFormatPercentageNumber(thousands, decimals)
                elif fmt_name in ('scientific', 'sci'):
                    fmt = model.GetModelFormatScientificNumber(decimals)
                else:
                    print(f"エラー: 未知の書式 '{fmt_name}'（general|whole|decimal|currency|percent|scientific）")
                    return False
            except Exception as e:
                print(f"エラー: 書式オブジェクトの取得に失敗: {str(e)[:120]}")
                return False
            desc = getattr(args, 'desc', None) or ''
            try:
                model.ModelMeasures.Add(measure_name, tbl, dax, fmt, desc)
            except Exception as e:
                # currency でグリフ等の無効な通貨コードだと Add 時に例外。
                # 既定の通貨記号で 1 回だけ再試行する。
                if fmt_name == 'currency' and symbol:
                    try:
                        model.ModelMeasures.Add(measure_name, tbl, dax,
                                                model.GetModelFormatCurrency('', decimals), desc)
                        print(f"  注意: 通貨コード '{symbol}' は無効。既定の通貨記号で作成しました（有効例: USD, JPY, EUR）。")
                        print(f"メジャー作成: {tbl_name}[{measure_name}] = {dax}  (書式=currency)")
                        print("（保存はしていません）")
                        return True
                    except Exception as e2:
                        e = e2
                print(f"エラー: メジャー作成に失敗しました: {str(e)[:200]}")
                print("  DAX 構文・テーブル/列名・シングルクォートを確認してください。")
                return False
            print(f"メジャー作成: {tbl_name}[{measure_name}] = {dax}  (書式={fmt_name})")
            print("（保存はしていません）")
            return True
        if sub == 'delete':
            if len(rest) < 3:
                print('使い方: datamodel measure delete <メジャー名>')
                return False
            measure_name = rest[2]
            ms = model.ModelMeasures
            for i in range(1, ms.Count + 1):
                if ms.Item(i).Name == measure_name:
                    ms.Item(i).Delete()
                    print(f"メジャー削除: {measure_name}")
                    print("（保存はしていません）")
                    return True
            print(f"エラー: メジャー '{measure_name}' が見つかりません（datamodel list で確認）。")
            return False
        print('使い方: datamodel measure <add|delete> ...')
        return False

    if action != 'list':
        print(f"未知のアクション: {action}（list|relation|measure）")
        return False

    # テーブル
    try:
        mts = model.ModelTables
        tn = mts.Count
    except Exception:
        mts = None
        tn = 0
    print(f"--- データモデル: テーブル ({tn}) ---")
    for i in range(1, tn + 1):
        mt = mts.Item(i)
        try:
            rc = mt.RecordCount
        except Exception:
            rc = '?'
        print(f"  {mt.Name}  ({rc}行)")
    if tn == 0:
        print("  (なし)")

    # リレーションシップ
    try:
        rels = model.ModelRelationships
        rn = rels.Count
    except Exception:
        rels = None
        rn = 0
    print(f"--- リレーションシップ ({rn}) ---")
    for i in range(1, rn + 1):
        r = rels.Item(i)
        try:
            fkt = r.ForeignKeyTable.Name
            fkc = r.ForeignKeyColumn.Name
            pkt = r.PrimaryKeyTable.Name
            pkc = r.PrimaryKeyColumn.Name
            active = ''
            try:
                active = '' if r.Active else '  (無効)'
            except Exception:
                pass
            print(f"  {fkt}[{fkc}] → {pkt}[{pkc}]{active}")
        except Exception:
            print(f"  (リレーション {i}: 読み取り不可)")
    if rn == 0:
        print("  (なし)")

    # メジャー（対応バージョンのみ）
    try:
        ms = model.ModelMeasures
        mn = ms.Count
        print(f"--- メジャー ({mn}) ---")
        for i in range(1, mn + 1):
            me = ms.Item(i)
            try:
                tbl = me.AssociatedTable.Name
            except Exception:
                tbl = '?'
            print(f"  {me.Name}  (所属={tbl})")
        if mn == 0:
            print("  (なし)")
    except Exception:
        pass

    return True


# ================================================================
# エントリポイント
# ================================================================

def build_parser():
    """argparse の構築（main と batch で共用）"""
    parser = argparse.ArgumentParser(
        description="VBAマネージャー (アクティブブック対応版)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python vba_manager.py list                                      # アクティブブックのマクロ一覧
  python vba_manager.py list 秀.xlsm                              # 指定ファイルのマクロ一覧
  python vba_manager.py list-modules                              # モジュール一覧
  python vba_manager.py get 空白行の削除                          # プロシージャ取得 → _last_proc.vba に保存
  python vba_manager.py get shu001 空白行の削除                   # モジュール指定してプロシージャ取得
  python vba_manager.py get アクティブマクロフォーム.CommandButton2_Click  # ドット区切りでも可
  python vba_manager.py replace-procedure                         # _last_proc.vba の内容で置換
  python vba_manager.py replace-procedure --code-file my.vba
  python vba_manager.py replace-module shu001 shu001_new.bas
  python vba_manager.py export-module shu001                      # shu001.bas にエクスポート
""")

    sub = parser.add_subparsers(dest="command")

    # diag
    sub.add_parser("diag")

    # setup-check（導入セルフ診断・初心者が最初に打つ1コマンド）
    p = sub.add_parser("setup-check", help="導入セルフ診断（Python/pywin32/Excel/VBOM信頼設定を○×表示）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # list-open
    p = sub.add_parser("list-open")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # list [excel_file]
    p = sub.add_parser("list")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--standard", action="store_true", help="標準モジュールのみを抽出")
    p.add_argument("--detail", action="store_true", help="所属モジュール・行数・先頭コメント付きで表示")
    p.add_argument("--module", dest="module_opt", default=None, help="対象モジュールを限定")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--personal", action="store_true", help="個人用マクロブック (PERSONAL.XLSB) を対象にする")
    p.add_argument("--addin", nargs="?", const=True, default=False,
                   help="アドインブック (.xlam/.xla) を対象にする。複数ロード時は名前(一部可)を指定")
    p.add_argument("--all", action="store_true", help="開いているすべてのブック・アドインを対象にする")

    # list-modules [excel_file]
    p = sub.add_parser("list-modules")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--personal", action="store_true", help="個人用マクロブック (PERSONAL.XLSB) を対象にする")
    p.add_argument("--addin", nargs="?", const=True, default=False,
                   help="アドインブック (.xlam/.xla) を対象にする。複数ロード時は名前(一部可)を指定")
    p.add_argument("--all", action="store_true", help="開いているすべてのブック・アドインを対象にする")

    # get [excel_file] <macro_name> [...]
    p = sub.add_parser("get")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="保存先ファイル（省略時は _last_proc.vba。参照用コピーを残したいときに）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # replace-procedure [excel_file] [code_file] [--code-file file] [--module name]
    p = sub.add_parser("replace-procedure")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--code-file", dest="code_file_opt", default=None)
    p.add_argument("--module", dest="module_opt", default=None,
                   help="適用先モジュール名を指定（同名プロシージャが複数ある場合に必須）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes",
                   help="確認プロンプトをスキップして自動で置換を実行します")
    p.add_argument("--force", action="store_true", dest="force",
                   help="構文エラー警告を無視して強制適用します")

    # add-procedure [excel_file] <module_name> [--code-file f] [-y]
    p = sub.add_parser("add-procedure", help="新規プロシージャをモジュール末尾に追加（コードは _last_proc.vba から）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--code-file", dest="code_file_opt", default=None)
    p.add_argument("-y", "--yes", action="store_true", dest="yes",
                   help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force",
                   help="構文エラー警告・バックアップ失敗を無視して強行")

    # add-module [excel_file] <module_name> [--type std|class|form]
    p = sub.add_parser("add-module", help="新規モジュールを追加（標準/クラス/フォーム）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--type", dest="type_opt", default="std",
                   help="モジュール種別 std|class|form（既定 std）")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # delete-procedure [excel_file] <macro_name> [--module name] [-y]
    p = sub.add_parser("delete-procedure", help="プロシージャを削除（削除コードを表示して確認）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module_opt", default=None,
                   help="対象モジュール名（同名が複数ある場合に必須）")
    p.add_argument("-y", "--yes", action="store_true", dest="yes",
                   help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # docs [excel_file] [--out f.md]
    p = sub.add_parser("docs", help="ブックの構成ドキュメント（取説）を自動生成")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力Markdownパス（省略時は _last_docs.md）")
    p.add_argument("--preview", dest="preview", default=None,
                   help="各シートの先頭N行をMarkdown表で含める")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # checkup(健康診断) [excel_file] [--out f.md]
    p = sub.add_parser("checkup", aliases=["健康診断"],
                       help="ブックの健康診断レポート（総合判定+壊れた参照+シート検査+前回との比較）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力Markdownパス（省略時は _last_checkup.md）")
    p.add_argument("--all", dest="all_books", action="store_true",
                   help="開いている全ブックを一括診断（PERSONAL.XLSB含む）")
    p.add_argument("--history", action="store_true",
                   help="診断せず過去の診断履歴（経過観察）を表で表示")
    p.add_argument("--detail", action="store_true",
                   help="--history で各回の間に起きた所見/マクロの増減も表示")
    p.add_argument("--note", default=None,
                   help="今回の診断にカルテのメモを添付（例: --note \"ボタン18を一語修正\"）")
    p.add_argument("--strict", action="store_true",
                   help="所見が1件でもあれば終了コード1（自動化のゲート用。既定は診断完了=0）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--ack-all", dest="ack_all", action="store_true",
                   help="今回のコード/フォーム所見を全て確認済み（意図的）として登録し、"
                        "以降の所見サマリ・総合判定から除外する"
                        "（存在しないマクロ呼び出し等の致命的所見は対象外）")
    p.add_argument("--show-ack", dest="show_ack", action="store_true",
                   help="診断はせず、確認済み（意図的）所見の一覧を表示")
    p.add_argument("--unack", default=None, metavar="文字列",
                   help="部分一致する確認済み所見を確認済みから外す（見直したくなった時）")

    # call-graph [excel_file] [--macro 名]
    p = sub.add_parser("call-graph", help="マクロの呼び出し関係を解析（未解決Call＝一語バグ検出つき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--macro", dest="macro_opt", default=None,
                   help="このマクロを起点に呼び出しツリーを展開")
    p.add_argument("--mermaid", nargs="?", const="_DEFAULT_", default=None,
                   help="Mermaid図をMarkdownに出力（省略時 _last_callgraph.md。GitHub/Qiitaで描画可）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # impact(影響範囲) [excel_file] <マクロ名>
    p = sub.add_parser("impact", aliases=["影響範囲"],
                       help="マクロ修正前の影響範囲予告（呼び元/呼び先を間接まで一覧）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # grep [excel_file] <pattern>
    p = sub.add_parser("grep", help="全モジュール横断のVBAコード検索")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--regex", action="store_true", help="正規表現として検索")
    p.add_argument("-i", "--ignore-case", dest="ignore_case", action="store_true",
                   help="大文字小文字を区別しない")
    p.add_argument("--module", dest="module_opt", default=None, help="検索対象モジュールを限定")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="表示件数の上限（既定200）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # code-replace [excel_file] <検索> <置換>
    p = sub.add_parser("code-replace", help="全マクロ横断の一括置換（diffプレビュー・バックアップ・確認つき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--regex", action="store_true", help="正規表現として置換")
    p.add_argument("--module", dest="module_opt", default=None, help="対象モジュールを限定")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")

    # list-backups [キーワード] / restore <バックアップファイル>
    p = sub.add_parser("list-backups", help="backups のバックアップ一覧（COM不要）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--max", dest="max_hits", type=int, default=None, help="表示件数の上限（既定30）")
    p = sub.add_parser("restore", help="モジュールバックアップ(.bas/.frm)を開いているブックへ書き戻す")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # list-shortcuts [excel_file]
    p = sub.add_parser("list-shortcuts")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # run-macro [excel_file] <macro_name> [args...]
    p = sub.add_parser("run-macro", help="Excel内の指定されたマクロを実行します")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--json", action="store_true", help="実行結果をJSON形式で出力")
    p.add_argument("--auto-dialog", dest="auto_dialog", default=None,
                   help="実行中に出るMsgBox/InputBoxを自動応答 ok|cancel|yes|no（既定は応答しない）")

    # test [excel_file] [絞り込み] [--module 名] [--auto-dialog ok] [--json]
    p = sub.add_parser("test", aliases=["テスト"],
                       help="テストSub（名前が「テスト」/test で始まる引数なしSub）を一括実行して成否一覧")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--module", dest="module", default=None, help="対象モジュールを限定")
    p.add_argument("--json", action="store_true", help="結果をJSON形式でも出力")
    p.add_argument("--auto-dialog", dest="auto_dialog", default=None,
                   help="テスト中に出るMsgBox等を自動応答 ok|cancel|yes|no")

    # replace-module [excel_file] <module_name> <bas_file>
    p = sub.add_parser("replace-module")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--force", action="store_true", dest="force",
                   help="バックアップ失敗時も強行する")

    # delete-module [excel_file] <module_name> [-y]
    p = sub.add_parser("delete-module", help="モジュール丸ごと削除（要約表示→確認→バックアップつき）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("-y", "--yes", action="store_true", dest="yes", help="確認プロンプトをスキップ")
    p.add_argument("--force", action="store_true", dest="force", help="バックアップ失敗時も強行する")

    # export-module [excel_file] <module_name>
    p = sub.add_parser("export-module")
    p.add_argument("posargs", nargs="+")

    # export-all [excel_file] [--dir 出力先] [--check]
    p = sub.add_parser("export-all", help="全モジュールを一括エクスポート（1接続）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--dir", dest="dir_opt", default=None, help="出力先フォルダ（省略時はSCRIPTS）")
    p.add_argument("--check", action="store_true", help="書き出した各ファイルに check-bas 相当の検査をかける")

    # reorder-macro <macro_name> <up|down>
    p = sub.add_parser("reorder-macro")
    p.add_argument("posargs", nargs="+")
    p.add_argument("--force", action="store_true",
                   help="バックアップが取れなくても実行する（未保存の新規ブック等）")

    # --- 目コマンド ---
    # read-range [excel_file] [range ...] [--formula] [--tsv [f]] [--width N] [--json]
    p = sub.add_parser("read-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--formula", action="store_true",
                   help="計算結果でなく数式(.Formula)を表示する")
    p.add_argument("--tsv", dest="tsv_out", nargs="?", const="_DEFAULT_", default=None,
                   help="TSVに書き出す（省略時 _last_values.tsv。編集して write-range で書き戻す往復用）")
    p.add_argument("--width", dest="width", default=None,
                   help="列の最大表示幅（既定40。超えた分は…付きで切り詰め）")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定。'!'入り・記号入りシート名向け）")

    # read-selection [excel_file] [--formula]
    p = sub.add_parser("read-selection")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--formula", action="store_true",
                   help="計算結果でなく数式(.Formula)を表示する")

    # sheet-info [excel_file] [--preview N]
    p = sub.add_parser("sheet-info")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--preview", dest="preview", default=None,
                   help="各シート使用範囲の先頭N行も表示（ブック俯瞰・1接続）")

    # snapshot [excel_file] [sheet] [--out file] [--sheet NAME] [--max-rows N]
    p = sub.add_parser("snapshot",
                       help="ブック(または1シート)を意味構造JSONに畳む＝開いたままブックLMの下ごしらえ")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力JSONパス（省略時は _last_snapshot.json）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（1シートだけ畳む。posargでも可）")
    p.add_argument("--max-rows", dest="max_rows", default=None,
                   help="1シートあたりのセル読み込み行上限（既定5000。超過は打ち切り注記）")

    # snapshot-diff <before.json> [after.json] [--max N]
    p = sub.add_parser("snapshot-diff",
                       help="2つのsnapshot JSONを比較＝前後差分の検分（COM不要）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--max", dest="max_opt", default=None,
                   help="1分類あたりの表示件数上限（既定20。超過は件数のみ表示）")

    # wiring [excel_file] [--json]
    p = sub.add_parser("wiring", aliases=["配線図"],
                       help="ボタン⇔マクロの配線図（OnAction一覧＋行き先のないボタン検出）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true",
                   help="JSON形式で出力（機械処理用）")

    # screenshot [excel_file] [range] [--out file]
    p = sub.add_parser("screenshot")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--out", dest="out_opt", default=None,
                   help="出力PNGパス（省略時は _last_view.png）")

    # --- 手コマンド (シートの編集・整形・構造操作) ---
    # write-range [excel_file] <range> [値] [--tsv file]
    p = sub.add_parser("write-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--tsv", dest="tsv_opt", default=None,
                   help="グリッドを読み込むTSVファイル（省略時は _last_values.tsv）")
    p.add_argument("--raw", action="store_true",
                   help="数値変換せず文字列として書き込む（セル書式を文字列にする。'007'等の先頭ゼロ保持）")
    p.add_argument("--append", action="store_true",
                   help="使用範囲の最終行の次に書く（rangeは「シート名!列文字」。ログ追記用）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # clear-range [excel_file] <range> [--contents|--formats|--all]
    p = sub.add_parser("clear-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--contents", action="store_true", help="値のみクリア")
    p.add_argument("--formats", action="store_true", help="書式のみクリア")
    p.add_argument("--all", action="store_true", help="すべてクリア（既定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # format-range [excel_file] <range> [書式オプション...]
    p = sub.add_parser("format-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--font")
    p.add_argument("--size")
    p.add_argument("--bold", action="store_true")
    p.add_argument("--unbold", action="store_true")
    p.add_argument("--italic", action="store_true")
    p.add_argument("--color")
    p.add_argument("--bg")
    p.add_argument("--number-format", dest="number_format")
    p.add_argument("--align", choices=['left', 'center', 'right', 'fill', 'justify'])
    p.add_argument("--valign", choices=['top', 'center', 'bottom'])
    p.add_argument("--wrap", action="store_true")
    p.add_argument("--border", choices=['thin', 'medium', 'thick', 'hairline', 'none'])
    p.add_argument("--col-width", dest="col_width")
    p.add_argument("--row-height", dest="row_height")
    p.add_argument("--merge", action="store_true")
    p.add_argument("--unmerge", action="store_true")
    p.add_argument("--autofit", action="store_true")
    p.add_argument("--lock", action="store_true", help="セルをロック（sheet protect 時に有効）")
    p.add_argument("--unlock", action="store_true", help="セルのロック解除（保護中も編集可に）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")

    # sheet [excel_file] <add|delete|rename|copy|activate|show|hide> ...
    p = sub.add_parser("sheet")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--after")
    p.add_argument("--before")
    p.add_argument("--clear", action="store_true", help="tab-color のクリア")
    p.add_argument("--password", default=None, help="protect/unprotect のパスワード")

    # table [excel_file] <create|list|delete> ...
    p = sub.add_parser("table")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--no-headers", dest="no_headers", action="store_true")
    p.add_argument("--at", dest="at", default=None, help="column add の挿入位置(1始まり)")
    p.add_argument("--desc", dest="desc", action="store_true", help="sort を降順に")
    p.add_argument("--tsv", dest="tsv_out", nargs="?", const="_DEFAULT_", default=None,
                   help="table read の結果をTSVに書き出す（省略時 _last_values.tsv）")

    # name [excel_file] <add|list|delete> ...
    p = sub.add_parser("name")
    p.add_argument("posargs", nargs="*")

    # --- 手コマンド 第2弾 ---
    # a. 編集の足回り
    p = sub.add_parser("row")          # row <insert|delete> <行番号> [本数]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet", default=None,
                   help="対象シート名（省略時はアクティブシート）")
    p = sub.add_parser("col")          # col <insert|delete> <列文字> [本数]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet", default=None,
                   help="対象シート名（省略時はアクティブシート）")
    p = sub.add_parser("copy-range")   # copy-range <src> <dst> [--values]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--values", action="store_true", help="値のみ貼り付け")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="コピー元シート名（srcと分離指定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="コピー元にシート名だけ（使用範囲全域）を許可する")
    p = sub.add_parser("fill")         # fill <range> [--right]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--right", action="store_true", help="右方向にフィル（既定は下）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p = sub.add_parser("sort")         # sort <range> [--key 列] [--desc] [--header|--no-header]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--key", help="並べ替えキー列（列文字）")
    p.add_argument("--desc", action="store_true", help="降順")
    p.add_argument("--header", action="store_true", help="先頭行を見出しとして扱う")
    p.add_argument("--no-header", dest="no_header", action="store_true", help="見出しなし")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート名だけの指定（使用範囲全域）を許可する")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")
    p = sub.add_parser("autofilter")   # autofilter [range] [--off]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--off", action="store_true", help="オートフィルタを解除")

    # b. 検索・置換
    p = sub.add_parser("find")         # find <文字> [--book] [--whole] [--formula]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--book", action="store_true", help="全シート横断で検索")
    p.add_argument("--whole", action="store_true", help="完全一致")
    p.add_argument("--formula", action="store_true", help="数式も検索対象にする")
    p.add_argument("--max", dest="max_hits", type=int, default=None,
                   help="表示件数の上限（既定200）")
    p = sub.add_parser("find-replace") # find-replace <検索> <置換> [range] [--whole]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--whole", action="store_true", help="完全一致のみ置換")
    p.add_argument("--match-case", dest="match_case", action="store_true",
                   help="大文字小文字を区別する（既定は区別しない）")
    p.add_argument("--wildcard", action="store_true",
                   help="検索文字列の * ? をワイルドカードとして扱う（既定は文字どおり）")
    p.add_argument("--sheet", dest="sheet_opt", default=None,
                   help="対象シート名（rangeと分離指定）")

    # c. 保存・印刷まわり
    p = sub.add_parser("save")         # save [excel_file]
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("save-as")      # save-as [excel_file] <path> [--overwrite]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--overwrite", action="store_true", help="出力先が既存でも上書きする")
    p = sub.add_parser("export-pdf")   # export-pdf <出力.pdf> [--sheet|--range]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--sheet", dest="sheet_opt", default=None, help="このシートだけをPDF化")
    p.add_argument("--range", dest="range_opt", default=None, help='この範囲だけをPDF化（例 "集計!A1:H50"）')
    p.add_argument("--overwrite", action="store_true", help="出力先が既存でも上書きする")
    p = sub.add_parser("print-setup")  # print-setup [opts]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--area", help="印刷範囲（例 A1:H50）")
    p.add_argument("--title-rows", dest="title_rows", help="印刷タイトル行（例 1:3）")
    p.add_argument("--title-cols", dest="title_cols", help="印刷タイトル列（例 A:B）")
    p.add_argument("--landscape", action="store_true", help="横向き")
    p.add_argument("--portrait", action="store_true", help="縦向き")
    p.add_argument("--fit-wide", dest="fit_wide", help="横N ページに収める")
    p.add_argument("--fit-tall", dest="fit_tall", help="縦N ページに収める")
    p.add_argument("--zoom", help="拡大縮小率(%%)")
    p.add_argument("--center-h", dest="center_h", action="store_true", help="水平中央")
    p.add_argument("--center-v", dest="center_v", action="store_true", help="垂直中央")

    # d. 仕上げ・見た目
    p = sub.add_parser("cond-format")  # cond-format <range> --gt 100 --bg '#...'
    p.add_argument("posargs", nargs="*")
    p.add_argument("--gt"); p.add_argument("--lt")
    p.add_argument("--ge"); p.add_argument("--le")
    p.add_argument("--eq"); p.add_argument("--ne")
    p.add_argument("--between", nargs=2, metavar=("V1", "V2"))
    p.add_argument("--formula", dest="formula_opt", default=None,
                   help='数式ベースのルール（例 --formula "=B2>AVERAGE($B$2:$B$20)"）')
    p.add_argument("--bg"); p.add_argument("--color")
    p.add_argument("--bold", action="store_true")
    p.add_argument("--clear", action="store_true", help="条件付き書式を全削除")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート全域（シート名だけの指定）を明示的に許可する")
    p = sub.add_parser("hyperlink")    # hyperlink <cell> <url> [--text t] / --remove / --list
    p.add_argument("posargs", nargs="*")
    p.add_argument("--text", help="表示文字")
    p.add_argument("--remove", action="store_true", help="ハイパーリンク削除")
    p.add_argument("--list", dest="list_links", nargs="?", const="__ACTIVE__", default=None,
                   help="シート内の全ハイパーリンクを一覧（--list シート名 で対象指定）")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="--remove でシート全域（シート名だけの指定）を明示的に許可する")
    p = sub.add_parser("validation")   # validation <range> --list 'A,B,C' / --clear
    p.add_argument("posargs", nargs="*")
    p.add_argument("--list", help="ドロップダウン候補（カンマ区切り）")
    p.add_argument("--clear", action="store_true", help="入力規則を削除")
    p.add_argument("--whole-sheet", dest="whole_sheet", action="store_true",
                   help="シート全域（シート名だけの指定）を明示的に許可する")
    p = sub.add_parser("freeze")       # freeze <cell> / freeze off
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("comment")      # comment <cell> <text> / --remove
    p.add_argument("posargs", nargs="*")
    p.add_argument("--remove", action="store_true", help="コメント削除")

    # 重量級(1) chart <create|list|delete>
    p = sub.add_parser("chart")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--type", help="column|bar|line|pie|scatter|area|doughnut")
    p.add_argument("--title", help="グラフタイトル")
    p.add_argument("--at", help="左上を合わせるセル")
    p.add_argument("--name", help="グラフ名")
    p.add_argument("--width", help="幅(pt)")
    p.add_argument("--height", help="高さ(pt)")

    # 重量級(1b) chart-config <action> <chart名> ...
    p = sub.add_parser("chart-config")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--min"); p.add_argument("--max")
    p.add_argument("--major"); p.add_argument("--minor")
    p.add_argument("--value", action="store_true"); p.add_argument("--percent", action="store_true")
    p.add_argument("--category", action="store_true"); p.add_argument("--series", action="store_true")
    p.add_argument("--position")
    p.add_argument("--series-name", dest="series_name")
    p.add_argument("--category-range", dest="category_range")
    p.add_argument("--marker-style", dest="marker_style")
    p.add_argument("--marker-size", dest="marker_size")
    p.add_argument("--marker-fg", dest="marker_fg")
    p.add_argument("--marker-bg", dest="marker_bg")
    p.add_argument("--invert", action="store_true")
    p.add_argument("--name")

    # 重量級(2) pivot <create|list|delete>
    p = sub.add_parser("pivot")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--rows", help="行フィールド（カンマ区切り）")
    p.add_argument("--cols", help="列フィールド（カンマ区切り）")
    p.add_argument("--values", help="値フィールド（カンマ区切り）")
    p.add_argument("--func", help="集計方法 sum|count|average|max|min（既定 sum）")
    p.add_argument("--sheet", help="出力シート名（無ければ作成）")
    p.add_argument("--at", help="出力先セル（同シート内に置く場合）")
    p.add_argument("--name", help="ピボットテーブル名")

    # 重量級(2b) pivot-field <action> <pivot> <field> ...
    p = sub.add_parser("pivot-field")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--func", help="add-value/set-func の集計 sum|count|average|max|min")
    p.add_argument("--name", help="add-value の表示名")

    # 重量級(2c) pivot-calc <action> <pivot> ...
    p = sub.add_parser("pivot-calc")
    p.add_argument("posargs", nargs="*")

    # 重量級(3) slicer <add|list|delete>
    p = sub.add_parser("slicer")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--at", help="左上を合わせるセル")
    p.add_argument("--name", help="スライサー名")

    # calc-mode [manual|auto|recalc]
    p = sub.add_parser("calc-mode")
    p.add_argument("posargs", nargs="*")

    # 重量級(4) powerquery <list|refresh|add|delete>
    p = sub.add_parser("powerquery")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--m-file", dest="m_file", default=None, help="add 用 M式ファイル（省略時 _last_query.m）")
    p.add_argument("--m", dest="m_opt", default=None, help="add 用 M式をインライン指定")
    p.add_argument("--desc", default=None, help="クエリの説明")
    p.add_argument("--to", dest="to", default=None, help="load 用 読み込み先: sheet|model")
    p.add_argument("--sheet", dest="sheet", default=None, help="load --to sheet の出力先シート（省略時アクティブ）")
    p.add_argument("--at", dest="at", default=None, help="load --to sheet の左上セル（省略時 A1）")

    # 重量級(5) connection <list|refresh|delete> / datamodel [list]
    p = sub.add_parser("connection")
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("datamodel")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--dax", dest="dax", default=None, help="measure add 用 DAX式をインライン指定")
    p.add_argument("--dax-file", dest="dax_file", default=None, help="measure add 用 DAXファイル（省略時 _last_dax.dax）")
    p.add_argument("--desc", dest="desc", default=None, help="measure の説明")
    p.add_argument("--format", dest="format", default=None,
                   help="measure の書式: general|whole|decimal|currency|percent|scientific（既定 general）")
    p.add_argument("--decimals", dest="decimals", default=None, help="小数桁数（decimal/currency/percent/scientific、既定2）")
    p.add_argument("--thousands", dest="thousands", action="store_true", help="桁区切りを使う（whole/decimal/percent）")
    p.add_argument("--symbol", dest="symbol", default=None, help="通貨コード（currency、例: USD/JPY/EUR。グリフ$¥は不可。無効なら既定）")

    # printer-list
    p = sub.add_parser("printer-list")
    p.add_argument("posargs", nargs="*")

    # printer-setup
    p = sub.add_parser("printer-setup")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--printer", help="対象プリンター名（省略時はActivePrinterまたは既定プリンター）")
    p.add_argument("--duplex", choices=['simplex', 'vertical', 'horizontal'], help="両面印刷（simplex:片面, vertical:長辺, horizontal:短辺）")
    p.add_argument("--color", choices=['mono', 'color'], help="カラーモード（mono:モノクロ, color:カラー）")
    p.add_argument("--orientation", choices=['portrait', 'landscape'], help="用紙の向き（portrait:縦, landscape:横）")

    # check [excel_file]
    p = sub.add_parser("check", help="全モジュールの構文チェックと診断を実行します")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # check-bas <file.bas> [--fix]  (COM不要・取り込み前の単体検査)
    p = sub.add_parser("check-bas", help="取り込み前に .bas を単体検査（文字コード/改行二重化/重複）。COM不要・複数可")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--fix", action="store_true", help="改行二重化を CP932 のまま自動修正する")
    p.add_argument("--json", action="store_true", help="結果をJSON形式で出力")

    # batch <コマンドファイル|->  （1接続・1プロセスでコマンド列を実行）
    p = sub.add_parser("batch", help="コマンド列を1回の接続で連続実行（ファイル or 標準入力 '-'）")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--keep-going", dest="keep_going", action="store_true",
                   help="途中の失敗で止まらず最後まで実行する")

    # shell（対話セッション。接続を張ったままコマンドを打ち続ける）
    sub.add_parser("shell", help="対話セッション（接続維持のREPL。2コマンド目から再接続なし）")

    return parser


def cmd_batch(args):
    """コマンド列を1プロセス・1COM接続で連続実行: batch <file|->

    各行は通常のCLI引数列そのもの（例: `get shu003 空白行の削除`）。
    空行と # 始まりは無視。get_workbook の接続キャッシュにより全行が同じ
    COM接続を使い回すため、「1コマンド毎の再接続で数分」級の一括作業が
    数秒に縮む。各行の実行は既存コマンドの機械的な再生のみ（判断はしない）。
    """
    import shlex
    src = args.posargs[0] if args.posargs else None
    if not src:
        print("使い方: batch <コマンドファイル|->   （- で標準入力から読む）")
        print("  例: get shu003 マクロA")
        print("      replace-procedure -y")
        return False
    if src == '-':
        # パイプ経由の入力はロケール(CP932)で誤読されるため UTF-8 に固定する（shell と同じ対処）
        if not sys.stdin.isatty():
            try:
                sys.stdin.reconfigure(encoding='utf-8')
            except Exception:
                pass
        text = sys.stdin.read()
    else:
        path = smart_path_resolve(src)
        if not path or not os.path.exists(path):
            print(f"エラー: コマンドファイルが見つかりません: {src}")
            return False
        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                text = f.read()
        except UnicodeDecodeError:
            # メモ帳等で CP932 (Shift-JIS) 保存されたコマンドファイルも受け付ける
            with open(path, 'r', encoding='cp932') as f:
                text = f.read()

    parser = build_parser()
    table = _command_table()
    keep_going = getattr(args, 'keep_going', False)
    total = ok_n = 0
    for lineno, line in enumerate(text.splitlines(), 1):
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        total += 1
        print(f"----- [batch:{lineno}] {line} -----")
        try:
            # Windows パスの \ をエスケープ扱いしない（クォートは通常どおり効く）
            lex = shlex.shlex(line, posix=True)
            lex.whitespace_split = True
            lex.escape = ''
            # shlex 既定のコメント文字 '#' を無効化。行頭 # は上で処理済みで、
            # 行中の # を生かすと「テスト#1」「--bg #FF0000」の # 以降が黙って消える
            lex.commenters = ''
            tokens = list(lex)
        except ValueError as e:
            print(f"[batch:{lineno}] 引数の解析に失敗: {e}")
            if keep_going:
                continue
            print("[batch] 停止（--keep-going で続行可）")
            return False
        try:
            sub_args, unknown = parser.parse_known_args(tokens)
        except SystemExit as e:
            if e.code in (0, None):
                # 行内の -h/--help はヘルプ表示済み。エラーではない
                ok_n += 1
                continue
            print(f"[batch:{lineno}] 引数エラー")
            if keep_going:
                continue
            print("[batch] 停止（--keep-going で続行可）")
            return False
        unknown = [u for u in unknown if u not in ("--visible", "-v")]
        if unknown:
            print(f"[batch:{lineno}] 不明な引数/オプション: {' '.join(unknown)}")
            if keep_going:
                continue
            print("[batch] 停止（--keep-going で続行可）")
            return False
        if not sub_args.command or sub_args.command in ('batch', 'shell'):
            print(f"[batch:{lineno}] このコマンドは batch 内で実行できません")
            if keep_going:
                continue
            return False
        try:
            res = table[sub_args.command](sub_args)
        except SystemExit as e:
            # reorder-macro 等は sys.exit で終了コードを返すため、ここで吸収する
            # （code=None の素の sys.exit() は正常終了。shell 側の判定と揃える）
            res = (e.code in (0, None))
        except Exception as e:
            print(f"[batch:{lineno}] エラー: {e}")
            res = False
        if res is not False:
            ok_n += 1
        elif not keep_going:
            print(f"[batch] {lineno}行目で失敗したため停止（--keep-going で続行可）")
            print(f"===== batch 結果: {ok_n}/{total} 成功 =====")
            return False
    print(f"===== batch 完了: {ok_n}/{total} 成功 =====")
    return ok_n == total


def cmd_shell(args):
    """対話セッション: shell

    接続を張ったままコマンドを打ち続ける REPL。batch のファイル版に対する対話版で、
    get_workbook の接続キャッシュにより2コマンド目からは COM 再接続なしで動く
    （1コマンド約1秒 → 体感即応）。exit / quit / Ctrl+C で終了。
    """
    import shlex
    # パイプ/リダイレクト経由の入力はロケール(CP932)で誤読されるため UTF-8 に固定する
    # （対話（コンソール直打ち）は Windows のコンソールAPIが処理するので触らない）
    if not sys.stdin.isatty():
        try:
            sys.stdin.reconfigure(encoding='utf-8')
        except Exception:
            pass
    parser = build_parser()
    table = _command_table()
    hist_cmds = []
    print("===== vba_manager 対話セッション =====")
    print("  コマンドをそのまま入力（例: list / get マクロ名 / read-range A1:D10）")
    print("  help で使い方、history で履歴（!番号 で再実行・!! は直前）、exit で終了。")
    while True:
        try:
            line = input("vba> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n終了します。")
            break
        if not line or line.startswith('#'):
            continue
        if line.lower() in ('exit', 'quit', 'q'):
            print("終了します。")
            break
        if line.lower() == 'help':
            parser.print_help()
            continue
        if line.lower() == 'history':
            for i, h in enumerate(hist_cmds, 1):
                print(f"  {i}: {h}")
            if not hist_cmds:
                print("  (まだ履歴はありません)")
            continue
        if line.startswith('!'):
            ref = line[1:].strip()
            if ref == '!':
                idx = len(hist_cmds)
            else:
                try:
                    idx = int(ref)
                except ValueError:
                    idx = 0
            if not (1 <= idx <= len(hist_cmds)):
                print("履歴にありません。history で番号を確認して !番号 で再実行（!! は直前）")
                continue
            line = hist_cmds[idx - 1]
            print(f"vba> {line}")
        hist_cmds.append(line)
        try:
            lex = shlex.shlex(line, posix=True)
            lex.whitespace_split = True
            lex.escape = ''                      # Windows パスの \ をエスケープ扱いしない
            lex.commenters = ''                  # 行中の # をコメント扱いしない（#FF0000 等が消える）
            tokens = list(lex)
        except ValueError as e:
            print(f"引数の解析に失敗: {e}")
            continue
        try:
            sub_args, unknown = parser.parse_known_args(tokens)
        except SystemExit:
            continue                             # 引数エラーは argparse が表示済み
        unknown = [u for u in unknown if u not in ("--visible", "-v")]
        if unknown:
            print(f"不明な引数/オプション: {' '.join(unknown)}")
            continue
        if not sub_args.command or sub_args.command in ('shell',):
            print("このコマンドはセッション内で実行できません")
            continue
        try:
            table[sub_args.command](sub_args)
        except SystemExit as e:
            if e.code not in (0, None):
                print(f"（終了コード {e.code}）")
        except KeyboardInterrupt:
            print("（中断しました）")
        except Exception as e:
            print(f"エラー: {e}")
    return True


def _command_table():
    """コマンド名→実装の対応表（main と batch で共用）"""
    return {
        "check":             cmd_check,
        "check-bas":         cmd_check_bas,
        "diag":              cmd_diag,
        "setup-check":       cmd_setup_check,
        "list-open":         cmd_list_open,
        "list":              cmd_list,
        "list-modules":      cmd_list_modules,
        "get":               cmd_get,
        "replace-procedure": cmd_replace_procedure,
        "add-procedure":     cmd_add_procedure,
        "add-module":        cmd_add_module,
        "delete-procedure":  cmd_delete_procedure,
        "grep":              cmd_grep,
        "code-replace":      cmd_code_replace,
        "docs":              cmd_docs,
        "call-graph":        cmd_call_graph,
        "checkup":           cmd_checkup,
        "健康診断":            cmd_checkup,
        "test":              cmd_test,
        "テスト":              cmd_test,
        "impact":            cmd_impact,
        "影響範囲":            cmd_impact,
        "replace-module":    cmd_replace_module,
        "delete-module":     cmd_delete_module,
        "export-module":     cmd_export_module,
        "export-all":        cmd_export_all,
        "list-backups":      cmd_list_backups,
        "restore":           cmd_restore,
        "reorder-macro":     cmd_reorder_macro,
        "list-shortcuts":    cmd_list_shortcuts,
        "run-macro":         cmd_run_macro,
        "read-range":        cmd_read_range,
        "read-selection":    cmd_read_selection,
        "sheet-info":        cmd_sheet_info,
        "snapshot":          cmd_snapshot,
        "snapshot-diff":     cmd_snapshot_diff,
        "wiring":            cmd_wiring,
        "配線図":              cmd_wiring,
        "screenshot":        cmd_screenshot,
        "write-range":       cmd_write_range,
        "clear-range":       cmd_clear_range,
        "format-range":      cmd_format_range,
        "sheet":             cmd_sheet,
        "table":             cmd_table,
        "name":              cmd_name,
        "row":               cmd_row,
        "col":               cmd_col,
        "copy-range":        cmd_copy_range,
        "fill":              cmd_fill,
        "sort":              cmd_sort,
        "autofilter":        cmd_autofilter,
        "find":              cmd_find,
        "find-replace":      cmd_find_replace,
        "save":              cmd_save,
        "save-as":           cmd_save_as,
        "export-pdf":        cmd_export_pdf,
        "print-setup":       cmd_print_setup,
        "cond-format":       cmd_cond_format,
        "hyperlink":         cmd_hyperlink,
        "validation":        cmd_validation,
        "freeze":            cmd_freeze,
        "comment":           cmd_comment,
        "chart":             cmd_chart,
        "chart-config":      cmd_chart_config,
        "pivot-field":       cmd_pivot_field,
        "pivot-calc":        cmd_pivot_calc,
        "pivot":             cmd_pivot,
        "slicer":            cmd_slicer,
        "calc-mode":         cmd_calc_mode,
        "powerquery":        cmd_powerquery,
        "connection":        cmd_connection,
        "datamodel":         cmd_datamodel,
        "printer-list":      cmd_printer_list,
        "printer-setup":     cmd_printer_setup,
        "batch":             cmd_batch,
        "shell":             cmd_shell,
    }


def main():
    setup_encoding()
    parser = build_parser()
    args, unknown = parser.parse_known_args()

    # 未知オプションの黙殺はタイポを事故に変える（例: clear-range --content が
    # 「値のみクリア」でなく既定の全消し Clear() に化ける）。グローバルの
    # --visible/-v だけ許容し、それ以外の残留はエラーで止める。
    unknown = [u for u in unknown if u not in ("--visible", "-v")]
    if unknown:
        print(f"エラー: 不明な引数/オプションです: {' '.join(unknown)}")
        print("  タイプミスの可能性があります。--help で正しいオプションを確認してください。")
        sys.exit(1)

    cmds = _command_table()

    if args.command in cmds:
        ok = False
        try:
            try:
                ok = cmds[args.command](args)
            except SystemExit:
                raise
            except Exception as e:
                print(f"エラー: {e}")
                sys.exit(1)
        finally:
            cleanup_excel()
        # 明示的に False を返したコマンドは失敗(1)、それ以外は成功(0)
        sys.exit(0 if ok is not False else 1)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
