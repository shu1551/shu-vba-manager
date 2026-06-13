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
import argparse
import time
import unicodedata
import pythoncom
import win32com.client

# ---- パス定数 ----
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR  = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', 'backups'))
LAST_PROC_FILE = os.path.join(SCRIPT_DIR, '_last_proc.vba')   # get の出力先
XL_EXTS     = ('.xlsm', '.xlam', '.xlsx', '.xls', '.xlsb')


# ================================================================
# ユーティリティ
# ================================================================

def setup_encoding():
    sys.stdout.reconfigure(encoding='utf-8')
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    pythoncom.CoInitialize()


def looks_like_xl_file(s):
    """文字列がExcelファイルパスっぽいか判定"""
    if not s:
        return False
    lower = s.lower()
    return (any(lower.endswith(e) for e in XL_EXTS)
            or os.sep in s or '/' in s
            or os.path.exists(s))


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


def get_workbook(target_file_arg=None):
    """
    target_file_arg が None/空 → アクティブExcelブックを自動使用
    それ以外 → 既に開いているか確認、なければ新規オープン
    戻り値: (xl, wb)
    """
    pythoncom.CoInitialize()

    if not target_file_arg:
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            raise Exception(
                "Excelが起動していません。Excelを開いてから実行してください。")
        wb = xl.ActiveWorkbook
        if wb is None:
            raise Exception(
                "アクティブなブックがありません。Excelでブックを開いてください。")
        print(f"対象ブック: {wb.Name}  (アクティブブック自動検出)")
        return xl, wb

    target_path = smart_path_resolve(target_file_arg)
    if not target_path:
        raise Exception(f"ファイルが見つかりません: {target_file_arg}")

    # 既に開いているか確認
    excel_running = False
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        excel_running = True
        for i in range(1, xl.Workbooks.Count + 1):
            wb = xl.Workbooks(i)
            if wb.FullName.lower() == target_path.lower():
                print(f"対象ブック: {wb.Name}  (既に開いています)")
                return xl, wb
    except Exception:
        pass

    # 新規オープン
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = True
    wb = xl.Workbooks.Open(target_path)
    print(f"対象ブック: {wb.Name}  (新規オープン)")
    if not excel_running:
        # COM起動のExcelは起動処理が走らず、アドインや PERSONAL.XLSB が読み込まれない
        print("注意: Excelが未起動だったため自動化用に新規起動しました。")
        print("      このExcelにはアドイン(秀.xlam等)・PERSONAL.XLSB が読み込まれていません。")
        print("      普段使いにはこのウィンドウを閉じて、手動起動した Excel で開き直してください。")
    return xl, wb


def make_backup(wb_fullname, label):
    """バックアップを作成（タイムスタンプ付き・同系列は直近5世代まで保持）"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ext = os.path.splitext(wb_fullname)[1] or '.xlsm'
    stamp = time.strftime("%Y%m%d_%H%M%S")
    prefix = os.path.basename(wb_fullname) + f".backup_before_{label}_"
    backup_path = os.path.join(BACKUP_DIR, prefix + stamp + ext)
    try:
        shutil.copy2(wb_fullname, backup_path)
        print(f"バックアップ作成: backups/{os.path.basename(backup_path)}")
    except Exception as e:
        print(f"警告: バックアップ失敗 ({e})")
        return
    # 同じ系列（同ブック・同ラベル）の古い世代を間引いて5世代までにする
    try:
        olds = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(prefix))
        for f in olds[:-5]:
            os.remove(os.path.join(BACKUP_DIR, f))
    except Exception:
        pass


def read_code_file(path):
    """コードファイルを UTF-8 / CP932 で読み込む"""
    for enc in ['utf-8-sig', 'utf-8', 'cp932']:
        try:
            with open(path, 'r', encoding=enc) as f:
                return f.read()
        except Exception:
            continue
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
    # コメント行を除外して検索
    clean_lines = []
    for line in code.split('\n'):
        stripped = line.strip()
        if stripped.startswith("'") or stripped.lower().startswith("rem "):
            continue
        clean_lines.append(stripped)
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
                return True
            except Exception as e:
                print(f"警告: モジュールバックアップ失敗 ({e})")
    return False


# ================================================================
# コマンド実装
# ================================================================

def cmd_diag(args):
    """動作確認"""
    print("Syntax OK")
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
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
    try:
        xl = win32com.client.GetObject(None, 'Excel.Application')
    except Exception:
        print('Excelは起動していません')
        return
    for wb in xl.Workbooks:
        print(wb.FullName)



def cmd_list(args):
    """マクロ(プロシージャ)一覧"""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)

    pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(\)]+)',
        re.IGNORECASE | re.MULTILINE
    )
    macros = []
    for comp in wb.VBProject.VBComponents:
        # --standard オプション指定時は標準モジュール(Type==1)のみを対象とする
        if getattr(args, 'standard', False) and comp.Type != 1:
            continue
        cm = comp.CodeModule
        if cm.CountOfLines == 0:
            continue
        for m in pattern.finditer(cm.Lines(1, cm.CountOfLines)):
            name = m.group(1)
            if name not in macros:
                macros.append(name)

    print(f"マクロ数: {len(macros)}")
    for name in macros:
        print(f"MACRO:{name}")
    return True


def cmd_list_modules(args):
    """モジュール一覧"""
    target_file, _ = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)

    type_names = {1: '標準モジュール', 2: 'クラスモジュール',
                  3: 'フォーム', 100: 'シート/ThisWorkbook'}
    for comp in wb.VBProject.VBComponents:
        tname = type_names.get(comp.Type, f'Type={comp.Type}')
        print(f"MODULE:{comp.Name}  ({tname})")
    return True


def cmd_get(args):
    """プロシージャのコードを取得・表示・ファイル保存

    書式:
      get <macro_name>                       全モジュールから検索
      get <module_name> <macro_name>         モジュール指定（スペース区切り）
      get <module_name>.<macro_name>         モジュール指定（ドット区切り）
    """
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: get [excel_file] <macro_name>  または  get [excel_file] <module_name> <macro_name>")
        return False

    # モジュール指定の解析
    module_name = None
    if len(rest) >= 2 and not looks_like_xl_file(rest[1]):
        # 書式: get <module_name> <macro_name>
        module_name = rest[0]
        macro_name  = rest[1]
    elif len(rest) == 1 and '.' in rest[0] and not looks_like_xl_file(rest[0]):
        # 書式: get <module_name>.<macro_name>
        module_name, macro_name = rest[0].split('.', 1)
    else:
        macro_name = rest[0]

    if module_name:
        print(f"モジュール指定: {module_name}")

    xl, wb = get_workbook(target_file)

    # モジュール未指定時：同名プロシージャが複数モジュールにある場合はエラー
    # （違うフォームの同名イベントを黙って掴む事故を防ぐ。replace-procedure と同じ流儀）
    if not module_name:
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
            print(f"  モジュールを指定してください。例: py vba_manager.py get {matched_modules[0]} {macro_name}")
            return False

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
        while lines:
            last = lines[-1].strip()
            if last == '' or re.match(
                r'^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+',
                last, re.IGNORECASE):
                lines.pop()
            else:
                break
        clean = '\n'.join(lines) + '\n'

        # _last_proc.vba に UTF-8 で保存
        with open(LAST_PROC_FILE, 'w', encoding='utf-8') as f:
            f.write(clean)

        print(f"モジュール  : {comp.Name}")
        print(f"プロシージャ: {macro_name}")
        print(f"保存先      : {LAST_PROC_FILE}")
        print("=" * 60)
        print(clean)
        print("=" * 60)
        return True

    print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
    return False


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

    # new_code の末尾に次のプロシージャの Sub/Function 宣言が混入していたら除去
    code_lines = new_code.rstrip('\n').split('\n')
    while len(code_lines) > 1:
        last = code_lines[-1].strip()
        if last == '':
            code_lines.pop()
        elif pattern.match(last) and code_lines[-1].strip() != code_lines[0].strip():
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

    make_backup(wb.FullName, macro_name)

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
        print("変更はありません。")

    # 確認プロンプト
    if not getattr(args, 'yes', False):
        ans = input(f"プロシージャ '{macro_name}' を置換しますか？ (y/N): ")
        if ans.strip().lower() not in ('y', 'yes'):
            print("キャンセルされました。")
            return False

    # モジュール単位のバックアップ
    make_module_backup(wb, target_comp.Name)

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
        end_pattern = re.compile(
            r'^\s*End\s+(?:Sub|Function)\s*$', re.IGNORECASE
        )

        sub_line_idx = None
        proc_end_idx = None
        attr_block = []

        for idx, line in enumerate(bas_lines):
            if sub_line_idx is None and proc_pattern.match(line):
                sub_line_idx = idx
                # Sub宣言の直後の Attribute行を収集
                check = idx + 1
                while check < len(bas_lines) and bas_lines[check].strip().startswith('Attribute '):
                    attr_block.append(bas_lines[check])
                    check += 1
            elif sub_line_idx is not None and end_pattern.match(line):
                proc_end_idx = idx
                break

        if sub_line_idx is None or proc_end_idx is None:
            os.remove(tmp_bas)
            continue

        if not attr_block:
            # Attribute行なし → 従来の InsertLines 方式（高速・モジュール順維持）
            # InsertLines は末尾改行を余分な空行として挿入するため取り除く
            os.remove(tmp_bas)
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
            os.remove(tmp_bas)
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
        try:
            wb.Save()
            time.sleep(0.5)
            pythoncom.PumpWaitingMessages()
            wb.VBProject.VBComponents.Remove(comp)
            time.sleep(1.5)
            pythoncom.PumpWaitingMessages()
            wb.VBProject.VBComponents.Import(tmp_bas)
            time.sleep(1.5)
            pythoncom.PumpWaitingMessages()
            wb.Save()
        finally:
            xl.DisplayAlerts = True
            if os.path.exists(tmp_bas):
                os.remove(tmp_bas)
        print(f"置換完了: [{module_name}] '{macro_name}' → 保存しました (Attribute保持)")
        return True

    print(f"エラー: プロシージャ '{macro_name}' が見つかりません")
    return False


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

    xl, wb = get_workbook(target_file)
    make_backup(wb.FullName, f"module_{module_name}")
    # モジュール単位のバックアップ
    make_module_backup(wb, module_name)
    print(f"モジュール '{module_name}' を Remove+Import で置換中...")

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() == module_name.lower():
            xl.DisplayAlerts = False
            try:
                wb.Save()
                time.sleep(0.5)
                pythoncom.PumpWaitingMessages()
                wb.VBProject.VBComponents.Remove(comp)
                time.sleep(1.5)
                pythoncom.PumpWaitingMessages()
                wb.VBProject.VBComponents.Import(resolved)
                time.sleep(1.5)
                pythoncom.PumpWaitingMessages()
                wb.Save()
            finally:
                xl.DisplayAlerts = True
            print(f"置換完了: モジュール '{module_name}' → 保存しました")
            return True

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
        r'(?P<kind>Sub|Function)\s+(?P<name>[^\s\(]+)',
        re.IGNORECASE
    )
    end_pattern = re.compile(r'^\s*End\s+(?:Sub|Function)\s*$', re.IGNORECASE)

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
            j = i + 1
            while j < n and not end_pattern.match(lines[j]):
                j += 1
            if j >= n:
                break
            blocks.append({
                'name':  m.group('name'),
                'kind':  m.group('kind').lower(),
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


# 一覧で非表示にしているマクロ（並べ替えの可視ブロック判定にも使う）
_HIDDEN_MACROS = {"ホイール有効化", "ホイール無効化", "マウスホイールフック"}


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
        print("使い方: reorder-macro <macro_name> <up|down>")
        sys.exit(1)
    macro_name = rest[0]
    direction  = rest[1].lower()
    if direction not in ('up', 'down'):
        print("方向は up または down を指定してください")
        sys.exit(1)

    xl, wb = get_workbook(None)  # ActiveWorkbook 自動検出

    # 対象マクロを含む標準モジュールを特定
    target_comp = None
    for comp in wb.VBProject.VBComponents:
        if comp.Type != 1:  # 標準モジュールのみ対象
            continue
        try:
            comp.CodeModule.ProcStartLine(macro_name, 0)
            target_comp = comp
            break
        except Exception:
            continue

    if target_comp is None:
        print(f"エラー: マクロ '{macro_name}' が標準モジュールに見つかりません")
        sys.exit(2)

    module_name = target_comp.Name

    # モジュールをエクスポートして CP932 のまま読み込む
    tmp_bas = os.path.join(SCRIPT_DIR, f"_tmp_reorder_{module_name}.bas")
    target_comp.Export(tmp_bas)
    with open(tmp_bas, 'rb') as f:
        bas_text = f.read().decode('cp932')

    header, blocks, trailing = _parse_module_blocks(bas_text)

    # 対象ブロックの index
    target_idx = None
    for i, b in enumerate(blocks):
        if b['name'] == macro_name:
            target_idx = i
            break
    if target_idx is None:
        os.remove(tmp_bas)
        print(f"エラー: モジュール {module_name} に Sub '{macro_name}' が見つかりません")
        sys.exit(2)

    # 一覧で見える Sub だけを「可視ブロック」として抽出
    visible_indices = [
        i for i, b in enumerate(blocks)
        if b['kind'] == 'sub' and b['name'] not in _HIDDEN_MACROS
    ]

    if target_idx not in visible_indices:
        os.remove(tmp_bas)
        print(f"エラー: '{macro_name}' は一覧表示対象ではありません")
        sys.exit(2)

    vis_pos = visible_indices.index(target_idx)

    if direction == 'up':
        if vis_pos == 0:
            os.remove(tmp_bas)
            print(f"BOUNDARY: [{module_name}] 内で既に最初です")
            sys.exit(3)
        swap_block_idx = visible_indices[vis_pos - 1]
    else:
        if vis_pos == len(visible_indices) - 1:
            os.remove(tmp_bas)
            print(f"BOUNDARY: [{module_name}] 内で既に最後です")
            sys.exit(3)
        swap_block_idx = visible_indices[vis_pos + 1]

    # ブロック単位で入れ替え（Attribute 行は各ブロック内に含まれているので一緒に動く）
    blocks[target_idx], blocks[swap_block_idx] = blocks[swap_block_idx], blocks[target_idx]

    new_text = _write_module(header, blocks, trailing)
    with open(tmp_bas, 'wb') as f:
        f.write(new_text.encode('cp932'))

    make_backup(wb.FullName, f"reorder_{macro_name}")
    print(f"並べ替え中: [{module_name}] '{macro_name}' を {direction}")

    # replace-module と同じ安定化手順（sleep + PumpWaitingMessages）に揃える
    xl.DisplayAlerts = False
    try:
        wb.Save()
        time.sleep(0.5)
        pythoncom.PumpWaitingMessages()
        wb.VBProject.VBComponents.Remove(target_comp)
        time.sleep(1.5)
        pythoncom.PumpWaitingMessages()
        wb.VBProject.VBComponents.Import(tmp_bas)
        time.sleep(1.5)
        pythoncom.PumpWaitingMessages()
        wb.Save()
    finally:
        xl.DisplayAlerts = True
        if os.path.exists(tmp_bas):
            os.remove(tmp_bas)

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

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() == module_name.lower():
            out_path = os.path.join(SCRIPT_DIR, f"{module_name}.bas")
            comp.Export(out_path)
            print(f"エクスポート完了: {out_path}")
            return True

    print(f"エラー: モジュール '{module_name}' が見つかりません")
    return False


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
            if os.path.exists(tmp_file):
                os.remove(tmp_file)
            continue
        finally:
            if os.path.exists(tmp_file):
                os.remove(tmp_file)

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

    if not shortcuts:
        print("ショートカットキーが設定されているマクロはありません。")
        return True

    print(f"設定されているショートカットキー一覧 (数: {len(shortcuts)})")
    print("-" * 60)
    for item in shortcuts:
        print(f"[{item['module']}] {item['macro']} -> {item['shortcut']}")
    print("-" * 60)
    return True


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


def _resolve_range(xl, wb, spec):
    """
    範囲指定を (ws, rng) に解決する。
      'A1:D20'         → アクティブシートの範囲
      'Sheet1!A1:D20'  → シート指定の範囲
      'Sheet1'         → そのシートの UsedRange
      None / ''        → アクティブシートの UsedRange
    """
    if not spec:
        ws = wb.ActiveSheet
        return ws, ws.UsedRange

    if '!' in spec:
        sheet_part, addr = spec.split('!', 1)
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


def _cell_str(v):
    """セル値を表示用文字列に"""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _disp_width(s):
    """全角文字を2幅として数えた表示幅"""
    w = 0
    for ch in s:
        w += 2 if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else 1
    return w


def _disp_truncate(s, width):
    """表示幅 width に収まるよう切り詰める"""
    out = []
    w = 0
    for ch in s:
        cw = 2 if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else 1
        if w + cw > width:
            break
        out.append(ch)
        w += cw
    return ''.join(out)


def _disp_pad(s, width, right=False):
    """表示幅基準で width までスペース埋め（right=Trueで右寄せ）"""
    pad = width - _disp_width(s)
    if pad <= 0:
        return s
    return (' ' * pad + s) if right else (s + ' ' * pad)


def _values_to_grid(rng, use_formula=False):
    """Range の値を、列文字＋行番号つきのテキスト格子にする

    use_formula=True のときは計算結果ではなく数式(.Formula)を表示する。
    数式のないセルは定数値がそのまま入る（write-range の .Value と同じ規約）。
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
        col_w.append(min(w, 40))

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
    """シートのセル値をテキスト格子で読み取る（目・テキスト版）"""
    target_file, rest = parse_target_and_rest(args.posargs)
    spec = rest[0] if rest else None
    use_formula = getattr(args, 'formula', False)
    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, spec)
    mode = "（数式表示）" if use_formula else ""
    print(f"シート: {ws.Name}   範囲: {rng.Address}{mode}")
    print("=" * 60)
    print(_values_to_grid(rng, use_formula))
    print("=" * 60)
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
    """ブックのシート構成・使用範囲を表示（見取り図）"""
    target_file, _ = parse_target_and_rest(args.posargs)
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
            dims = "(空)"
        vis = '' if sh.Visible == -1 else '  [非表示]'
        print(f"{mark} {sh.Name}: {dims}{vis}")
    print("-" * 60)
    return True


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
                    print(f"シート: {ws.Name}   範囲: {rng.Address}")
                    print(f"画像保存: {out_path}")
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

    # 選択を A1 に戻す（ChartObject 操作で選択が動くため）
    try:
        ws.Range("A1").Select()
    except Exception:
        pass

    print(f"エラー: スクリーンショットに失敗しました ({last_err})")
    return False


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
    """文字列をセル値に変換。'='始まりは数式、数値は数値、空は None。"""
    if s is None or s == '':
        return None
    if s.startswith('='):
        return s                      # 数式 (.Value への代入で Excel が数式と解釈)
    if re.fullmatch(r'-?\d+', s):
        try:
            return int(s)
        except ValueError:
            return s
    try:
        return float(s)
    except ValueError:
        return s


def _read_tsv_grid(path):
    """TSV(タブ区切り)をセル値の 2次元タプルに変換。列数は最大行に揃える。"""
    with open(path, 'r', encoding='utf-8-sig') as f:
        text = f.read()
    text = text.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n')
    if text == '':
        return ()
    rows = [tuple(_coerce_cell(c) for c in line.split('\t'))
            for line in text.split('\n')]
    ncols = max(len(r) for r in rows)
    return tuple(r + (None,) * (ncols - len(r)) for r in rows)


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

    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, spec)

    if inline_value is not None:
        # インライン単一値: 範囲全体に同じ値 (数式可)
        rng.Value = _coerce_cell(inline_value)
        print(f"書き込み: {ws.Name}!{rng.Address} ← {inline_value}")
    else:
        path = (smart_path_resolve(tsv_opt) if tsv_opt else _LAST_VALUES_FILE)
        if not path or not os.path.exists(path):
            print(f"エラー: TSVが見つかりません: {tsv_opt or _LAST_VALUES_FILE}")
            print("  単一値ならインラインで: write-range A1 \"値\"")
            return False
        grid = _read_tsv_grid(path)
        if not grid:
            print("エラー: TSVが空です")
            return False
        nrows = len(grid); ncols = len(grid[0])
        top = ws.Cells(rng.Row, rng.Column)
        if nrows == 1 and ncols == 1:
            top.Value = grid[0][0]
            print(f"書き込み: {ws.Name}!{top.Address} ← {grid[0][0]}")
        else:
            target = ws.Range(top, ws.Cells(rng.Row + nrows - 1,
                                            rng.Column + ncols - 1))
            target.Value = grid
            print(f"書き込み: {ws.Name}!{target.Address} ← TSV {nrows}行 x {ncols}列")

    print("（保存はしていません。Excelで確認後に保存してください）")
    return True


def cmd_clear_range(args):
    """セル範囲をクリア (既定: すべて)"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: clear-range [excel_file] <range> [--contents|--formats|--all]")
        return False
    spec = rest[0]
    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, spec)
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
    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, spec)
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
        rng.Merge(); applied.append("merge")
    if getattr(args, 'unmerge', False):
        rng.UnMerge(); applied.append("unmerge")
    if getattr(args, 'autofit', False):
        rng.Columns.AutoFit(); applied.append("autofit")

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
        for sh in wb.Sheets:
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
        for sh in wb.Sheets:
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
        for sh in wb.Sheets:
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
        refers = "='" + ws.Name + "'!" + rng.Address
        wb.Names.Add(nm_name, refers)
        print(f"名前付き範囲を追加: {nm_name} → {refers}")
        print("（保存はしていません）")
        return True

    if action == 'delete':
        if len(rest) < 2:
            print("使い方: name delete <name>"); return False
        nm_name = rest[1]
        for nm in wb.Names:
            if nm.Name == nm_name or nm.Name.split('!')[-1] == nm_name:
                nm.Delete()
                print(f"名前付き範囲を削除: {nm_name}")
                print("（保存はしていません）")
                return True
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
    ws = wb.ActiveSheet
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
    ws = wb.ActiveSheet
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
    xl, wb = get_workbook(target_file)
    ws_s, rng_s = _resolve_range(xl, wb, rest[0])
    ws_d, rng_d = _resolve_range(xl, wb, rest[1])
    if getattr(args, 'values', False):
        rng_s.Copy()
        rng_d.PasteSpecial(-4163)          # xlPasteValues
        xl.CutCopyMode = False
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
    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, rest[0])
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
    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, rest[0])
    keycol = getattr(args, 'key', None)
    key_idx = _col_num(keycol) if keycol else rng.Column
    keycell = ws.Cells(rng.Row, key_idx)
    order = 2 if getattr(args, 'desc', False) else 1       # xlDescending=2 / xlAscending=1
    if getattr(args, 'header', False):
        header = 1                                          # xlYes
    elif getattr(args, 'no_header', False):
        header = 2                                          # xlNo
    else:
        header = 0                                          # xlGuess
    rng.Sort(Key1=keycell, Order1=order, Header=header)
    print(f"並べ替え: {ws.Name}!{rng.Address}  キー列={keycol or _col_letter(rng.Column)}  "
          f"{'降順' if order == 2 else '昇順'}")
    print("（保存はしていません）")
    return True


def cmd_autofilter(args):
    """オートフィルタ: autofilter [range] [--off]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    xl, wb = get_workbook(target_file)
    if getattr(args, 'off', False):
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
        rng.AutoFilter()
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
    sheets = list(wb.Sheets) if getattr(args, 'book', False) else [wb.ActiveSheet]
    look_in = -4123 if getattr(args, 'formula', False) else -4163  # xlFormulas / xlValues
    look_at = 1 if getattr(args, 'whole', False) else 2            # xlWhole / xlPart
    total = 0
    for ws in sheets:
        rng = ws.UsedRange
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
            print(f"[{ws.Name}] {addr}: {cell.Value}")
            total += 1
            cell = rng.FindNext(cell)
    if total == 0:
        print(f"'{needle}' は見つかりませんでした。")
    else:
        print(f"--- {total}件 ヒット ---")
    return True


def cmd_find_replace(args):
    """一括置換: find-replace <検索> <置換> [range] [--whole]"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if len(rest) < 2:
        print("使い方: find-replace <検索> <置換> [range] [--whole]")
        return False
    needle, repl = rest[0], rest[1]
    spec = rest[2] if len(rest) >= 3 else None
    xl, wb = get_workbook(target_file)
    if spec:
        ws, rng = _resolve_range(xl, wb, spec)
    else:
        ws = wb.ActiveSheet
        rng = ws.UsedRange
    look_at = 1 if getattr(args, 'whole', False) else 2
    rng.Replace(What=needle, Replacement=repl, LookAt=look_at, MatchCase=False)
    print(f"置換: {ws.Name}!{rng.Address}  '{needle}' → '{repl}'")
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
    """別名保存: save-as <path>（アクティブブックを対象）"""
    rest = list(args.posargs)
    if not rest:
        print("使い方: save-as <path>")
        return False
    out = os.path.abspath(rest[0])
    ext = os.path.splitext(out)[1].lower()
    fmt = {'.xlsx': 51, '.xlsm': 52, '.xlsb': 50, '.xls': 56,
           '.csv': 6, '.txt': -4158}.get(ext, 51)
    xl, wb = get_workbook(None)
    xl.DisplayAlerts = False
    try:
        wb.SaveAs(out, FileFormat=fmt)
    finally:
        xl.DisplayAlerts = True
    print(f"別名保存しました: {out}")
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
    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, rest[0])

    if getattr(args, 'clear', False):
        rng.FormatConditions.Delete()
        print(f"条件付き書式を全削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True

    op = None; f1 = None; f2 = None
    for name in ('gt', 'lt', 'eq', 'ne', 'ge', 'le'):
        v = getattr(args, name, None)
        if v is not None:
            op = _XL_COND_OP[name]; f1 = str(v); break
    if op is None and getattr(args, 'between', None):
        op = _XL_COND_OP['between']; f1, f2 = args.between[0], args.between[1]
    if op is None:
        print("比較条件がありません。--gt 100 などを指定してください。")
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
    """ハイパーリンク: hyperlink <cell> <url> [--text 表示文字] / --remove"""
    target_file, rest = parse_target_and_rest(args.posargs)
    if not rest:
        print("使い方: hyperlink <cell> <url> [--text 表示文字]  /  hyperlink <cell> --remove")
        return False
    xl, wb = get_workbook(target_file)
    ws, rng = _resolve_range(xl, wb, rest[0])
    if getattr(args, 'remove', False):
        rng.Hyperlinks.Delete()
        print(f"ハイパーリンク削除: {ws.Name}!{rng.Address}")
        print("（保存はしていません）")
        return True
    if len(rest) < 2:
        print("使い方: hyperlink <cell> <url> [--text 表示文字]")
        return False
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
    xl, wb = get_workbook(target_file)
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
        for sh in wb.Sheets:
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
        for sh in wb.Sheets:
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
        for sh in wb.Sheets:
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
            tidx = int(rest[4]) if len(rest) >= 5 else 1
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
        sid = int(rest[2]) if len(rest) >= 3 else 1
        ch.ChartStyle = sid
        print(f"スタイル設定: {cname} = {sid}"); print("（保存はしていません）"); return True

    if action == 'placement':
        pl = int(rest[2]) if len(rest) >= 3 else 1
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
            if pos.lower() in posmap:
                try:
                    ch.SeriesCollection(1).DataLabels().Position = posmap[pos.lower()]
                except Exception:
                    pass
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
        idx = int(rest[2]) if len(rest) >= 3 else 1
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
        for sh in wb.Sheets:
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
        for sh in wb.Sheets:
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
        if sheet_opt:
            dws = None
            for sh in wb.Sheets:
                if sh.Name == sheet_opt:
                    dws = sh; break
            if dws is None:
                dws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
                dws.Name = sheet_opt
            dest = dws.Range("A3")
        elif at:
            dest = ws_s.Range(at)
        else:
            dws = wb.Sheets.Add(None, wb.Sheets(wb.Sheets.Count))
            dws.Name = _unique_sheet_name(wb, "ピボット")
            dest = dws.Range("A3")

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

        funcname = (getattr(args, 'func', None) or 'sum').lower()
        func = _XL_PIVOT_FUNC.get(funcname, -4157)
        values = getattr(args, 'values', None)
        if values:
            for f in values.split(','):
                f = f.strip()
                if f:
                    df = pt.AddDataField(pt.PivotFields(f), f"{funcname}/{f}", func)

        print(f"ピボット作成: [{pt.Parent.Name}] {pt.Name}  ソース={ws_s.Name}!{rng.Address}")
        print(f"  行={getattr(args,'rows',None) or '-'}  列={getattr(args,'cols',None) or '-'}  "
              f"値={values or '-'}({funcname})")
        print("（保存はしていません）")
        return True

    print(f"未知のアクション: {action}（create|list|delete）")
    return False


def _find_pivot(wb, name):
    for sh in wb.Sheets:
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
        p = pf(field)
        if p is None:
            print(f"エラー: フィールド '{field}' が見つかりません。"); return False
        p.Caption = newname
        print(f"表示名変更: {pname}[{field}] → {newname}"); print("（保存はしていません）"); return True

    if action == 'set-format':
        code = rest[3] if len(rest) >= 4 else None
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
        for i in range(1, p.PivotItems().Count + 1):
            it = p.PivotItems().Item(i)
            try:
                it.Visible = (it.Name in wanted)
            except Exception:
                pass
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
            pt.PivotFields(cf_name).Delete()
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
    for sh in wb.Sheets:
        for pt in sh.PivotTables():
            if pt.Name == name:
                return pt, 'pivot', sh
    for sh in wb.Sheets:
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
                print(f"{sl.Name}  (フィールド={sc.SourceName}, シート={sl.Parent.Name})")
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

        sc = wb.SlicerCaches.Add2(src, field)

        # 配置先シートと座標
        at = getattr(args, 'at', None)
        dws = sh
        if at:
            anchor = dws.Range(at)
            top, left = anchor.Top, anchor.Left
        else:
            top, left = 10.0, 400.0
        sl = sc.Slicers.Add(SlicerDestination=dws, Caption=field,
                            Top=top, Left=left, Width=144.0, Height=180.0)
        # Slicers.Add の Name 引数は効かないことがあるので作成後に明示セット
        req_name = getattr(args, 'name', None)
        if req_name:
            try:
                sl.Name = req_name.replace(' ', '')
            except Exception:
                pass
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
            if sheet_name:
                for i in range(1, wb.Worksheets.Count + 1):
                    if wb.Worksheets.Item(i).Name == sheet_name:
                        ws = wb.Worksheets.Item(i); break
                if ws is None:
                    ws = wb.Worksheets.Add()
                    ws.Name = sheet_name
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
            qt.Refresh(False)                  # BackgroundQuery:=False
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
            cn_name = f"Query - {name}"
            for cn in list(wb.Connections):
                if cn.Name == cn_name:
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

_XL_CONN_TYPE = {1: 'OLEDB', 2: 'ODBC', 4: 'XMLMAP', 5: 'TEXT',
                 6: 'WEB', 7: 'DATAFEED', 8: 'MODEL', 9: 'WORKSHEET', 10: 'NOSOURCE'}


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
                                                                 [--decimals N] [--thousands] [--symbol ¥]   メジャー(DAX)作成
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
        return True

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
            try:
                decimals = int(getattr(args, 'decimals', None) or 2)
            except (TypeError, ValueError):
                decimals = 2
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

def main():
    setup_encoding()

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

    # list-open
    sub.add_parser("list-open")

    # list [excel_file]
    p = sub.add_parser("list")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--standard", action="store_true", help="標準モジュールのみを抽出")

    # list-modules [excel_file]
    p = sub.add_parser("list-modules")
    p.add_argument("posargs", nargs="*")

    # get [excel_file] <macro_name>
    p = sub.add_parser("get")
    p.add_argument("posargs", nargs="+")

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

    # list-shortcuts [excel_file]
    p = sub.add_parser("list-shortcuts")
    p.add_argument("posargs", nargs="*")

    # replace-module [excel_file] <module_name> <bas_file>
    p = sub.add_parser("replace-module")
    p.add_argument("posargs", nargs="+")

    # export-module [excel_file] <module_name>
    p = sub.add_parser("export-module")
    p.add_argument("posargs", nargs="+")

    # reorder-macro <macro_name> <up|down>
    p = sub.add_parser("reorder-macro")
    p.add_argument("posargs", nargs="+")

    # --- 目コマンド ---
    # read-range [excel_file] [range] [--formula]
    p = sub.add_parser("read-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--formula", action="store_true",
                   help="計算結果でなく数式(.Formula)を表示する")

    # read-selection [excel_file] [--formula]
    p = sub.add_parser("read-selection")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--formula", action="store_true",
                   help="計算結果でなく数式(.Formula)を表示する")

    # sheet-info [excel_file]
    p = sub.add_parser("sheet-info")
    p.add_argument("posargs", nargs="*")

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

    # clear-range [excel_file] <range> [--contents|--formats|--all]
    p = sub.add_parser("clear-range")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--contents", action="store_true", help="値のみクリア")
    p.add_argument("--formats", action="store_true", help="書式のみクリア")
    p.add_argument("--all", action="store_true", help="すべてクリア（既定）")

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

    # sheet [excel_file] <add|delete|rename|copy|activate|show|hide> ...
    p = sub.add_parser("sheet")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--after")
    p.add_argument("--before")
    p.add_argument("--clear", action="store_true", help="tab-color のクリア")

    # table [excel_file] <create|list|delete> ...
    p = sub.add_parser("table")
    p.add_argument("posargs", nargs="*")
    p.add_argument("--no-headers", dest="no_headers", action="store_true")
    p.add_argument("--at", dest="at", default=None, help="column add の挿入位置(1始まり)")
    p.add_argument("--desc", dest="desc", action="store_true", help="sort を降順に")

    # name [excel_file] <add|list|delete> ...
    p = sub.add_parser("name")
    p.add_argument("posargs", nargs="*")

    # --- 手コマンド 第2弾 ---
    # a. 編集の足回り
    p = sub.add_parser("row")          # row <insert|delete> <行番号> [本数]
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("col")          # col <insert|delete> <列文字> [本数]
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("copy-range")   # copy-range <src> <dst> [--values]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--values", action="store_true", help="値のみ貼り付け")
    p = sub.add_parser("fill")         # fill <range> [--right]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--right", action="store_true", help="右方向にフィル（既定は下）")
    p = sub.add_parser("sort")         # sort <range> [--key 列] [--desc] [--header|--no-header]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--key", help="並べ替えキー列（列文字）")
    p.add_argument("--desc", action="store_true", help="降順")
    p.add_argument("--header", action="store_true", help="先頭行を見出しとして扱う")
    p.add_argument("--no-header", dest="no_header", action="store_true", help="見出しなし")
    p = sub.add_parser("autofilter")   # autofilter [range] [--off]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--off", action="store_true", help="オートフィルタを解除")

    # b. 検索・置換
    p = sub.add_parser("find")         # find <文字> [--book] [--whole] [--formula]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--book", action="store_true", help="全シート横断で検索")
    p.add_argument("--whole", action="store_true", help="完全一致")
    p.add_argument("--formula", action="store_true", help="数式も検索対象にする")
    p = sub.add_parser("find-replace") # find-replace <検索> <置換> [range] [--whole]
    p.add_argument("posargs", nargs="*")
    p.add_argument("--whole", action="store_true", help="完全一致のみ置換")

    # c. 保存・印刷まわり
    p = sub.add_parser("save")         # save [excel_file]
    p.add_argument("posargs", nargs="*")
    p = sub.add_parser("save-as")      # save-as <path>
    p.add_argument("posargs", nargs="*")
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
    p.add_argument("--bg"); p.add_argument("--color")
    p.add_argument("--bold", action="store_true")
    p.add_argument("--clear", action="store_true", help="条件付き書式を全削除")
    p = sub.add_parser("hyperlink")    # hyperlink <cell> <url> [--text t] / --remove
    p.add_argument("posargs", nargs="*")
    p.add_argument("--text", help="表示文字")
    p.add_argument("--remove", action="store_true", help="ハイパーリンク削除")
    p = sub.add_parser("validation")   # validation <range> --list 'A,B,C' / --clear
    p.add_argument("posargs", nargs="*")
    p.add_argument("--list", help="ドロップダウン候補（カンマ区切り）")
    p.add_argument("--clear", action="store_true", help="入力規則を削除")
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

    args = parser.parse_args()

    cmds = {
        "diag":              cmd_diag,
        "list-open":         cmd_list_open,
        "list":              cmd_list,
        "list-modules":      cmd_list_modules,
        "get":               cmd_get,
        "replace-procedure": cmd_replace_procedure,
        "replace-module":    cmd_replace_module,
        "export-module":     cmd_export_module,
        "reorder-macro":     cmd_reorder_macro,
        "list-shortcuts":    cmd_list_shortcuts,
        "read-range":        cmd_read_range,
        "read-selection":    cmd_read_selection,
        "sheet-info":        cmd_sheet_info,
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
    }

    if args.command in cmds:
        try:
            ok = cmds[args.command](args)
        except SystemExit:
            raise                      # reorder-macro 等は自前の終了コードを使う
        except Exception as e:
            print(f"エラー: {e}")
            sys.exit(1)
        # 明示的に False を返したコマンドは失敗(1)、それ以外は成功(0)
        sys.exit(0 if ok is not False else 1)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
