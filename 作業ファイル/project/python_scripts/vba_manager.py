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
"""

import sys
import os
import re
import shutil
import argparse
import time
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
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
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
    return xl, wb


def make_backup(wb_fullname, label):
    """バックアップを作成"""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ext = os.path.splitext(wb_fullname)[1] or '.xlsm'
    backup_name = os.path.basename(wb_fullname) + f".backup_before_{label}{ext}"
    backup_path = os.path.join(BACKUP_DIR, backup_name)
    try:
        shutil.copy2(wb_fullname, backup_path)
        print(f"バックアップ作成: backups/{os.path.basename(backup_path)}")
    except Exception as e:
        print(f"警告: バックアップ失敗 ({e})")


def read_code_file(path):
    """コードファイルを UTF-8 / CP932 で読み込む"""
    for enc in ['utf-8-sig', 'utf-8', 'cp932']:
        try:
            with open(path, 'r', encoding=enc) as f:
                return f.read()
        except Exception:
            continue
    raise Exception(f"ファイルを読み込めません: {path}")


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

    for comp in wb.VBProject.VBComponents:
        if module_name and comp.Name.lower() != module_name.lower():
            continue
        cm = comp.CodeModule
        try:
            proc_start = cm.ProcStartLine(macro_name, 0)
            body_start = cm.ProcBodyLine(macro_name, 0)
            count      = cm.ProcCountLines(macro_name, 0)
            # ProcCountLines は ProcStartLine からの行数なので
            # body_start から取る場合は差分を引く
            body_count = count - (body_start - proc_start)
            code  = cm.Lines(body_start, body_count)
        except Exception:
            continue

        # 末尾の余分な Sub/Function 宣言行を除去
        lines = code.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n').split('\n')
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
    print(f"プロシージャ '{macro_name}' を置換中...")

    for comp in wb.VBProject.VBComponents:
        if module_opt and comp.Name.lower() != module_opt.lower():
            continue
        cm = comp.CodeModule
        try:
            start = cm.ProcStartLine(macro_name, 0)
            count = cm.ProcCountLines(macro_name, 0)
        except Exception:
            continue

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
            os.remove(tmp_bas)
            cm.DeleteLines(start, count)
            cm.InsertLines(start, new_code)
            wb.Save()
            print(f"置換完了: [{comp.Name}] '{macro_name}' → 保存しました")
            return True

        # Attribute行あり → .bas編集 → replace-module 方式
        print(f"  (Attribute行検出 → replace-module方式で処理)")

        # new_code の行を準備（Sub宣言の直後に Attribute行を挿入）
        new_lines = new_code.rstrip('\n').split('\n')
        sub_decl_pattern = re.compile(
            r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+',
            re.IGNORECASE
        )
        insert_pos = 0
        for ni, nl in enumerate(new_lines):
            if sub_decl_pattern.match(nl):
                insert_pos = ni + 1
                break
        for ai, al in enumerate(attr_block):
            new_lines.insert(insert_pos + ai, al)

        # .bas 内の対象プロシージャを置換（Sub宣言行から End Sub まで）
        bas_lines[sub_line_idx:proc_end_idx + 1] = new_lines
        new_bas = '\r\n'.join(bas_lines)

        with open(tmp_bas, 'wb') as f:
            f.write(new_bas.encode('cp932'))

        # Remove + Import
        xl.DisplayAlerts = False
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
        xl.DisplayAlerts = True
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

    xl, wb = get_workbook(target_file)
    make_backup(wb.FullName, f"module_{module_name}")
    print(f"モジュール '{module_name}' を Remove+Import で置換中...")

    for comp in wb.VBProject.VBComponents:
        if comp.Name.lower() == module_name.lower():
            xl.DisplayAlerts = False
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

    wb.VBProject.VBComponents.Remove(target_comp)
    wb.VBProject.VBComponents.Import(tmp_bas)
    wb.Save()
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

    # replace-module [excel_file] <module_name> <bas_file>
    p = sub.add_parser("replace-module")
    p.add_argument("posargs", nargs="+")

    # export-module [excel_file] <module_name>
    p = sub.add_parser("export-module")
    p.add_argument("posargs", nargs="+")

    # reorder-macro <macro_name> <up|down>
    p = sub.add_parser("reorder-macro")
    p.add_argument("posargs", nargs="+")

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
    }

    if args.command in cmds:
        try:
            cmds[args.command](args)
        except Exception as e:
            print(f"エラー: {e}")
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
