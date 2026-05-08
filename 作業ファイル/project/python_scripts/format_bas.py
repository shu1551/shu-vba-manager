"""
format_bas.py - VBA .bas ファイル自動整形ツール

自動でやること:
  1. Sub 内の Dim 宣言をまとめて先頭へ移動 + ' --- 変数宣言 --- を付与
  2. On Error / Exit Sub|Function / GoTo ラベル の前後に空行
  3. End Sub|Function の直前に空行
  4. Sub と Sub の間の空行を1行に統一
  5. 連続する空行を最大1行に整理

使い方:
  py format_bas.py <モジュール名>          # 例: py format_bas.py 台帳マクロ
  py format_bas.py <モジュール名> --apply  # 整形後 Excel へも反映

引数なしで起動すると scripts ディレクトリの .bas 一覧を表示。
"""

import os
import re
import sys
import shutil
from datetime import datetime

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# CP932 読み書き
# ============================================================

def read_bas(path):
    for enc in ('cp932', 'utf-8-sig', 'utf-8'):
        try:
            with open(path, 'r', encoding=enc) as f:
                return f.read().replace('\r\n', '\n').replace('\r', '\n')
        except (UnicodeDecodeError, LookupError):
            continue
    raise ValueError(f'読み込めません: {path}')

def write_bas(path, content):
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = path.replace('.bas', f'_backup_{ts}.bas')
    shutil.copy2(path, backup)
    print(f'バックアップ: {os.path.basename(backup)}')
    with open(path, 'w', encoding='cp932', newline='\n') as f:
        f.write(content)

# ============================================================
# Sub/Function の解析
# ============================================================

SUB_START = re.compile(
    r'^(Public\s+|Private\s+)?(Sub|Function)\s+\w', re.IGNORECASE)
SUB_END   = re.compile(r'^End\s+(Sub|Function)\s*$', re.IGNORECASE)
DIM_LINE  = re.compile(r'^(Dim|Public|Private|Static)\s+', re.IGNORECASE)
LABEL_LINE = re.compile(r'^\w[\w\d]*:\s*$')   # GoTo ラベル（行末が :）
ON_ERROR  = re.compile(r'^On\s+Error\s+', re.IGNORECASE)
EXIT_KW   = re.compile(r'^Exit\s+(Sub|Function|For|Do)\s*$', re.IGNORECASE)


# ============================================================
# 自動セクションコメント判定パターン
# ============================================================

INPUT_CHECK_RX = [
    re.compile(r'^If\s+Not\s+\w+\s*\(', re.IGNORECASE),         # If Not データ行判定(...)
    re.compile(r'^If\s+ActiveSheet\.Name', re.IGNORECASE),
    re.compile(r'^If\s+ActiveCell\.', re.IGNORECASE),
    re.compile(r'^If\s+Intersect\s*\(', re.IGNORECASE),
    re.compile(r'^If\s+Selection\.', re.IGNORECASE),
    re.compile(r'^Set\s+Target\s*=\s*Selection', re.IGNORECASE),
]
FILE_DIALOG_RX    = re.compile(r'Application\.FileDialog\b', re.IGNORECASE)
WORKBOOKS_OPEN_RX = re.compile(r'Workbooks\.Open\b', re.IGNORECASE)
SET_TARGET_RX     = re.compile(
    r'^(Set\s+\w+\s*=\s*Active\w+|\w+\s*=\s*ActiveCell\.)', re.IGNORECASE)
SCREEN_OFF_RX     = re.compile(
    r'^Application\.ScreenUpdating\s*=\s*False', re.IGNORECASE)
SCREEN_ON_RX      = re.compile(
    r'^Application\.ScreenUpdating\s*=\s*True', re.IGNORECASE)
SAVE_CLOSE_RX     = re.compile(
    r'^(ActiveWorkbook\.Save|ActiveWindow\.Close|\w+\.Save\b|\w+\.Close\b)',
    re.IGNORECASE)
ERROR_LABEL_RX    = re.compile(
    r'^(err\w*|Cleanup\w*|\w*Error\w*):\s*$', re.IGNORECASE)
LOOP_START_RX     = re.compile(
    r'^(Do\s+(While|Until)|For\s+(Each|\w+\s*=))', re.IGNORECASE)
MAIN_LOOP_HINT_RX = re.compile(
    r'(Workbooks\.Open|ActiveCell\.Offset\(1\)|EntireRow\.Hidden)',
    re.IGNORECASE)


def split_into_blocks(lines):
    """空行で区切られたブロックリストを返す"""
    blocks = []
    cur = []
    for line in lines:
        if line.strip() == '':
            if cur:
                blocks.append(cur)
                cur = []
        else:
            cur.append(line)
    if cur:
        blocks.append(cur)
    return blocks


def block_has_leading_comment(block):
    """ブロックの先頭行が ' で始まるコメントか"""
    if not block:
        return False
    return block[0].strip().startswith("'")


def classify_block(block):
    """
    ブロックの内容からセクションコメント文字列を判定して返す。
    判定不能なら None を返す。誤検出を避けるため確信できるパターンのみ。
    """
    if not block:
        return None
    first = block[0].strip()
    joined = '\n'.join(l.strip() for l in block)

    # エラーハンドララベル
    if ERROR_LABEL_RX.match(first):
        return 'エラーハンドラ'

    # 入力チェック（If文 + Exit Sub/Function を含む）
    if 'Exit Sub' in joined or 'Exit Function' in joined:
        if any(rx.match(first) for rx in INPUT_CHECK_RX):
            return '入力チェック'

    # 対象の特定（Set ws = ActiveSheet, ACRC = ActiveCell.Row 等）
    if SET_TARGET_RX.match(first) and len(block) <= 5:
        # 「Set Target = Selection + If Intersect」のパターンなら入力チェック扱い
        if 'Intersect' in joined or 'Exit Sub' in joined:
            return '入力チェック'
        return '対象の特定'

    # ファイル選択ダイアログ
    if FILE_DIALOG_RX.search(joined):
        return 'ファイル選択'

    # ファイルを開く
    if WORKBOOKS_OPEN_RX.search(joined) and 'FileDialog' not in joined:
        return 'ファイルを開く'

    # メインループ
    if LOOP_START_RX.match(first) and MAIN_LOOP_HINT_RX.search(joined):
        return 'メインループ'

    # 描画停止（単独行）
    if SCREEN_OFF_RX.match(first) and len(block) == 1:
        return '描画停止'

    # 後処理（ScreenUpdating = True を含む短いブロック）
    if any(SCREEN_ON_RX.match(l.strip()) for l in block) and len(block) <= 5:
        if any(SAVE_CLOSE_RX.match(l.strip()) for l in block):
            return '保存して閉じる'
        return '後処理'

    return None


def split_into_subs(lines):
    """
    lines を (前置き, [(sub_lines, end_idx), ...]) に分割して返す。
    戻り値: (header_lines, sub_blocks)
      header_lines : Sub が始まる前の行リスト
      sub_blocks   : [(sub_lines_list, original_start_lineno), ...]
    """
    header = []
    blocks = []
    i = 0
    while i < len(lines):
        if SUB_START.match(lines[i].strip()):
            start = i
            sub_lines = []
            while i < len(lines):
                sub_lines.append(lines[i])
                if SUB_END.match(lines[i].strip()):
                    break
                i += 1
            blocks.append((sub_lines, start))
            i += 1
        else:
            if not blocks:
                header.append(lines[i])
            i += 1
    return header, blocks


# ============================================================
# 1Sub の整形
# ============================================================

def format_one_sub(lines):
    """
    1つの Sub/Function のコード行リスト（Sub〜End Sub）を整形して返す。
    1. Dim 宣言を先頭に集約 + ' --- 変数宣言 ---' を付与
    2. 空行で区切られたブロック単位にコメント自動判定
       入力チェック / 対象の特定 / ファイル選択 / ファイルを開く /
       メインループ / 後処理 / 保存して閉じる / エラーハンドラ
    3. ブロック間に空行を1行ずつ
    """
    if not lines:
        return lines

    first_line = lines[0]
    last_line  = lines[-1]
    body = lines[1:-1]

    # ---- Step 1: Dim 宣言を抽出 ----
    dim_lines   = []
    other_lines = []
    dim_done    = False
    for line in body:
        s = line.strip()
        if not dim_done and (DIM_LINE.match(s) or s == ''):
            if DIM_LINE.match(s):
                dim_lines.append(line)
        else:
            if DIM_LINE.match(s) and not dim_done:
                dim_lines.append(line)
            else:
                dim_done = True
                other_lines.append(line)

    # ---- Step 2: その他をブロック単位に分解 ----
    blocks = split_into_blocks(other_lines)

    # ---- Step 3: 新しい本体を組み立てる ----
    new_body = []

    # Dim ブロック
    if dim_lines:
        new_body.append("    ' --- 変数宣言 ---")
        new_body.extend(dim_lines)

    # 各ブロックを順に追加（コメント自動付与）
    for block in blocks:
        if new_body and new_body[-1].strip() != '':
            new_body.append('')

        comment = classify_block(block)
        if comment and not block_has_leading_comment(block):
            new_body.append(f"    ' --- {comment} ---")

        new_body.extend(block)

    # End Sub の前に空行
    if new_body and new_body[-1].strip() != '':
        new_body.append('')

    # 連続空行を1行に整理
    cleaned = []
    prev_blank = False
    for line in new_body:
        is_blank = (line.strip() == '')
        if is_blank and prev_blank:
            continue
        cleaned.append(line)
        prev_blank = is_blank

    return [first_line] + cleaned + [last_line]


# ============================================================
# ファイル全体の整形
# ============================================================

def format_content(content):
    lines = content.split('\n')
    header, blocks = split_into_subs(lines)

    # ヘッダーの末尾空行を整理
    while header and header[-1].strip() == '':
        header.pop()

    result = list(header)

    for i, (sub_lines, _) in enumerate(blocks):
        result.append('')   # Sub 間に1行空行
        formatted = format_one_sub(sub_lines)
        result.extend(formatted)

    # 末尾の空行を1行に
    while len(result) > 1 and result[-1].strip() == '' and result[-2].strip() == '':
        result.pop()
    result.append('')  # ファイル末尾は1行空行

    return '\n'.join(result)


# ============================================================
# メイン
# ============================================================

def format_module(module_name, apply_to_excel=False):
    bas_path = os.path.join(SCRIPTS_DIR, f'{module_name}.bas')
    if not os.path.exists(bas_path):
        print(f'ERROR: {bas_path} が見つかりません')
        return False

    print(f'整形中: {bas_path}')
    original = read_bas(bas_path)
    orig_lines = len(original.split('\n'))

    formatted = format_content(original)
    new_lines = len(formatted.split('\n'))

    write_bas(bas_path, formatted)
    print(f'整形完了: {orig_lines}行 → {new_lines}行')

    if apply_to_excel:
        print('Excelへ反映中...')
        import subprocess
        vba_mgr = os.path.join(SCRIPTS_DIR, 'vba_manager.py')
        r = subprocess.run(
            [sys.executable, vba_mgr, 'replace-module',
             module_name, bas_path],
            encoding='utf-8', capture_output=False,
            cwd=SCRIPTS_DIR)
        return r.returncode == 0

    return True


def list_bas_files():
    files = sorted(f for f in os.listdir(SCRIPTS_DIR) if f.endswith('.bas'))
    if not files:
        print('.bas ファイルが見つかりません')
        return
    print('整形可能な .bas ファイル:')
    for f in files:
        path = os.path.join(SCRIPTS_DIR, f)
        content = read_bas(path)
        n = len(content.split('\n'))
        subs = sum(1 for l in content.split('\n')
                   if SUB_START.match(l.strip()))
        print(f'  {f:<30} {n}行  {subs}Sub')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        list_bas_files()
        print('\n使い方: py format_bas.py <モジュール名> [--apply]')
        sys.exit(0)

    module_name = sys.argv[1].replace('.bas', '')
    apply_flag  = '--apply' in sys.argv

    ok = format_module(module_name, apply_to_excel=apply_flag)
    sys.exit(0 if ok else 1)
