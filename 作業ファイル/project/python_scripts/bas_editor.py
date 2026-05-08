"""
bas_editor.py - CP932 .bas ファイルの安全な編集ユーティリティ

【対策】
1. 読み込み時に改行コードを \n に統一（\r\n 混在問題を排除）
2. Sub/Function の置換は行単位で行い、文字列位置に頼らない
3. 書き込み後に必ず検証（行数・Sub名の存在確認）
4. 書き込み前に自動バックアップ
"""

import os
import shutil
from datetime import datetime


def read_bas(path: str) -> str:
    """CP932で読み込み、改行を \n に統一して返す"""
    with open(path, 'r', encoding='cp932') as f:
        content = f.read()
    return content.replace('\r\n', '\n').replace('\r', '\n')


def write_bas(path: str, content: str):
    """バックアップを取ってからCP932で書き込む"""
    # バックアップ
    if os.path.exists(path):
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup = path.replace('.bas', f'_backup_{ts}.bas')
        shutil.copy2(path, backup)
        print(f'バックアップ: {os.path.basename(backup)}')

    with open(path, 'w', encoding='cp932', newline='\n') as f:
        f.write(content)


def validate_bas(path: str, expected_min_lines: int, expected_max_lines: int,
                 required_subs: list = None):
    """書き込み後の検証"""
    content = read_bas(path)
    lines = content.split('\n')
    n = len(lines)

    errors = []

    if not (expected_min_lines <= n <= expected_max_lines):
        errors.append(f'行数異常: {n}行（期待: {expected_min_lines}〜{expected_max_lines}行）')

    if required_subs:
        for sub in required_subs:
            if sub not in content:
                errors.append(f'Sub/Function が見つからない: {sub}')

    # 重複チェック
    sub_names = []
    for line in lines:
        s = line.strip()
        for prefix in ('Sub ', 'Function ', 'Public Sub ', 'Private Sub ',
                       'Public Function ', 'Private Function '):
            if s.startswith(prefix):
                name = s[len(prefix):].split('(')[0].strip()
                if name in sub_names:
                    errors.append(f'Sub/Function 重複: {name}')
                sub_names.append(name)

    if errors:
        print('【検証エラー】')
        for e in errors:
            print(f'  ✗ {e}')
        raise ValueError('bas_editor: 検証失敗 — ファイルは保存されましたがエラーを確認してください')
    else:
        print(f'検証OK: {n}行, {len(sub_names)}個のSub/Function')


def replace_sub(path: str, sub_name: str, new_code: str,
                expected_line_range: tuple = None):
    """
    指定したSub/Functionを新しいコードで置換する（行単位）

    sub_name: 'Sub 計算式名前定義を適用' のように Sub/Function キーワード込みで指定
    new_code: 置換後のコード（Sub〜End Sub/Function まで）
    expected_line_range: 置換後の想定行数 (min, max)
    """
    content = read_bas(path)
    lines = content.split('\n')

    # 開始行を探す
    start_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith(sub_name):
            start_idx = i
            break

    if start_idx is None:
        raise ValueError(f'見つかりません: {sub_name}')

    # 対応する End Sub/End Function を探す
    is_func = 'Function' in sub_name
    end_keyword = 'End Function' if is_func else 'End Sub'
    end_idx = None
    for i in range(start_idx + 1, len(lines)):
        if lines[i].strip() == end_keyword:
            end_idx = i
            break

    if end_idx is None:
        raise ValueError(f'End が見つかりません: {sub_name}')

    print(f'置換対象: {start_idx+1}〜{end_idx+1}行目')

    # 置換
    new_lines = new_code.replace('\r\n', '\n').split('\n')
    lines = lines[:start_idx] + new_lines + lines[end_idx+1:]
    new_content = '\n'.join(lines)

    write_bas(path, new_content)

    # 検証
    if expected_line_range:
        validate_bas(path, expected_line_range[0], expected_line_range[1],
                     required_subs=[sub_name])
    else:
        # 行数の妥当性だけ確認（元ファイルの±500行）
        orig_n = len(content.split('\n'))
        validate_bas(path, max(1, orig_n - 500), orig_n + 500,
                     required_subs=[sub_name])


def append_sub(path: str, new_code: str, expected_line_range: tuple = None):
    """
    Subを末尾に追記する（重複チェック付き）

    new_code: 追記するコード（Sub/Function 名を含む）
    """
    content = read_bas(path)

    # 重複チェック: new_code に含まれる Sub/Function 名を抽出
    for line in new_code.split('\n'):
        s = line.strip()
        for prefix in ('Sub ', 'Function ', 'Public Sub ', 'Private Sub '):
            if s.startswith(prefix):
                name = s[len(prefix):].split('(')[0].strip()
                if f'Sub {name}(' in content or f'Function {name}(' in content:
                    raise ValueError(f'重複エラー: {name} は既に存在します。replace_sub を使ってください')

    new_content = content.rstrip('\n') + '\n\n' + new_code.strip() + '\n'
    write_bas(path, new_content)

    if expected_line_range:
        validate_bas(path, expected_line_range[0], expected_line_range[1])
    else:
        orig_n = len(content.split('\n'))
        added_n = len(new_code.split('\n'))
        validate_bas(path, orig_n + added_n - 5, orig_n + added_n + 5)


if __name__ == '__main__':
    # 動作確認
    import sys
    if len(sys.argv) > 1:
        path = sys.argv[1]
        content = read_bas(path)
        lines = content.split('\n')
        print(f'{path}: {len(lines)}行')
        subs = [l.strip() for l in lines if l.strip().startswith('Sub ') or l.strip().startswith('Function ')]
        print(f'Sub/Function数: {len(subs)}')
