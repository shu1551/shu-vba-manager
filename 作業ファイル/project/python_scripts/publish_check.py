# -*- coding: utf-8 -*-
"""publish_check.py — xlsm/xlsx 公開前の個人情報チェック＆メタデータ空欄化を1コマンドに畳む。

2026-07-11 の実名メタデータ公開事故を受けた6項目チェック
（作成者実名 / 隠しシート / 定義名 / 外部リンク / コメント / 残留入力値）を機械で回す。

使い方:
  py publish_check.py <ブック.xlsm>              # チェックのみ（公開可否を判定表示）
  py publish_check.py <ブック.xlsm> --scrub      # dc:creator / cp:lastModifiedBy を空欄化
  py publish_check.py <A.xlsm> <B.xlsm> ...      # 複数まとめて

判定: 実名らしき作成者（イニシャル2〜4文字以外）が残っていれば終了コード1。
--scrub は対象ファイルを直接書き換える（公開リポ側のコピーに使い、live 本体には使わない）。
"""
import argparse
import re
import sys
import zipfile

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')


def _grab(z, name):
    try:
        return z.read(name).decode('utf-8', 'replace')
    except KeyError:
        return ''


def _is_initialish(v):
    """空欄・イニシャル級（英字と空白/ピリオドのみ・4文字以下）なら公開OK扱い"""
    v = v.strip()
    if not v:
        return True
    return bool(re.fullmatch(r'[A-Za-z][A-Za-z .]{0,3}', v))


def check(path):
    """6項目チェック。問題点のリストを返す（空なら公開OK）"""
    issues = []
    z = zipfile.ZipFile(path)
    core = _grab(z, 'docProps/core.xml')
    for tag in ('dc:creator', 'cp:lastModifiedBy'):
        m = re.search(f'<{tag}>(.*?)</{tag}>', core, re.S)
        val = m.group(1) if m else ''
        mark = 'OK' if _is_initialish(val) else '★実名の可能性'
        print(f'  {tag}: {val!r} [{mark}]')
        if not _is_initialish(val):
            issues.append(f'{tag} が実名の可能性: {val!r}')
    app = _grab(z, 'docProps/app.xml')
    m = re.search(r'<Company>(.*?)</Company>', app, re.S)
    if m and m.group(1).strip():
        print(f'  Company: {m.group(1)!r} [★要確認]')
        issues.append(f'Company: {m.group(1)!r}')
    wbx = _grab(z, 'xl/workbook.xml')
    hidden = re.findall(r'<sheet [^>]*state="(?:hidden|veryHidden)"[^>]*/?>', wbx)
    print(f'  隠しシート: {len(hidden)}件')
    if hidden:
        issues.append(f'隠しシート {len(hidden)}件（中身を目視すること）: {hidden}')
    dn = re.findall(r'<definedName name="([^"]+)"[^>]*>([^<]*)</definedName>', wbx)
    bad_dn = [(n, v) for n, v in dn if re.search(r'[A-Za-z]:\\|Users', v)]
    print(f'  定義名: {len(dn)}件（ローカルパス入り {len(bad_dn)}件）')
    if bad_dn:
        issues.append(f'定義名にローカルパス: {bad_dn}')
    ext = [n for n in z.namelist() if 'externalLink' in n]
    print(f'  外部リンク: {len(ext)}件')
    if ext:
        issues.append(f'外部リンク {len(ext)}件（参照先パスを確認）')
    com = [n for n in z.namelist() if re.search(r'comments\d*\.xml$', n)]
    print(f'  コメント: {len(com)}ファイル')
    if com:
        issues.append(f'コメントあり（中身を目視すること）')
    # 残留入力値: 各シート1〜6行目の値セル（検索語・氏名等の置き忘れ検出の手がかり）
    shared = re.findall(r'<si>(?:<t[^>]*>(.*?)</t>|.*?)</si>', _grab(z, 'xl/sharedStrings.xml'), re.S)
    for sn in sorted(n for n in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n)):
        sx = _grab(z, sn)
        cells = re.findall(r'<c r="([A-Z]+[1-6])"(?:[^>]*? t="(\w+)")?[^>]*>(?:<f>(.*?)</f>)?(?:<v>(.*?)</v>)?</c>', sx)
        vals = []
        for addr, t, f, v in cells:
            if not (f or v):
                continue
            if t == 's' and v.isdigit() and int(v) < len(shared):
                v = shared[int(v)]
            vals.append(f'{addr}={("=" + f) if f else ""}{v[:40]}')
        if vals:
            print(f'  {sn} 上部セル（残留値の目視用）: ' + ' | '.join(vals[:10]))
    z.close()
    return issues


def scrub(path):
    """dc:creator / cp:lastModifiedBy を空欄化（zip 外科手術・その場書き換え）"""
    src = zipfile.ZipFile(path, 'r')
    items = []
    for info in src.infolist():
        data = src.read(info.filename)
        if info.filename == 'docProps/core.xml':
            text = data.decode('utf-8')
            text = re.sub(r'<dc:creator>.*?</dc:creator>', '<dc:creator></dc:creator>', text, flags=re.S)
            text = re.sub(r'<cp:lastModifiedBy>.*?</cp:lastModifiedBy>',
                          '<cp:lastModifiedBy></cp:lastModifiedBy>', text, flags=re.S)
            data = text.encode('utf-8')
        items.append((info, data))
    src.close()
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            zout.writestr(info, data)
    print(f'  空欄化しました: {path}')


def main():
    ap = argparse.ArgumentParser(description='xlsm/xlsx 公開前チェック＆メタデータ空欄化')
    ap.add_argument('files', nargs='+', help='対象ブック（複数可）')
    ap.add_argument('--scrub', action='store_true',
                    help='dc:creator / cp:lastModifiedBy を空欄化（公開用コピーにだけ使う）')
    args = ap.parse_args()
    ng = 0
    for p in args.files:
        print(f'===== {p}')
        if args.scrub:
            scrub(p)
        issues = check(p)
        if issues:
            ng += 1
            print('  ⚠ 公開前に要対処:')
            for i in issues:
                print(f'    - {i}')
        else:
            print('  ✓ 機械チェックは全項目クリア（残留値は上の表示を目視）')
    sys.exit(1 if ng else 0)


if __name__ == '__main__':
    main()
