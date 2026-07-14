# -*- coding: utf-8 -*-
"""publish_check.py — xlsm/xlsx 公開前の個人情報チェック＆メタデータ空欄化を1コマンドに畳む。

2026-07-11 の実名メタデータ公開事故を受けた機械チェック。見る先:
  - docProps/core.xml   作成者・最終更新者・タイトル・説明・キーワード
  - docProps/app.xml    会社・マネージャー
  - docProps/custom.xml カスタムプロパティ（社内テンプレの名残が入りがち）
  - ハイパーリンク      *_rels/*.xml.rels の Target（file:///C:/Users/実名/… が一番残る）
  - xl/vbaProject.bin   VBA プロジェクト内に焼き付いたローカルパス・実名
  - 隠しシート / 定義名 / 外部リンク / コメント / 残留入力値

使い方:
  py publish_check.py <ブック.xlsm>              # チェックのみ（公開可否を判定表示）
  py publish_check.py <ブック.xlsm> --scrub      # dc:creator / cp:lastModifiedBy を空欄化
  py publish_check.py <A.xlsm> <B.xlsm> ...      # 複数まとめて

判定: 実名らしき作成者（イニシャル2〜4文字以外）が残っていれば終了コード1。
--scrub は対象ファイルを直接書き換える（公開リポ側のコピーに使い、live 本体には使わない）。
"""
import argparse
import os
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


def _tag_value(xml, tag):
    """<tag>値</tag> を取り出す（<tag/> の自己完結タグは空欄扱い）。

    開始タグの属性を許容する。<dc:creator xml:space="preserve">山田太郎</dc:creator>
    のような属性付きに当たらないと、値が空と判定されて [OK] 扱いになり、さらに
    実名リストが空のまま vbaProject.bin の実名検索まで素通りして、実名を抱えたまま
    「機械チェックは全項目クリア」＋終了コード0 を返す（PII ゲートとして致命的）。
    """
    m = re.search(rf'<{tag}(?:\s[^>]*)?>(.*?)</{tag}>', xml, re.S)
    return m.group(1) if m else ''


def check(path):
    """公開前チェック。問題点のリストを返す（空なら公開OK）"""
    issues = []
    z = zipfile.ZipFile(path)
    core = _grab(z, 'docProps/core.xml')
    # 実名らしき値はここで集めておき、後で vbaProject.bin の中も同じ語で探す
    realnames = []
    for tag in ('dc:creator', 'cp:lastModifiedBy'):
        val = _tag_value(core, tag)
        mark = 'OK' if _is_initialish(val) else '★実名の可能性'
        print(f'  {tag}: {val!r} [{mark}]')
        if not _is_initialish(val):
            issues.append(f'{tag} が実名の可能性: {val!r}')
            realnames.append(val.strip())
    # タイトル・説明・キーワードは社内文書の名残（担当者名・部署名）が残りやすい
    for tag in ('dc:title', 'dc:description', 'cp:keywords', 'cp:category'):
        val = _tag_value(core, tag).strip()
        if val:
            print(f'  {tag}: {val!r} [★要確認]')
            issues.append(f'{tag} に値が残っています: {val!r}')
    app = _grab(z, 'docProps/app.xml')
    for tag in ('Company', 'Manager'):
        val = _tag_value(app, tag).strip()
        if val:
            print(f'  {tag}: {val!r} [★要確認]')
            issues.append(f'{tag}: {val!r}')
    # カスタムプロパティ（社内テンプレ由来の作成者・部署が入っていることがある）
    custom = _grab(z, 'docProps/custom.xml')
    if custom:
        props = []
        for blk in re.findall(r'<property\b[^>]*>.*?</property>', custom, re.S):
            m = re.search(r'name="([^"]+)"', blk)
            # 値は <vt:lpwstr> 等の型タグに包まれる（vector 入れ子もある）ので
            # タグを剥がして中身のテキストだけを見る
            val = ' '.join(re.sub(r'<[^>]+>', ' ', blk.split('>', 1)[1]).split())
            props.append((m.group(1) if m else '?', val))
        if props:
            print(f'  カスタムプロパティ: {len(props)}件 [★要確認]')
            for n, v in props:
                print(f'      {n} = {v!r}')
            issues.append(f'カスタムプロパティ {len(props)}件: {[n for n, _ in props]}')
    # ハイパーリンク: file:///C:/Users/実名/… が一番実名パスの残る場所
    # （ファイル一覧.xlsm 系はハイパーリンクだらけ）。sheet だけでなく全 .rels を見る。
    linkhits = []
    for n in z.namelist():
        if not n.endswith('.rels'):
            continue
        for tgt in re.findall(r'Target="([^"]+)"', _grab(z, n)):
            if re.match(r'(?:file:///)?[A-Za-z]:[\\/]', tgt) or tgt.lower().startswith('file:///'):
                linkhits.append((n, tgt))
    print(f'  ハイパーリンク等のローカル絶対パス: {len(linkhits)}件')
    if linkhits:
        for n, tgt in linkhits[:10]:
            print(f'      {n}: {tgt[:100]}')
        if len(linkhits) > 10:
            print(f'      … 他 {len(linkhits) - 10}件')
        issues.append(f'ハイパーリンクにローカル絶対パス {len(linkhits)}件（実名パスが露出します）')
    # VBA プロジェクト（バイナリ）に焼き付いたローカルパス・実名。
    # 文字列は CP932 と UTF-16LE の両方でありうるので生バイトで両方探す。
    try:
        vba = z.read('xl/vbaProject.bin')
    except KeyError:
        vba = b''
    if vba:
        vba_hits = []
        needles = [r'C:\Users', r'C:/Users'] + [n for n in realnames if n]
        for needle in needles:
            for enc in ('cp932', 'utf-16-le'):
                try:
                    b = needle.encode(enc)
                except UnicodeEncodeError:
                    continue
                if b and b in vba:
                    vba_hits.append(f'{needle}（{enc}）')
        print(f'  vbaProject.bin 内のローカルパス/実名: {len(vba_hits)}件')
        if vba_hits:
            for h in vba_hits:
                print(f'      {h}')
            issues.append(f'vbaProject.bin に実名/ローカルパス: {vba_hits}'
                          '（VBAコード中のパス文字列を書き換えて作り直すこと）')
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


def _blank_tag(text, tag):
    """<tag>値</tag> と <tag/> の両方を空欄化。(新text, 空欄化した値 or None) を返す。

    自己完結タグ <dc:creator/> に当たらない正規表現だと「何もしていないのに
    空欄化しました」と報告してしまうため、両方の形を見る。
    """
    # 開始タグの属性を許容する（_tag_value と同じ理由）。
    # <dc:creator xml:space="preserve">山田太郎</dc:creator> のような属性付きに
    # 当たらないと、実名が残っているのに「空欄化する項目はありませんでした」と
    # 報告して何もしないまま公開GOになる。
    m = re.search(rf'<{tag}(?:\s[^>]*)?>(.*?)</{tag}>', text, re.S)
    if m:
        if not m.group(1):
            return text, None                      # 既に空欄
        return text.replace(m.group(0), f'<{tag}></{tag}>', 1), m.group(1)
    if re.search(rf'<{tag}(?:\s[^>]*)?/>', text):
        return text, None                          # 自己完結タグ＝既に空欄
    return text, None


def scrub(path):
    """dc:creator / cp:lastModifiedBy を空欄化（zip 外科手術）。

    その場で ZipFile(path, 'w') を開くと元ファイルを即座に切り詰めるため、途中で
    例外・中断が起きるとブックが復元不能になる。一時ファイルに書き切ってから
    os.replace で原子的に差し替える。
    """
    blanked = []
    src = zipfile.ZipFile(path, 'r')
    items = []
    for info in src.infolist():
        data = src.read(info.filename)
        if info.filename == 'docProps/core.xml':
            text = data.decode('utf-8')
            for tag in ('dc:creator', 'cp:lastModifiedBy'):
                text, old = _blank_tag(text, tag)
                if old is not None:
                    blanked.append(f'{tag}={old!r}')
            data = text.encode('utf-8')
        items.append((info, data))
    src.close()

    if not blanked:
        print('  空欄化する項目はありませんでした（dc:creator / cp:lastModifiedBy は既に空欄）。')
        return

    tmp = path + '.scrub.tmp'
    try:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for info, data in items:
                zout.writestr(info, data)
        os.replace(tmp, path)     # 書き切ってから差し替える（途中で落ちても元は無傷）
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    print(f'  空欄化しました: {path}')
    for b in blanked:
        print(f'    - {b} → 空欄')


def main():
    ap = argparse.ArgumentParser(description='xlsm/xlsx 公開前チェック＆メタデータ空欄化')
    ap.add_argument('files', nargs='+', help='対象ブック（複数可）')
    ap.add_argument('--scrub', action='store_true',
                    help='dc:creator / cp:lastModifiedBy を空欄化（公開用コピーにだけ使う）')
    args = ap.parse_args()
    ng = 0
    for p in args.files:
        print(f'===== {p}')
        # 必ず check を先に回す。scrub は dc:creator / cp:lastModifiedBy を空欄化するが、
        # vbaProject.bin の実名検索はその作成者名を手がかり(needle)にしている。
        # 先に scrub すると手がかりが消え、bin に実名が残っていても「0件・全項目クリア・
        # 終了コード0」と言い切る＝実名を抱えたまま公開GOを出す
        # （2026-07-14 実弾で確認。7/11 の公開事故の再演になるところだった）
        issues = check(p)
        if issues:
            ng += 1
            print('  ⚠ 公開前に要対処:')
            for i in issues:
                print(f'    - {i}')
        else:
            print('  ✓ 機械チェックは全項目クリア（残留値は上の表示を目視）')
        if args.scrub:
            scrub(p)
            print('  ※ 空欄化しました。上の検査結果は空欄化「前」の状態です。')
            print('     空欄化後の状態を確認するには --scrub なしでもう一度実行してください。')
    sys.exit(1 if ng else 0)


if __name__ == '__main__':
    main()
