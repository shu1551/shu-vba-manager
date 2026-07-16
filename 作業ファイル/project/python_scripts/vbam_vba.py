# -*- coding: utf-8 -*-
"""vbam_vba.py — vba_manager 分割パート: VBA コマンド実装（list/get/replace/export/フォーム lint 等）

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

from vbam_core import *  # noqa: F401,F403
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
    seen = {}       # 小文字名 -> [行番号, ...]（照合用）
    disp = {}       # 小文字名 -> 最初に現れた綴り（表示用）
    for idx, line in enumerate(norm_text.split('\n'), 1):
        m = sub_pattern.match(line)
        if m:
            # VBA の識別子は大小文字を区別しない。Calc と calc は「同名」で
            # コンパイルエラーになるため、照合は小文字化して行う
            # （区別して数えると「重複なし・取り込み可」と誤って報告する）。
            # ただし報告は元の綴りで出す（キーごと小文字に潰すと表示が化ける）
            key = m.group('name').lower()
            seen.setdefault(key, []).append(idx)
            disp.setdefault(key, m.group('name'))
    return {disp[key]: lns for key, lns in seen.items() if len(lns) > 1}


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
        # 「起動していません」は正常な状態報告であって異常ではない。
        # ここだけ None を返すと終了コードが 1 になり他の分岐と不揃いになる
        return True
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
    # 参照するだけのコマンド。閉じたブックを自動で開く場合に Workbook_Open /
    # Auto_Open を発火させないよう読み取り専用＋イベント無効で開く
    # （既に開いているブックには影響しない）
    xl, default_wb = get_workbook(target_file, load_addins=load_addins, readonly=True)

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
    # 参照するだけのコマンド → 読み取り専用で開く（Workbook_Open を起こさない）
    xl, default_wb = get_workbook(target_file, load_addins=load_addins, readonly=True)

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


def _narrow_proc_range(cm, start, count):
    """ProcStartLine/ProcCountLines の領域を「実体だけ」に絞る。

    ProcCountLines の領域には前後の空行に加え、次プロシージャの宣言行が
    食い込むことがある（1行完結 Sub「Sub X(): Call Main: End Sub」の直後など）。
    get(_extract_proc) と replace-procedure は同じ条件でその行を落としているが、
    delete-procedure だけが生の start/count を DeleteLines に渡しており、
    隣のプロシージャの宣言行ごと消して「削除完了」と報告していた。

    戻り値: (実効start, 実効count)
    """
    lines = cm.Lines(start, count).replace('\r\n', '\n').split('\n')
    lead = 0
    while lead < len(lines) and lines[lead].strip() == '':
        lead += 1
    end = len(lines)
    while end - lead > 1:
        last = lines[end - 1].strip()
        if last == '':
            end -= 1
        elif (re.match(
                r'^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+',
                last, re.IGNORECASE)
              and not re.search(r'\bEnd\s+(?:Sub|Function)\b', last, re.IGNORECASE)):
            end -= 1
        else:
            break
    return start + lead, end - lead


def _inline_body_after_decl(raw, decl_end):
    """宣言行に同居する本体（1行完結 Sub）を取り出す。無ければ空文字。

    「Sub X(): Call Main: End Sub」の "Call Main: End Sub" の部分を返す。
    引数リストの括弧は対応を取って読み飛ばす（既定値の文字列に ':' や ')' が
    入っていても誤らないように、文字列リテラルの中は数えない）。
    """
    s = raw[decl_end:]
    i = 0
    if i < len(s) and s[i] == '(':
        depth = 0
        in_str = False
        while i < len(s):
            c = s[i]
            if c == '"':
                in_str = not in_str
            elif not in_str:
                if c == '(':
                    depth += 1
                elif c == ')':
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
            i += 1
    rest = s[i:]
    # ':' を探す前に行末コメントを落とす。落とさないと
    # 「Sub X() ' 例: Run "整形" してから」のようなコメント内のコロンを
    # 本体の区切りと誤認し、コメント文をコードとして走査してしまう
    # （コメント内の Call/Run が偽の呼び出しとして call-graph に載る）
    rest = _strip_vba_comment(rest)
    # 引数リストの後ろに ':' があれば、それ以降が同じ行に書かれた本体
    # （'As String' のような戻り型指定はコロンの手前にある）
    if ':' not in rest:
        return ''
    return rest.split(':', 1)[1]


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

    # コードを読むだけ → 読み取り専用で開く（Workbook_Open を起こさない）
    xl, wb = get_workbook(target_file, readonly=True)

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

    # コードファイルに2本以上のプロシージャが入っていたら弾く。
    # get は複数のプロシージャを1ファイルに連結して書ける（cmd_get）が、
    # replace-procedure は先頭1本の「領域」に全文を流し込むため、2本目以降は
    # モジュール内の重複定義になりコンパイルエラーになる。MCP 経由は常に -y で
    # 確認プロンプトも挟まらないので、文言の注意ではなく機械的に止める。
    decl_names = []
    for ln in code_lines:
        s = ln.strip()
        if not s or s.startswith("'") or re.match(r'^Rem\b', s, re.IGNORECASE):
            continue
        md = pattern.match(s)
        if md:
            decl_names.append(md.group(1))
    if len(decl_names) > 1:
        print(f"エラー: コードファイルに {len(decl_names)} 本のプロシージャが入っています:")
        for dn in decl_names:
            print(f"  - {dn}")
        print("  replace-procedure は1本ずつ書き戻してください")
        print("  （先頭の1本しか対象にならず、2本目以降は重複定義になります）。")
        print("  まとめて差し替えるなら replace-module を使ってください。")
        return False

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
    # 末尾は空行だけでなく「次プロシージャの宣言行」も削除範囲から外す。
    # ProcCountLines の領域には次の宣言行が食い込むことがあり（1行完結 Sub の直後など）、
    # get(_extract_proc) と new_code のサニタイザは同じ条件でその行を落とす。
    # ここだけ削除範囲に含めると、消した宣言行が new_code から復元されず
    # 次のプロシージャが宣言を失って壊れる（両者を対称に保つ）。
    end = len(old_all)
    while end - lead > 1:
        last = old_all[end - 1].strip()
        if last == '':
            end -= 1
        elif (pattern.match(last)
              and not re.search(r'\bEnd\s+(?:Sub|Function)\b', last, re.IGNORECASE)):
            end -= 1
        else:
            break
    eff_start = proc_start + lead
    eff_count = end - lead

    import difflib
    old_lines = old_all[lead:end]
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
                # ここで終端を確定する（随伴 Attribute 行までが置換対象）。
                # 文字列リテラルだけでなくコメントも落とす。落とさないと
                # 「Public Sub 印刷実行()  ' 途中で Exit せず End Sub まで通す」のような
                # 行末コメントを1行完結Subと誤判定し、置換範囲が宣言行だけになって
                # 旧本体が残る（新旧の本体が並んで二重化・構文破壊）
                no_str = _strip_vba_comment(re.sub(r'"[^"]*"', '""', line))
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
        # Attribute行の挿入位置は「宣言の実体が終わった直後」。
        # 宣言が行継続 " _" で折り返している場合（Sub X(ByVal a As Long, _ ）に
        # 1行目の直後へ入れると Attribute が継続行の前に割り込んで構文破壊になる。
        # 読み取り側（上の attr_block 収集）と対称に、継続行を読み飛ばしてから入れる
        insert_at = 1
        while insert_at < len(new_lines) and re.search(r'\s_\s*$', new_lines[insert_at - 1]):
            insert_at += 1
        for ai, al in enumerate(attr_block):
            new_lines.insert(insert_at + ai, al)

        # .bas 内の対象プロシージャを置換（Sub宣言行から End Sub まで）
        bas_lines[sub_line_idx:proc_end_idx + 1] = new_lines
        new_bas = '\r\n'.join(bas_lines)

        # --force で CP932 に無い文字（絵文字・機種依存文字等）が混ざると
        # ここで UnicodeEncodeError になる。生の例外で落とすと一時 .bas が残るので
        # 握りつぶさず・分かる形で報告して片付ける
        try:
            encoded_bas = new_bas.encode('cp932')
        except UnicodeEncodeError as ex:
            bad = new_bas[ex.start:ex.end]
            _remove_export_artifacts(tmp_bas)
            print(f"エラー: 置換コードに CP932（Shift-JIS）で扱えない文字があります: {bad!r}")
            print("  VBA モジュールは CP932 で保存されます。該当文字を通常の文字に置き換えてください。")
            return False

        with open(tmp_bas, 'wb') as f:
            f.write(encoded_bas)

        # Remove + Import（例外時も DisplayAlerts を戻し、一時ファイルを残さない）
        xl.DisplayAlerts = False
        removed = False
        imported = False
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
            # ここから先の失敗は「保存の失敗」であって「モジュールの消失」ではない。
            # この印を立てずに復旧へ落ちると、正しく入った新モジュールにバックアップを
            # 重ね Import して、ツール自身が連番モジュールを作ってしまう（2026-07-14 発見）
            imported = True
            _save_with_retry(wb)
        except ModuleNameCollisionError as ex:
            # コードは連番付き別名側に生きている。バックアップ再 Import は三重化するので禁止
            _print_collision_guidance(ex, module_name, module_backup)
            return False
        except Exception as ex:
            print(f"エラー: 置換中に失敗しました: {ex}")
            if imported:
                # Import は成功済み＝消えていない。再 Import は連番モジュールを生むだけ
                _print_save_failed_guidance(module_name)
            elif removed:
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
    # コードファイルに複数本入っていることがある（get は3本以上を連結して
    # 1つの _last_proc.vba に書く）。先頭1本だけ見て重複検査すると、2本目以降が
    # 既存と同名でも素通りし、同一モジュール内に二重定義されてコンパイル不能になる
    decl_names = re.findall(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+([^\s\(]+)', new_code,
        re.MULTILINE | re.IGNORECASE)
    if not decl_names:
        print("エラー: コードに Sub/Function 宣言が見つかりません")
        return False
    proc_name = decl_names[0]
    # コードファイル自身の中の同名重複も弾く（VBA は大小文字を区別しない）
    _lowered = [n.lower() for n in decl_names]
    for n in decl_names:
        if _lowered.count(n.lower()) > 1:
            print(f"エラー: コードファイル内に同名のプロシージャが複数あります: {n}")
            return False

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
    # 同一モジュール内はコンパイルエラーになるため必ず止める）。
    # コードファイル内の全宣言について検査する（先頭1本だけでは 2本目以降が素通りする）
    exists = False
    for _n in decl_names:
        try:
            cm.ProcStartLine(_n, 0)
            proc_name = _n          # 実際に衝突した名前を報告する
            exists = True
            break
        except Exception:
            continue
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
    # 領域に食い込んだ「次プロシージャの宣言行」を削除範囲から外す。
    # 外さないと 1行完結 Sub の直後のプロシージャが宣言を失って壊れる
    # （get / replace-procedure は同じ絞り込みを既に持っている）
    start, count = _narrow_proc_range(cm, start, count)
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

    actual_name = comp.Name
    wb.VBProject.VBComponents.Remove(comp)
    # VBE の Remove は遅延完了することがある（対象モジュールのコードが実行中など）。
    # 消える前に Save すると「削除完了」と報告しながらファイルには残り、
    # あとから遅延 Remove だけが効いて未保存のまま消える——という食い違いが起きる
    pythoncom.PumpWaitingMessages()
    if not _wait_component_gone(wb, actual_name):
        print(f"エラー: モジュール '{actual_name}' の Remove が完了しませんでした"
              f"（そのモジュールのコードが実行中の可能性があります）。", file=sys.stderr)
        print("  保存はしていません。Excel 側の実行中マクロ・開いているフォームを"
              "閉じてから、もう一度実行してください。", file=sys.stderr)
        if backup:
            print(f"  モジュールのバックアップ: {backup}", file=sys.stderr)
        return False

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
            imported = False
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
                # ここから先の失敗は「保存の失敗」であって「モジュールの消失」ではない
                imported = True
                _save_with_retry(wb)
            except ModuleNameCollisionError as ex:
                # コードは連番付き別名側に生きている。バックアップ再 Import は三重化するので禁止
                _print_collision_guidance(ex, module_name, module_backup)
                return False
            except Exception as ex:
                print(f"エラー: 置換中に失敗しました: {ex}")
                if imported:
                    # Import は成功済み＝消えていない。再 Import は連番モジュールを生むだけ
                    _print_save_failed_guidance(module_name)
                elif removed:
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
            # その行（＋随伴 Attribute 行）でブロックを閉じる。
            # 文字列リテラルだけでなくコメントも落とす。落とさないと行末コメントの
            # 「End Sub」で誤判定し、ブロックが宣言行だけで閉じて本体が次の
            # プロシージャに吸収され、reorder で本体が隣へ道連れになる
            no_str = _strip_vba_comment(re.sub(r'"[^"]*"', '""', lines[i]))
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
    imported = False
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
        # ここから先の失敗は「保存の失敗」であって「モジュールの消失」ではない。
        # success は最後の Save の後にしか立たないので、復旧判定には使えない
        imported = True
        _save_with_retry(wb)
        success = True
    except ModuleNameCollisionError as ex:
        # コードは連番付き別名側に生きている。バックアップ再 Import は三重化するので禁止
        _print_collision_guidance(ex, module_name, module_backup, err=True)
        return False
    except Exception as ex:
        # Remove 成功後に Import が失敗するとモジュールが消えたままになる。
        # replace-module と同じく、モジュールバックアップからの自動復旧を先に試みる
        if imported:
            # Import は成功済み＝消えていない。再 Import は連番モジュールを生むだけ
            print(f"エラー: 並べ替え後の保存に失敗しました: {ex}", file=sys.stderr)
            _print_save_failed_guidance(module_name, err=True)
        elif removed:
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
        # 「Remove 済みで Import 失敗」のときだけ復旧用に残し、それ以外は掃除する。
        # Import 済み（保存だけ失敗）はブック側に新コードが入っているので残す必要はない
        if not (removed and not imported) and os.path.exists(tmp_bas):
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

    # 書き出すだけ（ブックは変更しない） → 読み取り専用で開く
    xl, wb = get_workbook(target_file, readonly=True)

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

    # 書き出すだけ（ブックは変更しない） → 読み取り専用で開く
    xl, wb = get_workbook(target_file, readonly=True)
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
    # `or 30` だと --max 0（件数だけ見たい）が偽値で既定に化ける（is None 判定にする）
    _m = getattr(args, 'max_hits', None)
    limit = 30 if _m is None else int(_m)
    if limit < 0:
        print("エラー: --max は 0 以上で指定してください（0 は件数のみ表示）")
        return False
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
    # Export して読むだけ（ブックは変更しない） → 読み取り専用で開く
    xl, wb = get_workbook(target_file, readonly=True)

    shortcuts = []
    unread = []      # Export に失敗して中身を見られなかったモジュール

    for comp in wb.VBProject.VBComponents:
        if comp.Type not in (1, 2, 3, 100):
            continue

        tmp_file = os.path.join(SCRIPT_DIR, f"_tmp_sc_{comp.Name}.bas")
        try:
            comp.Export(tmp_file)
            with open(tmp_file, 'rb') as f:
                content = f.read().decode('cp932', errors='replace')
        except Exception as ex:
            # 黙って飛ばすと最後に「ショートカットなし」と言い切ってしまう。
            # 読めなかったモジュールは必ず報告する
            unread.append((comp.Name, _com_error_text(ex)))
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
        out = {"success": True, "file": wb.Name, "shortcuts": shortcuts}
        if unread:
            out["unread_modules"] = [{"module": m, "error": e} for m, e in unread]
        print(json.dumps(out, ensure_ascii=False), file=sys.stdout)
        return True

    if unread:
        print(f"⚠ 次のモジュールは読めませんでした（一覧に反映されていません・{len(unread)}件）:")
        for m, e in unread:
            print(f"  - {m}: {e}")

    if not shortcuts:
        if unread:
            print("読めたモジュールの範囲では、ショートカットキーの設定は見つかりませんでした。")
        else:
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


def _collect_book_inventory(xl, wb, include_vba=True, quiet=False):
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
    sc_unread = []      # Export に失敗して読めなかったモジュール（黙って落とさない）
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
        except Exception as ex:
            # 握りつぶすと「ショートカットなし」と誤って断言することになる
            sc_unread.append((comp.Name, _com_error_text(ex)))
        finally:
            _remove_export_artifacts(tmp)
    inv['shortcuts'] = shortcuts
    inv['shortcuts_unread'] = sc_unread
    if sc_unread and not quiet:
        # MCP では stderr と stdout が同じバッファなので、--json 時にここへ書くと
        # JSON の前にゴミが混ざって呼び出し側の json.loads が落ちる。
        # --json の呼び出し元は quiet=True にし、内容は inv['shortcuts_unread'] で受け取る
        print(f"⚠ ショートカット走査で読めなかったモジュール（{len(sc_unread)}件）: "
              + '、'.join(f"{m}({e})" for m, e in sc_unread), file=sys.stderr)

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
        print("  シート側だけなら sheet-info / snapshot 等の目コマンドは使えます。")
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
    """呼び出し関係の解析本体（call-graph / impact 共用・COM 不要の純粋処理）。

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
            close_after = False
            scan = raw            # 走査対象（表示は raw のまま＝元の行を見せる）
            dm = defn_pat.match(raw)
            if dm:
                cur = dm.group(1)
                # 1行完結 Sub「Sub X(): Call Main: End Sub」は宣言行に本体が同居する。
                # ここで無条件に continue すると Call が走査されず、呼ばれている Main が
                # 「どこからも呼ばれていない」に誤って載り、誤記（Call 印刷実効）も
                # 未解決Callとして検出されない（call-graph の検出器がこの形だけ素通り）。
                inline = _inline_body_after_decl(raw, dm.end())
                if not inline.strip():
                    continue
                scan = inline
                # 同じ行で End Sub まで来ているなら、この行を見終えた時点で閉じる
                close_after = bool(re.search(
                    r'\bEnd\s+(?:Sub|Function|Property)\b',
                    _strip_vba_comment(inline), re.IGNORECASE))
            elif end_pat.match(raw):
                cur = None
                continue
            # 文字列リテラルを潰してからコメントを落とす（' の誤爆防止）
            line = re.sub(r'"[^"]*"', '""', scan).split("'")[0]
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
            for rm_ in run_pat.finditer(_strip_vba_comment(scan)):
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
            if close_after:
                # 1行完結 Sub は宣言行で閉じている。ここで戻さないと、後続行の
                # 呼び出しがこの Sub の名で計上され続ける
                cur = None

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
    """参考所見スキャン（COM 不要の純粋処理・事実の列挙のみ）。

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
    # `or 200` だと --max 0（件数だけ見たい）が偽値で既定に化ける（is None 判定にする）
    _m = getattr(args, 'max_hits', None)
    max_hits = 200 if _m is None else int(_m)
    if max_hits < 0:
        print("エラー: --max は 0 以上で指定してください（0 は件数のみ表示）")
        return False

    # コードを読むだけ → 読み取り専用で開く（Workbook_Open を起こさない）
    xl, wb = get_workbook(target_file, readonly=True)
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
        """ダイアログのタイトル＋本文（Static コントロールの文字）を採取して報告に使う。

        タイトルは「どのダイアログが開いたか」の一番強い証拠
        （例: セルの書式設定 / 検索と置換）。MsgBox 系はタイトルが
        「Microsoft Excel」等で情報が薄いので本文と併記する。
        """
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
        title = ''
        try:
            title = win32gui.GetWindowText(hwnd).strip()
        except Exception:
            pass
        body = ' / '.join(parts)
        if title:
            return f"タイトル「{title}」 本文: {body}" if body else f"タイトル「{title}」"
        return body

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
    # count は「窓が実際に消えたことを確認した数」。PostMessage は非同期で、
    # 送っただけでは閉じたことにならない（ボタンID解決が外れる／WM_CLOSE を
    # 無視するダイアログでは閉じない）。従来は送信した時点で数えて「解除しました」と
    # 報告し、その hwnd を二度と再送しなかったため、安全弁が仕事をせずハングし続けた。
    state = {'count': 0, 'last': '', 'unclosed': []}
    pending = {}                           # hwnd -> {tries, last_ts, text}
    _RESEND_SEC = 0.6                      # 消えなければこの間隔で再送する

    def _loop():
        while not stop_evt.is_set():
            targets = []

            def _cb(hwnd, _unused):
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return
                    # #32770=MsgBox/InputBox等の標準ダイアログ、
                    # bosa_sdm_XL9=Excel内蔵ダイアログ(xlDialog*/セルの書式設定等)、
                    # NUIDialog=Office描画ダイアログ。いずれもExcel PID所有のモーダルとして
                    # 放置するとxl.Runが無言ハングする(2026-07-12実害: xlDialogFormatNumberを
                    # 開くマクロでrun-macroが40秒超ブロック、#32770限定だったため素通り)
                    if win32gui.GetClassName(hwnd) not in (
                            '#32770', 'bosa_sdm_XL9', 'NUIDialog'):
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
            # 消えた hwnd ＝ 本当に閉じられたもの。ここで初めて「解除した」と数える
            for hwnd in [h for h in pending if h not in targets]:
                info = pending.pop(hwnd)
                state['count'] += 1
                if info.get('text'):
                    state['last'] = info['text']
            now = time.time()
            for hwnd in targets:
                info = pending.get(hwnd)
                if info is not None and now - info['last_ts'] < _RESEND_SEC:
                    continue                       # 送ったばかり。反応を待つ
                try:
                    txt = _dialog_text(hwnd)
                    if win32gui.GetClassName(hwnd) == '#32770':
                        cid = _resolve_button_id(hwnd)
                        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, cid, 0)
                    else:
                        # Excel内蔵/Office描画ダイアログは子がWin32 Buttonでないため
                        # ボタンID解決が効かない。WM_CLOSE(=×ボタン)がキャンセル相当
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    if info is None:
                        pending[hwnd] = {'tries': 1, 'last_ts': now, 'text': txt}
                    else:
                        info['tries'] += 1
                        info['last_ts'] = now
                        if txt:
                            info['text'] = txt
                except Exception:
                    pass
            stop_evt.wait(0.15)
        # 監視終了時点で残っているもの＝送っても閉じなかったダイアログ。
        # 「解除しました」と混ぜず、閉じられなかった事実として別に報告する
        for info in pending.values():
            state['unclosed'].append(info.get('text') or '(本文なし)')

    th = threading.Thread(target=_loop, daemon=True)
    th.start()

    class _Watcher:
        def stop(self):
            stop_evt.set()
            # ループが抜けきるのを待つ。待たずに戻ると、直後に unclosed を読む
            # 呼び出し側が「閉じられなかったダイアログ」を取りこぼす（競合）
            try:
                th.join(timeout=1.0)
            except Exception:
                pass

        @property
        def count(self):
            return state['count']

        @property
        def last_text(self):
            return state['last']

        @property
        def unclosed(self):
            """送っても閉じなかったダイアログの本文（安全弁が効かなかった証拠）"""
            return state['unclosed']

    return _Watcher()


def _dialog_watcher_note(watcher, mode):
    """ダイアログを自動解除したときの人向けの注記（無ければ空文字）。

    「黙って握りつぶした」にしないための報告。何が出て、どう閉じたかを一行で残す。
    """
    if watcher is None:
        return ""
    if not getattr(watcher, 'count', 0):
        # 1件も閉じられていなくても、閉じられずに残ったダイアログがあれば報告する
        stuck0 = list(getattr(watcher, 'unclosed', ()) or ())
        if stuck0:
            return ("⚠ 実行中にダイアログを検出しましたが、閉じられませんでした"
                    f"（{len(stuck0)}件・Excel 側に残っている可能性があります）: "
                    + " / ".join(s[:60] for s in stuck0[:3]))
        return ""
    how = "指定ボタンで応答" if mode else "安全側（キャンセル優先）で自動解除"
    body = watcher.last_text or "(本文なし)"
    n = watcher.count
    msg = (f"⚠ 実行中にダイアログを{n}件検出し、{how}しました。"
           f"マクロがメッセージを出しています → 内容: {body}")
    # 送っても閉じなかったもの＝安全弁が効かなかった証拠。「解除しました」と
    # 混ぜず、閉じられなかった事実として必ず別に出す（黙って成功にしない）
    stuck = list(getattr(watcher, 'unclosed', ()) or ())
    if stuck:
        msg += (f"\n⚠ うち {len(stuck)}件は閉じられませんでした"
                "（Excel 側に残っている可能性があります）: "
                + " / ".join(s[:60] for s in stuck[:3]))
    return msg


def dialog_safe(cmd_func):
    """cmd_* を「実行中のダイアログを安全側で自動解除する」に変える decorator。

    セルの書き換え・クリア・行列削除・並べ替え・置換は Worksheet_Change 等の
    イベントマクロを同期発火させる。そのマクロが MsgBox を出すと COM 呼び出しが
    そこでブロックし、人が手で閉じるまでコマンドが無言でハングする（2026-07-11 実害。
    write-range と run-macro では常設済みだったが、他の書き込み系には無かった）。
    検出したダイアログは終了後に必ず報告する（黙って握りつぶさない）。
    """
    import functools

    @functools.wraps(cmd_func)
    def wrapper(args):
        try:
            target_file, _rest = parse_target_and_rest(getattr(args, 'posargs', []) or [])
            xl, _wb = get_workbook(target_file)
        except Exception:
            # ブック解決に失敗した場合は元の関数に委ね、そちらのエラーを出させる
            return cmd_func(args)
        watcher = _start_dialog_watcher(xl)
        try:
            return cmd_func(args)
        finally:
            try:
                watcher.stop()
            except Exception:
                pass
            note = _dialog_watcher_note(watcher, None)
            if note:
                print(note, file=sys.stderr)
    return wrapper


def _project_book_name(xl, p):
    """VBProject から Application.Run 修飾に使うブック名を得る。

    p.Filename は未保存ブックだと例外／空になる。その場合は開いているブックを
    走査して同じプロジェクトのブック名（Book1 等）を拾う。Run はブック名で修飾
    できるので、未保存ブックでも取り違えずに名指しできる。見つからなければ None。
    """
    try:
        fn = p.Filename
    except Exception:
        fn = None
    if fn:
        return os.path.basename(fn)
    try:
        pname = p.Name
    except Exception:
        return None
    try:
        for w in xl.Workbooks:
            try:
                if not w.Path and w.VBProject.Name == pname:
                    return w.Name
            except Exception:
                continue
    except Exception:
        pass
    return None


def _find_macro_owner_books(xl, wb, macro_name):
    """macro_name を宣言しているプロジェクトを全部探す。

    戻り値: (対象ブック名 or None, 他候補[(ブック名 or None, プロジェクト名)])

    VBProjects の列挙順で先勝ちすると、アドインや PERSONAL.XLSB に同名 Sub が
    あるとき作業ブックではなくそちら側を実行してしまう。対象ブック（get_workbook が
    返した wb）に有ればそれを最優先し、他にも同名があれば警告できるよう全部返す。
    """
    pattern = re.compile(
        r'^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?'
        r'(?:Sub|Function)\s+' + re.escape(macro_name) + r'\b',
        re.IGNORECASE | re.MULTILINE
    )
    try:
        target_name = wb.Name
    except Exception:
        target_name = None

    target_hit = None
    others = []
    try:
        projects = list(xl.VBE.VBProjects)
    except Exception as ex:
        print(f"[DEBUG] Failed to enumerate VBProjects: {ex}", file=sys.stderr)
        return None, []

    for p in projects:
        hit = False
        try:
            for comp in p.VBComponents:
                try:
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0 and pattern.search(cm.Lines(1, cm.CountOfLines)):
                        hit = True
                        break
                except Exception:
                    # 読めないモジュールは飛ばすが、探索自体は止めない
                    continue
        except Exception:
            continue
        if not hit:
            continue
        bname = _project_book_name(xl, p)
        try:
            pname = p.Name
        except Exception:
            pname = '?'
        if bname and target_name and bname.lower() == target_name.lower():
            target_hit = bname
        else:
            others.append((bname, pname))
    return target_hit, others


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
        # 対象ブック（アクティブブック or --target）を最優先で探す。
        # 列挙順の先勝ちだと、アドイン/PERSONAL 側の同名マクロを実行してしまう
        target_hit, others = _find_macro_owner_books(xl, wb, macro_name)
        found_wb = target_hit
        if target_hit:
            if others:
                dup = '、'.join(
                    (b or f'{pn}（ブック名不明）') for b, pn in others
                )
                print(f"⚠ 同名マクロが他にもあります: {dup} → 対象ブック "
                      f"'{target_hit}' 側を実行します", file=sys.stderr)
        elif others:
            named = [(b, pn) for b, pn in others if b]
            if named:
                found_wb = named[0][0]
                if len(others) > 1:
                    dup = '、'.join(
                        (b or f'{pn}（ブック名不明）') for b, pn in others
                    )
                    print(f"⚠ 同名マクロが複数のブックにあります: {dup} → "
                          f"'{found_wb}' 側を実行します。意図が違う場合は "
                          f"\"ブック名!マクロ名\" で名指ししてください", file=sys.stderr)
            else:
                # ブック名で修飾できない（プロジェクトの帰属が取れない）ときだけ
                # 修飾なしの直接実行に任せる。黙って別ブックを掴まない
                print(f"[WARNING] Macro '{macro_name}' はブック名を特定できない"
                      f"プロジェクトにあります。修飾なしで実行します。", file=sys.stderr)

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
        # ツール側では切っていないが、実行したマクロが DisplayAlerts=False を立てたまま
        # 落ちている場合がある。そのままだとユーザーの Excel セッションに残り、以後の
        # 手動操作で保存確認などの警告が出なくなるため、必ず有効に戻す
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

    # ダイアログ対策は既定で常設（run-macro / write-range と同じ）。test は任意の
    # テスト Sub を Application.Run する＝最も VBA を発火させる経路で、テスト内の
    # MsgBox で無言ハングする（--auto-dialog 明示時だけ監視、では守れない）。
    # 明示時はそのボタンで応答、省略時は安全解除（キャンセル優先）。
    _auto_dialog = getattr(args, 'auto_dialog', None)
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
            # 失敗して抜ける経路でも、検出したダイアログは必ず報告する
            # （通常の finally は報告するのに、ここだけ黙って捨てていた）
            _note = _dialog_watcher_note(_dlg_watcher, _auto_dialog)
            if _note:
                print(_note)
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
            # 解除したダイアログがあれば必ず本文で報告する（無言で握りつぶさない）
            _note = _dialog_watcher_note(_dlg_watcher, _auto_dialog)
            if _note:
                print(_note)
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




__all__ = [
    '_AUTO_EXEC_PREFIXES',
    '_AUTO_EXEC_STD',
    '_all_procedure_names',
    '_analyze_calls',
    '_check_bas_one',
    '_collect_book_inventory',
    '_com_error_text',
    '_dialog_watcher_note',
    'dialog_safe',
    '_extra_code_scans',
    '_extract_proc',
    '_find_consecutive_dup_lines',
    '_find_duplicate_procedures',
    '_find_macro_owner_books',
    '_inline_body_after_decl',
    '_inventory_or_explain',
    '_narrow_proc_range',
    '_parse_module_blocks',
    '_project_book_name',
    '_select_addin_project',
    '_start_dialog_watcher',
    '_strip_vba_comment',
    '_suggest_similar',
    '_write_module',
    'cmd_add_module',
    'cmd_add_procedure',
    'cmd_call_graph',
    'cmd_check',
    'cmd_check_bas',
    'cmd_code_replace',
    'cmd_delete_module',
    'cmd_delete_procedure',
    'cmd_diag',
    'cmd_docs',
    'cmd_export_all',
    'cmd_export_module',
    'cmd_get',
    'cmd_grep',
    'cmd_impact',
    'cmd_list',
    'cmd_list_backups',
    'cmd_list_modules',
    'cmd_list_open',
    'cmd_list_shortcuts',
    'cmd_reorder_macro',
    'cmd_replace_module',
    'cmd_replace_procedure',
    'cmd_restore',
    'cmd_run_macro',
    'cmd_setup_check',
    'cmd_test',
]
