#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""vba_mcp_server.py — vba_manager を MCP サーバー化する薄い窓口。

実体は vba_manager.py の関数群そのもの（この層は判断をしない）。常駐プロセスが
get_workbook の接続キャッシュを持ち続けるため、CLI の「1コマンド毎の COM 再接続」
が消える。いわば shell/batch の常駐版。接続の鮮度（ブックが閉じられた等）は
get_workbook 側の生存確認＋自動再接続に任せる。

制約:
- stdout は JSON-RPC の通信線なので、コマンドの print は全て捕捉してツール結果で返す
- COM は専用ワーカースレッド1本に固定（呼び出しスレッドが変わっても STA を跨がない）
- input() 待ちで固まらないよう実行中は stdin を空にする（確認系は -y を付けて呼ぶ）
"""
import atexit
import io
import os
import queue
import sys
import threading
from contextlib import redirect_stdout, redirect_stderr

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
os.chdir(SCRIPT_DIR)

import pythoncom  # noqa: E402
import vba_manager  # noqa: E402
from mcp.server.fastmcp import FastMCP  # noqa: E402

LAST_PROC = os.path.join(SCRIPT_DIR, "_last_proc.vba")
BLOCKED = {"shell", "batch"}  # 対話・標準入力前提のコマンドは MCP では使えない

_jobs = queue.Queue()


def _tokenize(line):
    import shlex
    lex = shlex.shlex(line, posix=True)
    lex.whitespace_split = True
    lex.escape = ''  # Windows パスの \ をエスケープ扱いしない（shell と同じ）
    return list(lex)


def _run_line(parser, table, line, buf):
    tokens = _tokenize(line)
    sub_args, unknown = parser.parse_known_args(tokens)
    unknown = [u for u in unknown if u not in ("--visible", "-v")]
    if unknown:
        buf.write(f"不明な引数/オプション: {' '.join(unknown)}\n")
        return False
    if not sub_args.command or sub_args.command in BLOCKED:
        buf.write("このコマンドは MCP セッション内では実行できません\n")
        return False
    return table[sub_args.command](sub_args)


def _worker():
    pythoncom.CoInitialize()
    parser = vba_manager.build_parser()
    table = vba_manager._command_table()
    while True:
        line, box, done = _jobs.get()
        buf = io.StringIO()
        old_stdin = sys.stdin
        ok = False
        try:
            sys.stdin = io.StringIO("")
            with redirect_stdout(buf), redirect_stderr(buf):
                if line == "__cleanup__":
                    vba_manager.cleanup_excel()
                    ok = True
                else:
                    ok = _run_line(parser, table, line, buf) is not False
        except SystemExit as e:
            ok = e.code in (0, None)
        except Exception as e:
            buf.write(f"\nエラー: {e}")
        finally:
            sys.stdin = old_stdin
            box["ok"] = ok
            box["out"] = buf.getvalue()
            done.set()


def _run(line, timeout=600):
    box, done = {}, threading.Event()
    _jobs.put((line, box, done))
    if not done.wait(timeout):
        return False, f"タイムアウト（{timeout}秒）: {line}"
    return box.get("ok", False), (box.get("out") or "").strip()


def _submit(line, timeout=600):
    ok, out = _run(line, timeout)
    if not ok:
        out = (out + "\n（コマンドは失敗しました）").strip()
    return out or "（出力なし）"


mcp = FastMCP("vba-manager")


@mcp.tool()
def vba(command: str) -> str:
    """vba_manager のコマンドを1行実行する。対象は今アクティブに開いている Excel ブック。

    CLI と同じ引数列をそのまま渡す。例:
      "list"（マクロ一覧） / "list-open"（開いているブック一覧） /
      "get モジュール名 プロシージャ名" / "run-macro マクロ名" /
      "read-range A1:D10" / "write-range A1 値" / "sheet-info" /
      "grep ActiveSheet" / "checkup" / "impact マクロ名"
    コマンド一覧・各引数は vba_help で確認できる。
    注意: 確認プロンプトを出すコマンドは必ず -y を付ける（例: "replace-procedure -y"）。
    shell / batch は使えない（このセッション自体が常駐＝接続使い回しのため不要）。
    """
    return _submit(command)


@mcp.tool()
def vba_help(command: str = "") -> str:
    """コマンド一覧（引数なし）または個別コマンドの詳細ヘルプ（例: command="get"）を返す。"""
    return _submit(f"{command} --help" if command else "--help")


@mcp.tool()
def get_procedure(name: str, module: str = "") -> str:
    """プロシージャのコードを取得して返す（_last_proc.vba にも保存される）。

    同名プロシージャが複数モジュールにある場合は module を指定する。
    修正の流れ: get_procedure → set_procedure_code → replace_procedure
    """
    cmd = f'get "{module}" "{name}"' if module else f'get "{name}"'
    ok, out = _run(cmd)
    if not ok:
        return (out + "\n（取得に失敗しました）").strip()
    try:
        with open(LAST_PROC, 'r', encoding='utf-8') as f:
            code = f.read()
    except OSError:
        return out
    return out + "\n----- コード -----\n" + code


@mcp.tool()
def set_procedure_code(code: str) -> str:
    """修正後のプロシージャコードを _last_proc.vba に書き込む（UTF-8）。

    Sub〜End Sub（または Function〜End Function）まで丸ごと渡す。
    この後 replace_procedure を呼ぶと差分表示つきで適用される。
    """
    with open(LAST_PROC, 'w', encoding='utf-8') as f:
        f.write(code)
    n = code.count('\n') + 1
    return f"_last_proc.vba に {n} 行を書き込みました。replace_procedure で適用できます。"


@mcp.tool()
def replace_procedure(module: str = "") -> str:
    """_last_proc.vba の内容でプロシージャを置換する（自動バックアップ・差分表示つき）。

    同名プロシージャが複数ある場合は module で対象モジュールを明示する。
    """
    cmd = "replace-procedure -y" + (f' --module "{module}"' if module else "")
    return _submit(cmd)


def _shutdown():
    # ツールが自動起動した Excel があれば畳む（ユーザーの Excel には触れない）
    try:
        box, done = {}, threading.Event()
        _jobs.put(("__cleanup__", box, done))
        done.wait(10)
    except Exception:
        pass


if __name__ == "__main__":
    threading.Thread(target=_worker, daemon=True).start()
    atexit.register(_shutdown)
    mcp.run()
