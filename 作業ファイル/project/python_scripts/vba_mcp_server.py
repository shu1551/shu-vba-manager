#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""vba_mcp_server.py — vba_manager を MCP サーバー化する薄い窓口。

実体は vba_manager.py の関数群そのもの（この層は判断をしない）。常駐プロセスが
get_workbook の接続キャッシュを持ち続けるため、CLI の「1コマンド毎の COM 再接続」
が消える。いわば shell/batch の常駐版。接続の鮮度（ブックが閉じられた等）は
get_workbook 側の生存確認＋自動再接続に任せる。
ただし握りっぱなしにはしない: アイドル IDLE_RELEASE_SECS 秒で COM 参照を解放する
（常駐が参照を握ったままだと、ユーザーが×で閉じた Excel がゾンビ化するため）。

制約:
- stdout は JSON-RPC の通信線なので、コマンドの print は全て捕捉してツール結果で返す
- COM は専用ワーカースレッド1本に固定（呼び出しスレッドが変わっても STA を跨がない）
- input() 待ちで固まらないよう実行中は stdin を空にする（確認系は -y を付けて呼ぶ）
"""
import atexit
import gc
import io
import os
import queue
import sys
import threading

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
os.chdir(SCRIPT_DIR)

import pythoncom  # noqa: E402
import vba_manager  # noqa: E402
from mcp.server.fastmcp import FastMCP  # noqa: E402

LAST_PROC = os.path.join(SCRIPT_DIR, "_last_proc.vba")
BLOCKED = {"shell", "batch"}  # 対話・標準入力前提のコマンドは MCP では使えない

# 本物の標準入出力（JSON-RPC の通信線）を起動時に控えておく。
# ワーカーは実行中 sys.stdout を自分のバッファに差し替えるが、復元先を
# 「入った時点の sys.stdout」にすると、世代交代（_restart_worker）後に
# 起動した新ワーカーが「前の世代の捨てられたバッファ」を掴んでそこへ復元し続ける。
# 以後ジョブ外の print（cleanup の DEBUG 行など）が誰も読まないバッファに溜まる。
# 復元先は常にこの実体にする。
REAL_STDIN, REAL_STDOUT, REAL_STDERR = sys.stdin, sys.stdout, sys.stderr

# アイドルでこの秒数コマンドが来なければ Excel への COM 参照を解放する。
# 常駐が Application/Workbook 参照を握ったままだと、ユーザーが Excel を×で
# 閉じてもプロセスが数十秒死なずゾンビ化し、その隙にブックを開くと死にかけ
# プロセスに合流して XLSTART(PERSONAL.XLSB) を読まない実害がある
# （2026-07-10 対照実験: COM未接触/使い捨てスクリプトは×閉じ2.6秒で消滅、
# 常駐ワーカー接続時のみ25秒超のゾンビ化）。解放後の次コマンドは get_workbook
# が再接続するだけなので、連続コマンド中の速度は落ちない。
IDLE_RELEASE_SECS = 5

_jobs = queue.Queue()
_worker_lock = threading.Lock()


class _ThreadRouted(io.TextIOBase):
    """スレッドごとに出力先を振り分ける stdout/stderr/stdin の代理。

    従来はワーカーがジョブ毎に sys.stdout 自体を差し替えていた。しかし sys.stdout は
    プロセス全体のグローバルであってスレッドローカルではないため、次の穴が残る:

      1. W1 が MsgBox に捕まり 600 秒でタイムアウト → 世代交代で見捨てられる
      2. W2 がジョブを終え、finally で sys.stdout を REAL_STDOUT に戻す
      3. そこで人がダイアログを閉じ、W1 が復帰して print する
      4. その print は REAL_STDOUT ＝ JSON-RPC の通信線へ流れ、セッションが壊れる

    従来のガード（自分が据えたものが残っているときだけ戻す）が守っていたのは
    「復元時」だけで、「実行中の書き込み」は無防備だった。

    この代理を起動時に1回だけ据え、ワーカーは sys.stdout に二度と触らず
    「自分のスレッドのバッファ」を登録するだけにする。見捨てられた旧ワーカーは
    自分の（誰も読まない）バッファに書き続け、通信線には絶対に届かない。
    """
    _local = threading.local()

    def __init__(self, real, stream='out'):
        self._real = real
        self._stream = stream          # 'out' | 'err' | 'in'

    @classmethod
    def bind(cls, out_buf, err_buf, in_buf, json_mode=False):
        """このスレッドの出力先を登録する（ワーカーが自分のジョブ開始時に呼ぶ）"""
        cls._local.out = out_buf
        cls._local.err = err_buf
        cls._local.inp = in_buf
        cls._local.json_mode = json_mode

    @classmethod
    def unbind(cls):
        cls._local.out = None
        cls._local.err = None
        cls._local.inp = None
        cls._local.json_mode = False

    @classmethod
    def json_mode(cls):
        return getattr(cls._local, 'json_mode', False)

    def _target(self):
        if self._stream == 'out':
            return getattr(_ThreadRouted._local, 'out', None) or self._real
        if self._stream == 'err':
            return getattr(_ThreadRouted._local, 'err', None) or self._real
        return getattr(_ThreadRouted._local, 'inp', None) or self._real

    def write(self, s):
        return self._target().write(s)

    def flush(self):
        try:
            self._target().flush()
        except Exception:
            pass

    def readline(self, *a):
        return self._target().readline(*a)

    def read(self, *a):
        return self._target().read(*a)

    def isatty(self):
        return False


# 起動時に1回だけ据える。以後 sys.stdout/stderr/stdin は誰も差し替えない
sys.stdout = _ThreadRouted(REAL_STDOUT, 'out')
sys.stderr = _ThreadRouted(REAL_STDERR, 'err')
sys.stdin = _ThreadRouted(REAL_STDIN, 'in')


def _install_json_aware_print():
    """--json のジョブでは、素の print()（情報行）を stderr 側へ退避する。

    JSON 本体は print(..., file=sys.stdout) と明示して出力される一方、
    情報行（「対象ブック: ...」等）は素の print() で出る。CLI では
    setup_encoding が builtins.print にパッチを当ててこれを分離しているが、
    MCP は setup_encoding を呼ばないため、両者が同じ stdout に混ざり
    機械処理側が json.loads できなかった。同じ分離をここで行う。
    スレッドローカル判定にしてあるので、他スレッドの print は巻き込まない。
    """
    import builtins
    _orig_print = builtins.print

    def _print(*args, **kwargs):
        if 'file' not in kwargs and _ThreadRouted.json_mode():
            kwargs['file'] = sys.stderr
        _orig_print(*args, **kwargs)

    builtins.print = _print


_install_json_aware_print()


def _release_com_refs():
    """接続キャッシュの Excel COM 参照を手放す（COM を作ったワーカースレッド上で呼ぶ）。

    ①get_workbook の接続キャッシュを手放す（ユーザーの Excel に参照を残さない）。
    ②ツールが自動起動した非表示 Excel（_created_instances）も畳む。
      常駐サーバーでは終了時の cleanup_excel が何時間も来ないため、放置すると
      アドインを読まない非表示 Excel が残留し、ユーザーが開いたファイルが合流して
      「アドインが効かない」事故になる（2026-07-12 特定・当日3体残留の実害）。
      未保存の変更を持つインスタンスだけは温存する（無言の変更破棄をしない）。
    """
    try:
        if vba_manager._wb_cache:
            vba_manager._wb_cache.clear()
            gc.collect()  # 参照サイクルに残った COM ラッパも確実に Release させる
    except Exception:
        pass
    try:
        vba_manager.release_created_instances()
    except Exception:
        pass


def _tokenize(line):
    import shlex
    lex = shlex.shlex(line, posix=True)
    lex.whitespace_split = True
    lex.escape = ''  # Windows パスの \ をエスケープ扱いしない（shell と同じ）
    lex.commenters = ''  # 行中の # をコメント扱いしない（#FF0000 等が消える。shell/batch と同じ）
    return list(lex)


def _wants_json(line):
    """その行が --json を求めているか（CLI は sys.argv を見るが MCP には無いので行を見る）"""
    try:
        return "--json" in _tokenize(line)
    except Exception:
        return "--json" in line


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


def _worker(jobs):
    pythoncom.CoInitialize()
    parser = vba_manager.build_parser()
    table = vba_manager._command_table()
    while True:
        try:
            item = jobs.get(timeout=IDLE_RELEASE_SECS)
        except queue.Empty:
            # アイドル: 連続コマンド中は保持していた COM 参照をここで手放す。
            # 世代交代後に復帰した旧ワーカーは新世代のキャッシュに触らない
            if jobs is _jobs:
                _release_com_refs()
            continue
        if item is None:
            # 世代交代の停止合図（タイムアウト後に復帰した旧ワーカーはここで退場）
            break
        line, box, done = item
        with _worker_lock:
            if box.get("cancelled"):
                # 実行前にタイムアウト済みのジョブ。「タイムアウトしました」と報告済みなのに
                # 後から副作用だけ走る、を防ぐためここで捨てる
                continue
            box["started"] = True
        # stdout（JSON本体）と stderr（情報行・警告）を分ける。混ぜると --json の
        # 出力に情報行が紛れ込み、機械処理側が json.loads できない
        out_buf = io.StringIO()
        err_buf = io.StringIO()
        ok = False
        try:
            # sys.stdout をグローバルに差し替えるのではなく、自分のスレッドの
            # 出力先を登録するだけ（_ThreadRouted 参照）。こうすると世代交代で
            # 見捨てられた旧ワーカーが後から復帰して print しても、その出力は
            # 自分のバッファに行き、JSON-RPC の通信線には絶対に漏れない。
            _ThreadRouted.bind(out_buf, err_buf, io.StringIO(""),
                               json_mode=_wants_json(line))
            if line == "__cleanup__":
                vba_manager.cleanup_excel()
                ok = True
            else:
                ok = _run_line(parser, table, line, err_buf) is not False
        except SystemExit as e:
            ok = e.code in (0, None)
        except EOFError:
            err_buf.write("\nエラー: 確認プロンプト待ちになりました（MCP では応答できません）。"
                          "確認付きコマンドは -y を付けて再実行してください")
        except Exception as e:
            err_buf.write(f"\nエラー: {e}")
        finally:
            _ThreadRouted.unbind()
            box["ok"] = ok
            box["out"] = out_buf.getvalue()
            box["err"] = err_buf.getvalue()
            done.set()


def _restart_worker():
    """タイムアウトで詰まったワーカーを見捨てて新しい世代に交代する。

    ワーカーは1本のキュー直列なので、詰まったジョブを放置すると以後の
    全ツール呼び出しがタイムアウトになる（サーバー再起動まで回復不能）。
    旧キューには停止合図を置き、旧ワーカーが後で復帰しても新ジョブを食わせない。
    新ワーカーは別スレッド＝別 STA なので、旧スレッドの COM 参照は使えない。
    接続キャッシュを捨てて新規接続からやり直す。
    """
    global _jobs
    with _worker_lock:
        old = _jobs
        _jobs = queue.Queue()
        old.put(None)
        try:
            vba_manager._wb_cache.clear()
        except Exception:
            pass
        threading.Thread(target=_worker, args=(_jobs,), daemon=True).start()


def _run(line, timeout=600):
    box, done = {}, threading.Event()
    with _worker_lock:
        # _restart_worker と排他にする。参照と put の間に世代交代が挟まると、
        # ジョブが停止済みの旧キューに落ちて誰にも実行されない
        _jobs.put((line, box, done))
    if not done.wait(timeout):
        with _worker_lock:
            started = box.get("started", False)
            if not started:
                box["cancelled"] = True   # ワーカーは実行前にこれを見て捨てる
        if not started:
            # 前のジョブが長引いて未着手のままタイムアウト。ワーカー自体は健全なので
            # 世代交代せず、このジョブだけ取り下げる
            return False, (f"タイムアウト（{timeout}秒）: {line}\n"
                           "前のコマンドが長引いているため、このコマンドは未実行のまま"
                           "取り下げました。しばらくしてから再実行してください")
        _restart_worker()
        return False, (f"タイムアウト（{timeout}秒）: {line}\n"
                       "ワーカーを再起動しました。次の呼び出しから復旧します"
                       "（実行中だったコマンドは Excel 側で続いている可能性があります。"
                       "モーダルダイアログが開いていないか確認してください）")
    return (box.get("ok", False),
            (box.get("out") or "").strip(),
            (box.get("err") or "").strip())


def _submit(line, timeout=600):
    ok, out, err = _run(line, timeout)
    # --json 指定時は JSON 本体（stdout）だけを返す。情報行（stderr）を混ぜると
    # 機械処理側が json.loads できない（CLI では setup_encoding が情報行を stderr へ
    # 退避して分離しているが、MCP は setup_encoding を呼ばないので自前で分ける）
    if _wants_json(line):
        if ok and out:
            return out
        # 失敗時だけは理由が要る（黙って空を返さない）
        return ((out + "\n" + err).strip() + "\n（コマンドは失敗しました）").strip()
    body = "\n".join(p for p in (err, out) if p).strip()
    if not ok:
        body = (body + "\n（コマンドは失敗しました）").strip()
    return body or "（出力なし）"


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
    ok, out, err = _run(cmd)
    # get のコード本文は stdout、情報行・警告は stderr。人が読むので両方返す
    body = "\n".join(p for p in (err, out) if p).strip()
    if not ok:
        return (body + "\n（取得に失敗しました）").strip()
    # get の出力に保存先とコード全文が含まれる（_last_proc.vba と同一内容）ので
    # ファイルを読み直して二重に返さない
    return body


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
        with _worker_lock:
            _jobs.put(("__cleanup__", box, done))
        done.wait(10)
    except Exception:
        pass


if __name__ == "__main__":
    threading.Thread(target=_worker, args=(_jobs,), daemon=True).start()
    atexit.register(_shutdown)
    mcp.run()
