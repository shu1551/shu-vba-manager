import sys
import io
import os
import glob
import customtkinter as ctk
import subprocess
import threading

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
LAST_PROC_FILE = os.path.join(SCRIPTS_DIR, '_last_proc.vba')


class VBAManagerGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.cached_macros = {}
        self.cached_modules = {}

        self.title("VBA Manager")
        self.geometry("960x680")
        self.minsize(800, 500)

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # ---- サイドバー ----
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)

        ctk.CTkLabel(self.sidebar_frame, text="VBA Manager",
                     font=ctk.CTkFont(size=20, weight="bold")
                     ).grid(row=0, column=0, padx=20, pady=(20, 10))

        ctk.CTkLabel(self.sidebar_frame, text="対象ファイル:",
                     font=ctk.CTkFont(size=12)
                     ).grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")

        self.file_combobox = ctk.CTkComboBox(
            self.sidebar_frame, values=["読み込み中..."],
            command=self.on_file_selected)
        self.file_combobox.grid(row=2, column=0, padx=20, pady=(5, 10))

        self.status_label = ctk.CTkLabel(
            self.sidebar_frame, text="選択中: -", text_color="gray",
            font=ctk.CTkFont(size=11), wraplength=170)
        self.status_label.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="w")

        # last_proc.vba の状態表示
        self.proc_label = ctk.CTkLabel(
            self.sidebar_frame, text="コード: (未取得)",
            text_color="gray", font=ctk.CTkFont(size=11), wraplength=170)
        self.proc_label.grid(row=4, column=0, padx=20, pady=(0, 20), sticky="w")

        self.after(100, self.load_excel_files)
        self.after(500, self.refresh_proc_label)

        # ---- メインエリア ----
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # ボタンエリア
        self.buttons_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.buttons_frame.grid(row=0, column=0, sticky="ew")
        for i in range(3):
            self.buttons_frame.grid_columnconfigure(i, weight=1)

        buttons_info = [
            ("📄 マクロ一覧",          self.btn_show_macros),
            ("👁  コードを取得",        self.btn_show_code),
            ("✍  コードを適用",        self.btn_apply_code),
            ("🔧 マクロを修正",         self.btn_fix_macro),
            ("🧹 クリーンアップ(整形)", self.btn_cleanup),
            ("📦 モジュール一覧",       self.btn_list_modules),
            ("🔬 詳細解析",            self.btn_analyze),
            ("☑  構文チェック",        self.btn_syntax_check),
            ("🩺 診断",               self.btn_diag),
        ]

        row, col = 0, 0
        for text, command in buttons_info:
            btn = ctk.CTkButton(
                self.buttons_frame, text=text, height=42,
                font=ctk.CTkFont(size=13), command=command)
            btn.grid(row=row, column=col, padx=8, pady=6, sticky="ew")
            col += 1
            if col > 2:
                col = 0
                row += 1

        # ログエリア
        self.log_container = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.log_container.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        self.log_container.grid_rowconfigure(1, weight=1)
        self.log_container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.log_container, text="実行ログ:",
                     font=ctk.CTkFont(size=12, weight="bold")
                     ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.log_textbox = ctk.CTkTextbox(
            self.log_container, font=ctk.CTkFont("Consolas", size=13))
        self.log_textbox.grid(row=1, column=0, sticky="nsew")
        self.log_textbox.insert("0.0", "起動完了。待機中...\n")

    # ================================================================
    # ユーティリティ
    # ================================================================

    def append_log(self, text):
        self.log_textbox.insert("end", text + "\n")
        self.log_textbox.see("end")

    def get_vba_manager_path(self):
        return os.path.join(SCRIPTS_DIR, "vba_manager.py")

    def get_target_file_path(self):
        filename = self.file_combobox.get()
        return os.path.abspath(
            os.path.join(SCRIPTS_DIR, "..", "..", "..", filename))

    def on_file_selected(self, choice):
        self.status_label.configure(text=f"選択中: {choice}")
        if choice and choice not in ("ファイルが見つかりません", "読み込み中..."):
            threading.Thread(
                target=self._prefetch, args=(choice,), daemon=True).start()

    def _prefetch(self, filename):
        target = os.path.abspath(
            os.path.join(SCRIPTS_DIR, "..", "..", "..", filename))
        for cmd, cache in [("list", self.cached_macros),
                           ("list-modules", self.cached_modules)]:
            try:
                r = subprocess.run(
                    [sys.executable, self.get_vba_manager_path(), cmd, target],
                    capture_output=True, text=True, encoding='utf-8',
                    creationflags=0x08000000)
                prefix = "MACRO:" if cmd == "list" else "MODULE:"
                items = []
                for line in r.stdout.splitlines():
                    if line.startswith(prefix):
                        # "MODULE:名前  (型)" → 名前だけ取り出す
                        val = line[len(prefix):].split()[0].strip()
                        items.append(val)
                if items:
                    cache[target] = items
            except Exception:
                pass

    def load_excel_files(self):
        target_dir = os.path.abspath(os.path.join(SCRIPTS_DIR, "..", "..", ".."))
        excel_files = []
        for ext in ("*.xlsm", "*.xlsb", "*.xlsx"):
            for f in glob.glob(os.path.join(target_dir, ext)):
                excel_files.append(os.path.basename(f))
        if not excel_files:
            excel_files = ["ファイルが見つかりません"]
        self.file_combobox.configure(values=excel_files)
        self.file_combobox.set(excel_files[0])
        self.on_file_selected(excel_files[0])

    def refresh_proc_label(self):
        if os.path.exists(LAST_PROC_FILE):
            mtime = os.path.getmtime(LAST_PROC_FILE)
            import datetime
            dt = datetime.datetime.fromtimestamp(mtime).strftime("%m/%d %H:%M")
            # ファイルの先頭行（Sub名）を表示
            try:
                with open(LAST_PROC_FILE, encoding='utf-8') as f:
                    first = f.readline().strip()
                self.proc_label.configure(
                    text=f"コード取得済み:\n{first}\n({dt})",
                    text_color="#4CAF50")
            except Exception:
                self.proc_label.configure(
                    text=f"コード取得済み ({dt})", text_color="#4CAF50")
        else:
            self.proc_label.configure(text="コード: (未取得)", text_color="gray")
        self.after(3000, self.refresh_proc_label)

    def run_backend_command(self, cmd_args, callback=None):
        """vba_manager.py を実行してログに出力"""
        full_cmd = [sys.executable, self.get_vba_manager_path()] + cmd_args
        self.append_log(f"\n> vba_manager {' '.join(cmd_args)}")

        def _run():
            try:
                proc = subprocess.Popen(
                    full_cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding='utf-8')
                for line in iter(proc.stdout.readline, ''):
                    self.after(0, self.append_log, line.rstrip())
                proc.stdout.close()
                rc = proc.wait()
                self.after(0, self.append_log, f"完了 (コード: {rc})")
                if callback and rc == 0:
                    self.after(0, callback)
            except Exception as e:
                self.after(0, self.append_log, f"[エラー] {e}")

        threading.Thread(target=_run, daemon=True).start()

    def run_script(self, script_path, args, callback=None):
        """任意のスクリプトを実行してログに出力"""
        full_cmd = [sys.executable, script_path] + args
        self.append_log(f"\n> {os.path.basename(script_path)} {' '.join(args)}")

        def _run():
            try:
                proc = subprocess.Popen(
                    full_cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding='utf-8', cwd=SCRIPTS_DIR,
                    input=None)
                for line in iter(proc.stdout.readline, ''):
                    self.after(0, self.append_log, line.rstrip())
                proc.stdout.close()
                rc = proc.wait()
                self.after(0, self.append_log, f"完了 (コード: {rc})")
                if callback and rc == 0:
                    self.after(0, callback)
            except Exception as e:
                self.after(0, self.append_log, f"[エラー] {e}")

        threading.Thread(target=_run, daemon=True).start()

    def open_in_editor(self, filepath):
        try:
            subprocess.Popen(['notepad', filepath])
            self.append_log(f"メモ帳で開きました: {os.path.basename(filepath)}")
        except Exception as e:
            self.append_log(f"[エラー] 開けませんでした: {e}")

    # ================================================================
    # ボタンハンドラ
    # ================================================================

    def btn_show_macros(self):
        self.open_macro_dialog("確認するマクロを選択してください")

    def btn_show_code(self):
        """マクロを選択してコードを _last_proc.vba に取得"""
        self.open_macro_dialog("コードを取得するマクロを選択してください",
                               action="get")

    def btn_apply_code(self):
        """_last_proc.vba の内容をExcelに適用"""
        if not os.path.exists(LAST_PROC_FILE):
            self.append_log("_last_proc.vba がありません。先にコードを取得してください。")
            return
        target_file = self.get_target_file_path()
        self.run_backend_command(["replace-procedure", target_file],
                                 callback=self.refresh_proc_label)

    def btn_fix_macro(self):
        """マクロを選択してコードを取得し、GUIのコードビューアで表示"""
        self.open_macro_dialog("コードを確認・修正するマクロを選択してください",
                               action="fix")

    def btn_cleanup(self):
        """モジュールを選択してクリーンアップ用スキャフォルドを生成・開く"""
        self.open_module_dialog()

    def btn_list_modules(self):
        target_file = self.get_target_file_path()
        self.run_backend_command(["list-modules", target_file])

    def btn_analyze(self):
        """詳細解析: .bas ファイルの行数・Sub数を一覧表示"""
        self.append_log("\n=== 詳細解析 ===")

        def _run():
            sys.path.insert(0, SCRIPTS_DIR)
            from bas_editor import read_bas
            bas_files = sorted(glob.glob(os.path.join(SCRIPTS_DIR, "*.bas")))
            if not bas_files:
                self.after(0, self.append_log, "  .bas ファイルが見つかりません")
                return
            for bas_path in bas_files:
                try:
                    content = read_bas(bas_path)
                    lines = content.split('\n')
                    subs = [l.strip() for l in lines
                            if any(l.strip().startswith(p)
                                   for p in ('Sub ', 'Function ',
                                             'Public Sub ', 'Private Sub ',
                                             'Public Function ', 'Private Function '))
                            and '(' in l]
                    name = os.path.basename(bas_path)
                    self.after(0, self.append_log,
                               f"  {name}: {len(lines)}行  /  {len(subs)} Sub/Function")
                except Exception as e:
                    self.after(0, self.append_log,
                               f"  {os.path.basename(bas_path)}: ERROR {e}")
            self.after(0, self.append_log, "解析完了")

        threading.Thread(target=_run, daemon=True).start()

    def btn_syntax_check(self):
        """構文チェック: .bas ファイルの重複Sub・End Sub対応を確認"""
        self.append_log("\n=== 構文チェック ===")

        def _run():
            sys.path.insert(0, SCRIPTS_DIR)
            from bas_editor import read_bas
            bas_files = sorted(glob.glob(os.path.join(SCRIPTS_DIR, "*.bas")))
            if not bas_files:
                self.after(0, self.append_log, "  .bas ファイルが見つかりません")
                return
            all_ok = True
            for bas_path in bas_files:
                name = os.path.basename(bas_path)
                try:
                    content = read_bas(bas_path)
                    lines = content.split('\n')
                    errors = []
                    sub_names = []
                    depth = 0
                    for i, line in enumerate(lines, 1):
                        s = line.strip()
                        # Sub/Function 開始
                        for p in ('Sub ', 'Function ', 'Public Sub ',
                                  'Private Sub ', 'Public Function ',
                                  'Private Function '):
                            if s.startswith(p) and '(' in s:
                                n = s[len(p):].split('(')[0].strip()
                                if n in sub_names:
                                    errors.append(f"L{i} 重複: {n}")
                                sub_names.append(n)
                                depth += 1
                                break
                        # End Sub / End Function
                        if s in ('End Sub', 'End Function'):
                            depth -= 1
                            if depth < 0:
                                errors.append(f"L{i} 対応するSubなしのEnd Sub")
                                depth = 0
                    if depth != 0:
                        errors.append(f"End Sub 不足 ({depth}個)")

                    if errors:
                        all_ok = False
                        self.after(0, self.append_log, f"  ✗ {name}:")
                        for e in errors:
                            self.after(0, self.append_log, f"      {e}")
                    else:
                        self.after(0, self.append_log,
                                   f"  ✓ {name}: 問題なし ({len(sub_names)} Sub)")
                except Exception as e:
                    self.after(0, self.append_log, f"  ! {name}: {e}")
            self.after(0, self.append_log,
                       "チェック完了" + (" — 問題なし" if all_ok else " — エラーあり"))

        threading.Thread(target=_run, daemon=True).start()

    def btn_diag(self):
        self.run_backend_command(["diag"])

    # ================================================================
    # ダイアログ
    # ================================================================

    def _make_select_dialog(self, title, items, on_select, loading=False):
        """汎用選択ダイアログを生成して返す"""
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("420x520")
        dialog.attributes("-topmost", True)
        dialog.transient(self)

        label = ctk.CTkLabel(
            dialog,
            text=title + ("\n(読込中...)" if loading else ""),
            font=ctk.CTkFont(size=14))
        label.pack(pady=15)

        frame = ctk.CTkScrollableFrame(dialog)
        frame.pack(fill=ctk.BOTH, expand=True, padx=20, pady=(0, 20))

        def populate(items_list):
            label.configure(text=title)
            if not items_list:
                ctk.CTkLabel(frame, text="項目が見つかりません").pack(pady=20)
                return
            for item in items_list:
                ctk.CTkButton(
                    frame, text=item,
                    fg_color="transparent", border_width=1,
                    text_color=("gray10", "#DCE4EE"),
                    hover_color=("gray70", "gray30"),
                    command=lambda v=item: (dialog.destroy(), on_select(v))
                ).pack(fill="x", padx=10, pady=4)

        if items is not None:
            populate(items)
        return dialog, populate

    def open_macro_dialog(self, title, action=None):
        target_file = self.get_target_file_path()
        cached = self.cached_macros.get(target_file)

        def on_select(macro_name):
            if action in ("get", "fix"):
                def after_get():
                    self.refresh_proc_label()
                    if action == "fix":
                        self.open_code_viewer(macro_name)
                self.run_backend_command(
                    ["get", target_file, macro_name], callback=after_get)
            else:
                self.append_log(f"選択: {macro_name}")

        dialog, populate = self._make_select_dialog(
            title, cached, on_select, loading=(cached is None))

        if cached is None:
            def _fetch():
                try:
                    r = subprocess.run(
                        [sys.executable, self.get_vba_manager_path(),
                         "list", target_file],
                        capture_output=True, text=True, encoding='utf-8',
                        creationflags=0x08000000)
                    items = [l[len("MACRO:"):].strip()
                             for l in r.stdout.splitlines()
                             if l.startswith("MACRO:")]
                    if items:
                        self.cached_macros[target_file] = items
                    self.after(0, populate, items)
                except Exception as e:
                    self.after(0, self.append_log, f"[エラー] {e}")
            threading.Thread(target=_fetch, daemon=True).start()

    def open_module_dialog(self):
        target_file = self.get_target_file_path()
        cached = self.cached_modules.get(target_file)

        def on_select(module_name):
            self._do_cleanup(module_name, target_file)

        dialog, populate = self._make_select_dialog(
            "クリーンアップするモジュールを選択",
            cached, on_select, loading=(cached is None))

        if cached is None:
            def _fetch():
                try:
                    r = subprocess.run(
                        [sys.executable, self.get_vba_manager_path(),
                         "list-modules", target_file],
                        capture_output=True, text=True, encoding='utf-8',
                        creationflags=0x08000000)
                    items = [l[len("MODULE:"):].split()[0].strip()
                             for l in r.stdout.splitlines()
                             if l.startswith("MODULE:")]
                    if items:
                        self.cached_modules[target_file] = items
                    self.after(0, populate, items)
                except Exception as e:
                    self.after(0, self.append_log, f"[エラー] {e}")
            threading.Thread(target=_fetch, daemon=True).start()

    def _do_cleanup(self, module_name, target_file):
        """クリーンアップ:
           optimize_xxx.py があれば自動実行 → Excelへ反映
           なければ「Claudeに依頼を」と案内
        """
        self.append_log(f"\n=== クリーンアップ: {module_name} ===")
        optimize_path = os.path.join(SCRIPTS_DIR, f"optimize_{module_name}.py")

        def after_format():
            """整形後にExcelへ反映"""
            self.append_log("Excelへ反映中...")
            self.run_backend_command(
                ["replace-module", target_file, module_name,
                 f"{module_name}.bas"],
                callback=lambda: self.append_log(
                    f"✓ {module_name} のクリーンアップ完了"))

        if os.path.exists(optimize_path):
            # optimize_xxx.py がある → 詳細整形スクリプトを実行
            self.append_log(f"  optimize_{module_name}.py を実行します")
            self.run_script(optimize_path, [], callback=after_format)
        else:
            # ない → format_bas.py で基本自動整形
            self.append_log(
                f"  optimize_{module_name}.py が未作成のため、\n"
                f"  format_bas.py で基本整形（Dim整理・空行統一）を実行します")
            format_path = os.path.join(SCRIPTS_DIR, "format_bas.py")
            self.run_script(format_path, [module_name], callback=after_format)

    def open_code_viewer(self, macro_name):
        """_last_proc.vba の内容をGUI内のポップアップで表示"""
        if not os.path.exists(LAST_PROC_FILE):
            self.append_log("コードファイルが見つかりません")
            return
        try:
            with open(LAST_PROC_FILE, encoding='utf-8') as f:
                code = f.read()
        except Exception as e:
            self.append_log(f"[エラー] {e}")
            return

        win = ctk.CTkToplevel(self)
        win.title(f"コード: {macro_name}")
        win.geometry("700x600")
        win.attributes("-topmost", True)
        win.transient(self)

        win.grid_rowconfigure(1, weight=1)
        win.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            win,
            text=f"{macro_name}  —  _last_proc.vba  |  修正後は「✍ コードを適用」で反映",
            font=ctk.CTkFont(size=12), text_color="gray"
        ).grid(row=0, column=0, padx=16, pady=(12, 4), sticky="w")

        textbox = ctk.CTkTextbox(win, font=ctk.CTkFont("Consolas", size=13))
        textbox.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="nsew")
        textbox.insert("0.0", code)
        textbox.configure(state="disabled")


def main():
    app = VBAManagerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
