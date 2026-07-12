# VBAマネージャー — Python と Claude Code で Excel VBA を操作するツールキット

Excel の VBA モジュールや UserForm を、**Python 経由でコマンドラインから操作する**ためのツールキットです。
Claude Code（または他の AI コーディングアシスタント）と組み合わせれば、
**「マクロを追加して」「フォームを直して」と話しかけるだけで** Excel の VBA が書き換わるという、
まるで魔法のような開発環境を構築できます。

> 📝 **開発の背景や思想はこちらの Qiita 記事にまとめています**：
> [Excel VBA × Claude Code で「会話するだけでマクロが直る」開発環境を作った話](https://qiita.com/shu15511551/items/d3bbba2dac4007327db6)

## ダウンロード

### ワンクリックで ZIP ダウンロード（一般ユーザー向け）

**[📦 最新版を ZIP でダウンロード](https://github.com/shu1551/shu-vba-manager/archive/refs/heads/main.zip)**

ダウンロードした ZIP を解凍すれば、すぐに使えます。

### git で取得（開発者向け）

```powershell
git clone https://github.com/shu1551/shu-vba-manager.git
```

---

## このツールキットでできること

- ✅ **VBA モジュール (.bas) の取り出し／差し替え／プロシージャ単位の置換・追加・削除**
- ✅ **全マクロ横断のコード検索（grep）と一括置換（code-replace・diffプレビュー付き）**
- ✅ **バックアップからの復元（restore）**、**コマンド列の一括実行（batch・接続1回）**
- ✅ **開いたままのブックを直接読み書き**（セル読取・書込・書式・行列操作・検索置換・PDF出力など）
- ✅ **グラフ／ピボット／スライサー／PowerQuery／データモデル(DAX) まで CLI で操作**
- ✅ **ブックの検分**——健康診断（checkup）・意味構造JSON化（snapshot）・前後差分（snapshot-diff）・ボタン⇔マクロ配線図と壊れた配線の検出（wiring）
- ✅ **UserForm を「宣言」で構築**（行構造を書くだけで整列済みフォームが完成・イベント雛形も自動生成）
- ✅ **フォームの検査（lint）・実表示スクリーンショット・既存フォームの宣言コード逆変換**
- ✅ **導入セルフ診断（setup-check）**——環境が揃っているか1コマンドで○×確認
- ✅ **VBE と `.bas` ファイルのライブ同期**
- ✅ **登録済みマクロの GUI ランチャー**（customtkinter 製）
- ✅ **任意の `.xlsm` を対象にできる**（自分のブックでも、配布ブックでも）
- ✅ **Claude Code 連携で「話しかけるだけ」開発が可能**

---

## ツール一覧（`作業ファイル/project/python_scripts/`）

| ツール | 役割 |
|---|---|
| **`vba_manager.py`** | 中核ツール。VBA の `list / get / replace-module / replace-procedure / add-procedure / delete-procedure` に、コード検索 `grep`・一括置換 `code-replace`・復元 `restore`・一括実行 `batch`・導入診断 `setup-check`、シート読取（`read-range` / `sheet-info` / `screenshot`）・編集（`write-range` / `format-range` / `find-replace` / `export-pdf` 等）・グラフ／ピボット／PowerQuery／データモデル・検分（`checkup` / `snapshot` / `snapshot-diff` / `wiring`）まで **72コマンド**。実装は入口＋5パート構成（`vbam_core / vbam_vba / vbam_view / vbam_edit / vbam_heavy`・2026-07-12 分割）で、使い方は従来どおり `py vba_manager.py <コマンド>` |
| **`vbam_*.py`（5本）** | vba_manager の分割パート。**単体では使わない**が、`vba_manager.py` と同じフォルダに必要（clone / zip 取得なら自動で揃う） |
| **`vba_mcp_server.py`** | vba_manager を **MCP サーバー化**する薄い窓口（2026-07 追加）。常駐 COM 接続でコマンドごとの再接続が消え、応答は実測 0.01〜0.2 秒級。`vba`（1行で全コマンド）＋ `get_procedure / set_procedure_code / replace_procedure / vba_help` の5ツール |
| **`form_layout.py`** | UserForm を**宣言で構築**するレイアウトエンジン。行構造を書くだけでラベル整列・ボタンバー・タブ順・イベント雛形まで自動。Excel 不要の配置プレビューも（`py form_layout.py preview 宣言.py`） |
| **`form_inspect.py`** | フォームの点検。コントロール配置＋コードを1接続で取得、実表示PNG撮影（`--png` / `--png-all`）、機械検査（`--lint`）、既存フォームの宣言コード逆変換（`--to-layout`） |
| **`form_tool.py`** | フォームの幾何操作 CLI。`scale / set / move / align / size-match / distribute / tab-order / rename-control / delete-control / copy-form` |
| **`form_builder.py`** | UserForm 構築の低レベル部品。`add_btn`, `add_lbl` などのヘルパーと `Grid` 配置（カレンダー等の自由配置向け） |
| **`test_tools.py`** | 自動テスト（pytest・67件）。レイアウト計算・エンコーディングガード・lint・名前衝突ガード等の COM 不要部分を検証 |
| **`live_sync_vba.py`** | VBE と `.bas` ファイルをリアルタイム同期 |
| **`menu_launcher.py`** / **`select_macro_gui.py`** | customtkinter 製のマクロ検索／実行 GUI |
| **`bas_editor.py`** | `.bas` ファイル編集ヘルパー（プロシージャ単位の `replace_sub`, `read_bas`） |
| **`format_bas.py`** | `.bas` 整形（空行・インデント正規化） |
| **`optimize_vba_modules.py`** | VBA モジュール最適化 |
| **`publish_check.py`** | ブック公開前の個人情報チェック（作成者実名・隠しシート・定義名のローカルパス・外部リンク・コメント・残留入力値）＋ `--scrub` でメタデータ空欄化。ブックを配布・公開する前に1コマンドで |

---

## 動作環境

| 必須 | 備考 |
|---|---|
| Windows 10 / 11 | 日本語環境 |
| Microsoft Excel | デスクトップ版（Microsoft 365 / Office 2019 以降）|
| Python 3.9 以上 | `py` コマンドが使えること |
| Claude Code（推奨） | <https://claude.com/claude-code> 「話しかけるだけ」開発を実現するために |

---

## セットアップ

### ① Python のインストール

1. <https://www.python.org/downloads/> から最新版を DL
2. インストーラーで **「Add Python to PATH」を必ずチェック**
3. PowerShell で `py --version` 確認

### ② Python パッケージのインストール

```powershell
cd "$env:USERPROFILE\Desktop\VBAマネージャー"
py -m pip install -r requirements.txt
```

依存パッケージ:
- `pywin32` — Excel COM 操作
- `customtkinter` — GUI ランチャー

### ③ Excel のセキュリティ設定（★最重要★）

**この設定をしないと Python から VBA を一切操作できません。**

1. Excel を起動
2. **ファイル → オプション → トラスト センター → トラスト センターの設定**
3. **マクロの設定** タブを開く
4. **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック**
5. OK で閉じる

ここで詰まる人が多いので必ず確認してください。

### ④ 導入セルフ診断（ここまでの確認）

設定が揃ったかどうかは、1コマンドで確認できます：

```powershell
cd "作業ファイル\project\python_scripts"
py vba_manager.py setup-check
```

Python / pywin32 / Excel / 上記③の信頼設定などを ○× で表示し、NG には対処法が出ます。
分からないところは、この出力をそのまま AI に貼って聞いてください。

### ⑤ Claude Code のインストール（推奨）

<https://claude.com/claude-code> の手順に従ってインストール。
インストール後、配布フォルダで `claude` コマンドを起動するとここを作業ディレクトリとして AI と対話できます。

---

## 主要ツールの使い方

### `vba_manager.py` — VBA モジュール CRUD ＋ Excel 操作

```powershell
cd "作業ファイル\project\python_scripts"

# 全モジュール一覧
py vba_manager.py list "C:\path\to\任意.xlsm"

# .bas を取り出し
py vba_manager.py get "C:\path\to\任意.xlsm" モジュール名

# .bas モジュールを置き換え（破壊的、自動バックアップあり）
py vba_manager.py replace-module "C:\path\to\任意.xlsm" モジュール名 ファイル.bas

# プロシージャ単位で置き換え（既存プロシージャをピンポイント差し替え）
# 非対話実行（Claude Code 連携等）では -y で確認プロンプトをスキップ
py vba_manager.py replace-procedure -y "C:\path\to\任意.xlsm" patched.vba
```

ファイルパスを省略すると**アクティブな（今 Excel で開いている）ブック**を自動対象にする。
VBA だけでなく、開いたままのブックのシートを直接読み書きできる：

```powershell
# 読む（目）：セル値・シート構成・範囲のPNG化
py vba_manager.py read-range A1:D10
py vba_manager.py sheet-info
py vba_manager.py screenshot A1:H30

# 書く（手）：値・数式・書式・行列・検索置換・印刷設定など
py vba_manager.py write-range C1 "=SUM(A1:A10)"
py vba_manager.py format-range A1:D1 --bold --bg "#FFFF00"
py vba_manager.py find-replace 旧 新 A1:Z99

# 重量級：グラフ・ピボット・スライサー・PowerQuery・データモデル(DAX)
py vba_manager.py chart create A1:B5 --type column --title "月別売上"
py vba_manager.py pivot create 元データ!A1:C100 --rows 部門 --values 売上
py vba_manager.py powerquery list
py vba_manager.py datamodel list

# 横断系：検索・一括置換・復元・一括実行（2026-07 追加）
py vba_manager.py grep "ActiveSheet"              # 全マクロからコード検索（モジュール/行番号つき）
py vba_manager.py code-replace "旧" "新" -y       # 一括置換（diffプレビュー・バックアップ・Attribute温存）
py vba_manager.py list-backups                    # 自動バックアップの一覧（5世代）
py vba_manager.py restore <バックアップ.bas>       # 置換前の状態に復元（undo）
py vba_manager.py batch cmds.txt                  # コマンド列を接続1回で一括実行（実測: 6コマンド約1秒）
py vba_manager.py export-pdf 出力.pdf --sheet 集計  # PDF出力

# 検分（現地調査）：健康診断・意味構造JSON・前後差分・配線図（2026-07 追加）
py vba_manager.py checkup                         # ブックの健康診断（総合判定A/B/C・定期健診・カルテ）
py vba_manager.py snapshot                        # ブックを意味構造JSONに畳む（→ _last_snapshot.json）
py vba_manager.py snapshot-diff before.json       # snapshot同士の差分＝マクロが実際に何を変えたか（COM不要）
py vba_manager.py wiring                          # ボタン⇔マクロ配線図・壊れた配線の検出（別名: 配線図）
```

全コマンド（72個）の一覧と使い方は `vba_manager.py` 冒頭のコマンド表を参照。
※ Excel が未起動の状態でパス指定すると自動化用の Excel が新規起動される。この Excel には
アドインや PERSONAL.XLSB が読み込まれないため、普段使いには手動起動した Excel を使うこと
（ツールが警告を表示する）。

### `form_layout.py` — UserForm を「宣言」で構築（2026-07 追加・推奨）

座標の手計算は不要になりました。行構造を書くだけで、ラベル整列・ボタンバー右寄せ・
Enter/Esc 割り当て・タブ順・イベントコード雛形まで自動です：

```python
from form_layout import build_form, row, lbl, txt, combo, button_bar, ok, cancel, spacer

build_form("F_Input", "顧客登録", rows=[
    row(lbl("顧客名"), txt("txtName", required=True)),   # required でラベルに＊＋入力チェック雛形
    row(lbl("区分"), combo("cmbKind", items=["法人", "個人"])),
    spacer(),
    button_bar(ok("btnSave", "登録"), cancel("btnClose", "閉じる")),
], vba_stub=True, png=True)   # png=True で構築後に実表示スクリーンショット
```

- `py form_layout.py preview 宣言.py` — **Excel を起動せずに**配置図PNGで設計確認
- `py form_inspect.py <フォーム> --lint` — 重なり・不揃い・タブ順・孤児ハンドラの機械検査
- `py form_inspect.py <フォーム> --to-layout` — **既存フォームを宣言コードに逆変換**（改修が宣言編集になる）
- `py form_tool.py move/align/size-match/tab-order …` — 微調整もコマンドで
- タブ付きフォーム（MultiPage）・枠（Frame）・範囲選択欄・スピン付き数値なども対応

### `form_builder.py` — 低レベル部品（自由配置向け）

カレンダーの7×6ボタン格子のような自由配置は、従来どおり `add_btn` 等で直接置けます。
`Grid` / `vstack` / `hstack` の座標計算ヘルパー付き。イベントコードの注入（`inject_vba`）もこちら。

### `live_sync_vba.py` — ライブ同期

```powershell
py live_sync_vba.py
```

VBE 上での編集が `.bas` ファイルへ即時反映されます。
逆方向（`.bas` の変更 → VBE 反映）にも対応。

### `menu_launcher.py` — GUI ランチャー

```powershell
py menu_launcher.py
```

customtkinter 製 GUI が起動。登録された Public Sub を一覧から検索・実行できます。

### `vba_mcp_server.py` — MCP サーバー（2026-07 追加）

Claude Code などの MCP クライアントに登録すると、AI がこのツールキットを直接呼び出せます。
CLI が1コマンドごとに払っていた COM 再接続（全工程で最も重い処理）が消え、
常駐接続により 2 回目以降の応答は **実測 0.01〜0.2 秒級** になります。

```powershell
# Claude Code への登録（ユーザースコープ）
claude mcp add vba-manager --scope user -- py "<このフォルダの絶対パス>\作業ファイル\project\python_scripts\vba_mcp_server.py"
```

ツールは5つだけの薄い窓口です（実体は vba_manager.py そのもの）：

- `vba(command)` — CLI と同じ引数列を1行で渡す。72コマンド全部使える
- `vba_help` — コマンド一覧・個別ヘルプ
- `get_procedure` / `set_procedure_code` / `replace_procedure` — プロシージャ修正の定番3手

対象は常に「**今アクティブに開いている Excel ブック**」。特定のブックやパスには依存しません。

#### Antigravity など他の MCP クライアントで使う場合

MCP は接続後の規格なので、サーバーが立ち上がりさえすれば **Gemini など他社の AI からも同じツールがそのまま使えます**（Antigravity + Gemini で動作確認済み）。ただし「立ち上げ方」はクライアントごとの流儀なので、以下の4点に注意してください。

1. **`command` は Python の絶対パスで書く**（最重要）。
   `"command": "py"` は Claude Code では動きますが、クライアントによっては `py` を解決できず、
   **エラーも出さずにサーバーが起動しない**ことがあります（Microsoft Store 版 / Python Install Manager 環境で発生を確認）。
   自分の Python の絶対パスは次で調べられます：

   ```powershell
   py -c "import sys; print(sys.executable)"
   ```

   Antigravity の場合は `%USERPROFILE%\.gemini\antigravity-ide\mcp_config.json` に登録し、IDE を再起動します：

   ```json
   "vba-manager": {
     "command": "C:\\Users\\<ユーザー名>\\AppData\\Local\\Python\\bin\\python.exe",
     "args": ["<このフォルダの絶対パス>/作業ファイル/project/python_scripts/vba_mcp_server.py"]
   }
   ```

2. **Antigravity のエージェントターミナルから CLI を直接叩くのは不可**。
   現行の Antigravity では、エージェントのターミナルは実機内の**別デスクトップ**で実行されるため（2026-07 実測）、
   COM が開いている Excel に届かず「Excel が起動していません」になります。
   MCP サーバーは IDE 本体が通常デスクトップ側で起動するため、この問題を受けません。**必ず MCP 経由で使ってください**。

3. **接続確認は「点呼」で**。
   AI に `list-open`（vba ツール経由）を実行させ、実際に開いているブック名と一致するか突き合わせます。
   一致すれば生きた Excel に届いている証拠です。「接続できました」という AI の申告だけでは確認になりません。

4. **AI に「ファイル直接操作」の代替をさせない**。
   COM 接続に失敗した AI が、開いている .xlsm をファイルとして直接読み書きする回避策に切り替えることがあります。
   読み取りは保存時点の古い内容しか見えず、書き込みは開いているブックと衝突して**ブックを破壊します**。
   「開いているブックへのファイル直接操作は禁止。必ずこのツール経由で」と指示してください。

---

---

## Claude Code との連携：「話しかけるだけ」開発

配布フォルダで `claude` を起動した状態で、たとえばこんなプロンプトを投げます：

> **「ターゲットブックの shu001 モジュールに、アクティブセルに今日の日付を入れるマクロを追加して」**

Claude Code は自動的に：

1. `vba_manager.py get` で既存モジュールを取得
2. `.bas` を編集（CP932 エンコーディングで保存）
3. `vba_manager.py replace-module` で Excel に反映

これで **手作業ゼロでマクロが追加** されます。

### よく使うプロンプト例

- 「`〇〇.xlsm` の `Module1` に △△ するマクロを追加して」
- 「`データ入力フォーム` という UserForm を作って、ボタン3個とテキストボックス2個を配置して」
- 「`計算表.bas` の `合計計算` プロシージャを高速化して」
- 「マウスホイールでスクロールできるようにフォームに対応コードを追加して」

### ガードレールは同梱済み（CLAUDE.md と専用スキル）

このリポジトリには、Claude Code が安全に VBA を操作するための設定が
**最初から入っています**。クローン／解凍してそのまま使えます。

- **`CLAUDE.md`**（リポジトリ直下）
  Claude Code が起動時に自動で読み込むプロジェクトルール。
  「`.bas` を Edit / Write ツールで編集して文字化けさせない」
  「`get → replace-procedure` の手順を守る」「`python` ではなく `py` を使う」
  などを Claude に指示します。

- **`.claude/skills/shu-addin-manager/`**（同梱スキル）
  「マクロを追加して」「フォームを直して」などのキーワードで自動的に起動し、
  Claude に正しい作業フロー（`list → get → 修正 → replace-procedure`）を守らせます。

どちらも自分で用意する必要はありません。このフォルダで `claude` を起動すれば、
そのまま有効になります。

---

## サンプルとして同梱：`秀.xlsm`

配布フォルダ直下の `秀.xlsm` は、**作者（shu）が個人利用しているサンプルアドイン**です。
このツールキットの動作確認用に同梱しています。

### サンプルが提供する機能

- `shu001〜005.bas` 内の Public Sub をメニュー自動表示
- アドイン自身を `.xlam` に変換して登録／解除するマクロ（`アドインの更新登録`, `アドインの登録解除`）
- カレンダー入力フォーム、マクロ一覧フォーム などの UserForm 例

### サンプルを試す

1. `秀.xlsm` をダブルクリックで通常ブックとして開く
2. `Alt + F8` でマクロ一覧 → `アドインの更新登録` を実行
3. `%AppData%\Roaming\Microsoft\AddIns\` に `秀.xlam` が保存・登録される
4. Excel を再起動 → リボンに `秀` メニューが現れる

**注意**：このサンプルは作者個人の業務用に組まれたもので、内部マクロは参考程度にご覧ください。
このツールキットの本体はあくまで `作業ファイル/project/python_scripts/` 配下の Python ツール群です。

---

## ファイル構成

```
VBAマネージャー/
├── README.md                       ← このファイル
├── CLAUDE.md                       ← Claude Code 用プロジェクトルール（同梱）
├── LICENSE                         ← MIT
├── requirements.txt
├── 秀.xlsm                         ← サンプルアドイン（個人利用例）
├── .claude/
│   └── skills/
│       └── shu-addin-manager/      ← Claude Code 用の同梱スキル
└── 作業ファイル/
    ├── VBAマネージャー起動.bat
    ├── 一覧.bat
    ├── 修正.bat
    └── project/
        └── python_scripts/         ← ★ ツールキット本体
            ├── vba_manager.py      ← 中核：VBA モジュール CRUD
            ├── form_builder.py     ← UserForm ビルダー
            ├── live_sync_vba.py    ← ライブ同期
            ├── menu_launcher.py    ← GUI ランチャー入口
            ├── select_macro_gui.py ← customtkinter 製 GUI
            ├── bas_editor.py       ← .bas 編集ヘルパー
            ├── format_bas.py       ← .bas 整形
            ├── optimize_vba_modules.py
            └── *.bas / *.frm / *.frx  ← サンプル用 VBA モジュール／フォーム
```

---

## 注意事項

### `.bas` は **必ず CP932 (Shift-JIS)** で保存

UTF-8 のエディタで開いて保存し直すと、日本語が全部文字化けして VBA モジュールが破壊されます。

```python
with open(path, 'r', encoding='cp932') as f:
    content = f.read()

with open(path, 'w', encoding='cp932') as f:
    f.write(content)
```

### バックアップ

`vba_manager.py replace-module` などの破壊的な操作は、事前に `backups/` フォルダへ自動でバックアップを作成します。

---

## トラブルシューティング

| 症状 | 対処 |
|---|---|
| `プログラムによる Visual Basic プロジェクトへのアクセスは信頼性に欠けます` | Excel のトラスト センター設定（手順③）を再確認 |
| `pywin32` インストール後に `ImportError` | `py -m pywin32_postinstall -install` を実行 |
| `.bas` を編集したら全部文字化けした | UTF-8 で保存してしまった。`backups/` から復元 |
| `vba_manager.py list` が空を返す | 対象 `.xlsm` が開いていない or トラストセンター設定が未済 |

---

## ライセンス

MIT License — [LICENSE](LICENSE) を参照してください。

## 作者

shu

このツールキットは個人開発で、**無保証** で公開しています。
バグ報告や改善提案は歓迎ですが、対応を保証するものではありません。
