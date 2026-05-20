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

- ✅ **VBA モジュール (.bas) の取り出し／差し替え／プロシージャ単位の置換**
- ✅ **UserForm を Python コードで構築**（コントロール配置〜イベントコード注入まで）
- ✅ **VBE と `.bas` ファイルのライブ同期**
- ✅ **登録済みマクロの GUI ランチャー**（customtkinter 製）
- ✅ **任意の `.xlsm` を対象にできる**（自分のブックでも、配布ブックでも）
- ✅ **Claude Code 連携で「話しかけるだけ」開発が可能**

---

## ツール一覧（`作業ファイル/project/python_scripts/`）

| ツール | 役割 |
|---|---|
| **`vba_manager.py`** | 中核ツール。`.xlsm` 内の VBA モジュールを `list / get / replace-module / replace-procedure` で操作 |
| **`form_builder.py`** | UserForm をコードで構築。`add_btn`, `add_lbl`, `add_txt`, `add_lst` などのヘルパー付き |
| **`live_sync_vba.py`** | VBE と `.bas` ファイルをリアルタイム同期 |
| **`menu_launcher.py`** / **`select_macro_gui.py`** | customtkinter 製のマクロ検索／実行 GUI |
| **`bas_editor.py`** | `.bas` ファイル編集ヘルパー（プロシージャ単位の `replace_sub`, `read_bas`） |
| **`format_bas.py`** | `.bas` 整形（空行・インデント正規化） |
| **`optimize_vba_modules.py`** | VBA モジュール最適化 |

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

### ④ Claude Code のインストール（推奨）

<https://claude.com/claude-code> の手順に従ってインストール。
インストール後、配布フォルダで `claude` コマンドを起動するとここを作業ディレクトリとして AI と対話できます。

---

## 主要ツールの使い方

### `vba_manager.py` — VBA モジュール CRUD

```powershell
cd "作業ファイル\project\python_scripts"

# 全モジュール一覧
py vba_manager.py list "C:\path\to\任意.xlsm"

# .bas を取り出し
py vba_manager.py get "C:\path\to\任意.xlsm" モジュール名

# .bas モジュールを置き換え（破壊的、自動バックアップあり）
py vba_manager.py replace-module "C:\path\to\任意.xlsm" モジュール名 ファイル.bas

# プロシージャ単位で置き換え（既存プロシージャをピンポイント差し替え）
py vba_manager.py replace-procedure "C:\path\to\任意.xlsm" patched.vba
```

### `form_builder.py` — UserForm をコードで構築

```python
from form_builder import FormBuilder, add_btn, add_lbl, add_txt, add_lst

with FormBuilder.connect(wb_keyword='ターゲットブック名') as fb:
    fb.create_form('SampleForm', width=300, height=200)
    add_lbl(fb, 'lblTitle', 'タイトル', x=10, y=10)
    add_txt(fb, 'txtInput', x=10, y=40, width=200)
    add_btn(fb, 'btnOK', 'OK', x=10, y=80)
```

これで `SampleForm` という名前の UserForm がターゲットブックに作成されます。
さらにイベントコード（`btnOK_Click` など）も Python から注入可能。

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
