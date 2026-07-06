---
name: excel-userform-builder
description: >
  Excel VBA UserForm の作成・修正・再構築を行うスキル。form_builder.py モジュールを使って
  Python から UserForm のコントロール配置と VBA コード注入を自動化する。
  ユーザーが「フォームを作って」「UserForm を追加して」「フォームを修正して」「ボタンを追加して」
  「Excel のフォーム」「倍率回転フォーム」「カレンダーフォーム」など、Excel UserForm に関する
  作成・修正・再構築の話をしたときに必ずこのスキルを使うこと。
  フォームのコントロール配置だけでなく、VBA イベントコードの作成も含む。
---

# Excel UserForm Builder

Excel VBA の UserForm を Python (win32com) 経由で作成・修正するためのガイド。

## 前提条件

- 対象の .xlsm を Excel で開いて**アクティブ**にしておくこと（対象は常に**アクティブな開いているブック**。特定ブック名・パスに依存しない）
- ツール（`form_builder.py` / `vba_manager.py` / `form_inspect.py`）は作業中プロジェクトの `…\作業ファイル\project\python_scripts\` にある。**絶対パスを決め打ちしない** — 場所が不明なら Glob `**/form_builder.py` で特定し、そのディレクトリで実行する
- Python は `py` コマンドを使う（`python` ではない）

## まず最短ルートを選ぶ（重要）

UserForm の作業は2種類。混同すると遠回りになる。

### A. 既存フォームの「コードだけ」直す（イベント処理の修正・バグ取り）
**form_builder で作り直さない。** vba_manager のプロシージャ単位置換が最短：

1. 把握： `py form_inspect.py <フォーム名> [proc ...]`
   → **1回のCOM接続**でコントロール配置＋VBAコード（全体 or 指定プロシージャ）をまとめて出す。
     export-module と get を別々に何度も叩く（毎回接続し直し＝遅い）のを避ける入口。
   → `py form_inspect.py --list` で開いているブックの全フォーム名。
2. 取得： `py vba_manager.py get <フォーム名>.<プロシージャ>` → `_last_proc.vba`(UTF-8) に保存
3. 修正版を **UTF-8 の .vba** に書く（**.bas は絶対に Edit/Write しない**）
4. 置換： `py vba_manager.py replace-procedure --code-file fix.vba --module <フォーム名> -y`
   - **ブック保存まで自動**。Attribute 行の無い普通のプロシージャは InsertLines＝**速い経路**。
   - 同名プロシージャが複数モジュールにある時だけ `--module` が必須。

### B. コントロールの配置・追加・全面再構築
レイアウトを変える／コントロールを足す／作り直すとき**だけ**、下の form_builder フロー
（get_or_create → clear_controls → 再配置 → inject_vba）を使う。

## デザイン原則（AI がフォームを設計するときの規範）

MSForms の素の見た目は古いが、**整列・余白・統一・順序**の4つが揃うだけで見違える。
座標を暗算で一発置きせず、以下を規範として設計し、必ず PNG で検証する。

### 寸法のリズム（8の倍数で刻む）
- フォーム外周の余白: **12pt**。要素間の縦間隔: **8pt**（強い関連は 6、セクション区切りは 16）
- 高さの標準: TextBox/ComboBox **22**、ボタン **24〜28**、ラベル **18**、リスト行数×12+6
- **同種のコントロールは同じ幅・高さに揃える**（特にボタン。バラバラの高さは素人臭さの最大要因）

### 整列（縦ラインを最少に）
- 基本は「ラベル列＋入力列」の**2本の縦ライン**。左端が3本以上あると散らかって見える
- ラベルは入力欄と**上下中央**を合わせる（top = 入力top + (入力H − ラベルH)/2）
- 右端も揃える（一番幅の広い入力に合わせる）

### ボタンの作法（Windows 規約）
- OK/キャンセルは**右下・同サイズ**（最低 72×24）。並びは「実行系 → キャンセルが右端」
- Default=True（Enter）を実行ボタンに、Cancel=True（Esc）をキャンセルに設定する
- 破壊的操作（削除等）のボタンは OK 群から**離して**置く（左下など）

### フォント・色（増やさない）
- 全コントロール **12pt 統一**（ツールの既定）。強調は **Bold のみ**。フォント種は増やさない
- 色は最小限: 背景は既定のまま、強調色は**1色まで**。多色は判断コストを増やすだけ

### 順序と操作
- **TabIndex は視線順**（左→右、上→下）。配置後に必ず整える
- 主要ボタンに Accelerator（&文字）を設定
- UserForm_Initialize で最初の入力欄に SetFocus

### 命名
- `btn/txt/lbl/lst/cmb/chk/opt/fra` ＋意味名（btnSave、txtName）。イベントプロシージャ名と一致させる

### 検証ループ（必須・省略禁止）
1. 作る（form_layout / form_builder）
2. **`py form_inspect.py <フォーム> --png` で実表示を見る**（座標表ではなく画像で判断）
3. `--lint` の機械検査（重なり・はみ出し・不揃い）を通す
4. ズレは form_tool（set/align/scale）で直して再撮影
「一発で置けたはず」は禁物——必ず目で確認してから納品する。

## 定番レシピ（form_layout の宣言をそのまま流用してよい）

### 1. 確認ダイアログ
```python
build_form("F_Confirm", "確認", rows=[
    row(lbl("この操作は元に戻せません。実行しますか？")),
    spacer(),
    button_bar(ok("btnYes", "実行する"), cancel("btnNo", "やめる")),
], vba_stub=True)
```

### 2. 検索フォーム
```python
build_form("F_Search", "検索", rows=[
    row(lbl("キーワード"), txt("txtKey", required=True), btn("btnFind", "検索", accel="F")),
    row(lbl("結果"), lst("lstResult", rows_visible=10)),
    spacer(),
    button_bar(btn("btnJump", "移動"), cancel("btnClose", "閉じる")),
], vba_stub=True)
```

### 3. 範囲を選んで実行（シート処理ツールの型）
```python
build_form("F_Range", "範囲処理", rows=[
    row(lbl("対象範囲"), refedit("refTarget")),   # TextBox+選択ボタン+InputBox(Type:=8) の複合部品
    row(lbl("件数"), spin_txt("txtN", value="1", min_=1, max_=999), lbl("件")),
    spacer(),
    button_bar(ok("btnRun", "実行"), cancel("btnClose", "閉じる")),
], vba_stub=True)
```

### 4. 設定画面（タブ付き）
```python
build_form("F_Setting", "設定", rows=[
    multipage("mpMain",
        page("基本", row(lbl("名前"), txt("txtName", required=True)),
                     row(lbl("区分"), combo("cmbKind", items=["A", "B"]))),
        page("詳細", row(lbl("メモ"), txt("txtMemo", height=54, multiline=True)),
                     row(chk("chkAdv", "詳細モードを有効にする")))),
    spacer(),
    button_bar(ok("btnSave", "保存", accel="S"), cancel("btnCancel", "キャンセル")),
], vba_stub=True, png=True)
```

いずれも `preview_layout(rows)` で Excel 無しの配置確認 → `build_form(..., png=True)` →
`form_inspect --lint` → `form_tool` 微調整、のループで仕上げる。
required=True はラベルに ＊ が付き、実行ボタンに空チェック雛形が入る。

## 作業フロー

### 0. まず form_layout.py を検討する（推奨・宣言的レイアウト）

行構造で書けるフォームなら、座標計算を自前でやらず form_layout.py に任せる
（ラベル列幅の自動揃え・余白リズム・ボタンバー・TabIndex 整列が全部機械計算になる）。
使い方は excel-vba-manager スキルまたは form_layout.py の docstring を参照。
自由配置（カレンダー格子等）だけ form_builder の add_* + Grid を直接使う。

### 1. 要件を確認する

- フォーム名、キャプション、サイズ
- 必要なコントロール（ボタン、テキストボックス、リスト等）とレイアウト
- 各コントロールのイベント処理（クリック時の動作等）
- 対象ブック（秀.xlsm、ポスター.xlsm 等）

### 2. Python スクリプトを作成する

`create_<フォーム名>.py` として保存する。基本構造:

```python
"""
<フォーム名>を<ブック名>に作成するスクリプト
実行前に<ブック名>をExcelで開いておくこと。
"""
import os
import traceback
from form_builder import FormBuilder, add_btn, add_lbl, add_txt, add_lst, add_combo

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VBA_FILE = os.path.join(BASE_DIR, "_<フォーム名>_code.vba")


def build_controls(f):
    # コントロールを配置
    add_btn(f, "BtnOK", "OK", 100, 160, 60, 20)
    # ... 他のコントロール


def main():
    try:
        with FormBuilder.connect() as fb:   # 既定はアクティブブック（特定ブックに縛らない）
            frm = fb.get_or_create("フォーム名",
                                   caption="タイトル", width=300, height=200)
            f = fb.clear_controls(frm)
            build_controls(f)
            print(f"コントロール合計: {f.Controls.Count} 個")
            fb.inject_vba(frm, VBA_FILE)
            fb.save()
    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    main()
```

### 3. VBA コードファイルを作成する

`_<フォーム名>_code.vba` として **UTF-8** で保存する。フォームのイベントプロシージャを記述:

```vba
Private Sub CommandButton1_Click()
    ' ボタンクリック時の処理
End Sub

Private Sub UserForm_Initialize()
    ' フォーム初期化処理
End Sub
```

### 4. スクリプトを実行する

```bash
# form_builder.py のあるディレクトリで実行（py コマンド・対象はアクティブブック）
py create_<フォーム名>.py
```

## form_builder.py API リファレンス

### 接続方法

```python
FormBuilder.connect(wb_path=r"C:\...\file.xlsm")   # フルパス指定
FormBuilder.connect(wb_keyword="ポスター")           # ブック名キーワード検索
FormBuilder.connect()                                # アクティブブック
```

### FormBuilder メソッド

| メソッド | 説明 |
|---------|------|
| `get_or_create(name, caption=, width=, height=)` | フォームを取得 or 新規作成 |
| `delete_form(name)` | フォームを削除 |
| `clear_controls(frm)` | 全コントロール削除、Designer を返す |
| `inject_vba(frm, vba_file)` | UTF-8 の .vba ファイルからコード注入 |
| `save()` | ブックを保存 |
| `list_forms()` | フォーム名一覧を返す |

### コントロール追加関数

全関数の共通引数: `(frm, name, left, top, width, height)`
デフォルトフォントサイズ: 12pt（`font_size=` で個別変更可）

| 関数 | 追加引数 | 説明 |
|------|---------|------|
| `add_btn(frm, name, caption, ...)` | `font_bold=False` | CommandButton |
| `add_lbl(frm, name, caption, ...)` | `align=1, fore=, back=, font_bold=` | Label (align: 1左,2中央,3右) |
| `add_txt(frm, name, ...)` | `value="", multiline=False, scrollbars=0` | TextBox |
| `add_lst(frm, name, ...)` | | ListBox |
| `add_combo(frm, name, ...)` | | ComboBox |
| `add_chk(frm, name, caption, ...)` | | CheckBox |
| `add_opt(frm, name, caption, ...)` | | OptionButton |
| `add_frame(frm, name, caption, ...)` | ※font_size なし | Frame |
| `add_img(frm, name, ...)` | ※font_size なし | Image |
| `add_spin(frm, name, ...)` | ※font_size なし | SpinButton |
| `add_scroll(frm, name, ...)` | `orientation=0` (0水平,1垂直) | ScrollBar |

## 重要な注意事項

- **VBA コードファイルのエンコーディング**: `.vba` ファイルは必ず **UTF-8** で保存する（`inject_vba` が UTF-8 で読む）
- **.bas ファイルのエンコーディング**: 標準モジュールの `.bas` は **CP932 (Shift-JIS)**
- **⚠ .bas ファイルに Edit ツール・Write ツールを絶対に使うな ⚠**: Claude の Edit/Write は UTF-8 で書き込むため CP932 が破壊される。修正は必ず Python で `open(path, 'r', encoding='cp932')` → 修正 → `open(path, 'w', encoding='cp932')` とする
- **初回作成時のリネーム失敗**: VBComponents.Add(3) 後の Name 設定が失敗することがある → UserFormN になった場合は手動リネーム後にスクリプト再実行
- **既存フォームの修正**: `get_or_create` は既存フォームを再利用し、`clear_controls` で全コントロールを削除してから再配置する
- **コントロール名**: VBA イベントプロシージャ名と一致させること（例: `CommandButton1` → `Private Sub CommandButton1_Click()`）

## 既存の実装例

`create_*.py` / `_*_code.vba` 形式の作成スクリプトが**プロジェクト内に残っていれば**参考にする
（例：カレンダーフォーム、ノード編集フォーム）。ただし**存在を前提にしない** — 無ければ上の
雛形（「作業フロー」§2）から書き起こす。あるか確かめるときは Glob `**/create_*.py` で探す。
