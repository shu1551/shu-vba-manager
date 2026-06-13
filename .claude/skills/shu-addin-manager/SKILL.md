---
name: shu-addin-manager
description: |
  秀.xlsm（Excel VBAアドインソフト）のマクロ追加・修正・管理を行うスキル。
  秀.xlsm はアドインとして登録し、どのExcelファイルからでもマクロをメニュー表示・実行できる
  ユニークなアドインフレームワーク。shu001モジュールにマクロを追加するだけでメニューに自動反映される。

  このスキルは以下のときに必ず使うこと：
  - ユーザーが「秀」「アドイン」「shu001」「shu003」に言及したとき
  - 「マクロを追加して」「マクロを修正して」「VBAマクロ」などExcel VBAアドイン関連の作業
  - vba_manager.py や form_builder.py を使う作業
  - 秀.xlsm のプロシージャやモジュールの操作
  - Excel VBA のショートカットキー設定やAttribute行の操作
---

# 秀.xlsm アドインマネージャー

## 概要

秀.xlsm は Excel VBA アドインソフト。shu001 モジュールにマクロ（Sub）を書くだけで、
アクティブマクロフォームのメニューに自動表示され、選択して実行できる。
専門知識なしでオリジナルのアドインソフトが作成できるフレームワーク。

## 重要パス

このスキルはリポジトリのルートで `claude` を起動して使う前提。
パスはすべてリポジトリルートからの相対。

```
秀.xlsm:            秀.xlsm（リポジトリルート直下）
SCRIPTS:            作業ファイル\project\python_scripts\
vba_manager.py:     SCRIPTS\vba_manager.py（メイン管理ツール）
form_builder.py:    SCRIPTS\form_builder.py（UserForm作成モジュール）
_last_proc.vba:     SCRIPTS\_last_proc.vba（get の出力先、UTF-8）
backups:            作業ファイル\backups\（自動バックアップ格納場所）
```

標準モジュールの正規ソース（shu001.bas / shu003.bas 等）はリポジトリには含まれない。
必要なときは `export-module` で 秀.xlsm 本体から取り出す。

## 秀.xlsm の構成

- **標準モジュール**: shu001, shu003, Module1
- **主要フォーム**: アクティブマクロフォーム ほか
- 構成はアドインの育て方によって変わる。`list-modules` で実際の状態を確認すること。

## vba_manager.py コマンド一覧

すべてのコマンドは SCRIPTS ディレクトリから実行する。Excelで秀.xlsmを開いた状態で使う。
**コマンドは `py` を使うこと（`python` ではない）。**

```bash
cd 作業ファイル\project\python_scripts

# マクロ一覧表示
py vba_manager.py list

# モジュール一覧表示
py vba_manager.py list-modules

# プロシージャのコード取得 → _last_proc.vba に保存
py vba_manager.py get <Sub名>

# モジュール指定してプロシージャ取得（同名プロシージャが複数ある場合に使う）
py vba_manager.py get <モジュール名> <Sub名>
py vba_manager.py get <モジュール名>.<Sub名>   # ドット区切りも可

# _last_proc.vba の内容でプロシージャを置換（バックアップ自動作成）
py vba_manager.py replace-procedure -y
#  ※ -y/--yes で確認プロンプト(y/N)をスキップ。Claude等の非対話実行では必ず -y を付ける。
#    差分(Diff)は -y を付けても表示される。--module <名> で対象モジュールの明示も可。

# モジュール全体を Remove+Import で置換
py vba_manager.py replace-module <モジュール名> <basファイル>

# モジュールを .bas にエクスポート
py vba_manager.py export-module <モジュール名>
```

## 標準作業フロー（マクロ修正）

この手順を必ず守ること。勝手にコードを変更しない。

1. `py vba_manager.py list` でマクロ一覧を確認
2. `py vba_manager.py get <Sub名>` で対象コードを取得
   - 同名プロシージャが複数フォームにある場合は **モジュール指定** を使う：
     `py vba_manager.py get <モジュール名> <Sub名>`
3. `_last_proc.vba` を Read ツールで読み、修正内容を検討
4. 修正後のコードを `_last_proc.vba` に Write
5. `py vba_manager.py replace-procedure -y` で適用（非対話実行のため -y 必須）
6. ユーザーに動作確認を依頼

## 標準作業フロー（フォームの .bas を修正して適用）

フォームや複数プロシージャをまとめて修正する場合（.bas ファイルを直接編集）：

1. `py vba_manager.py export-module <モジュール名>` で最新の .bas をエクスポート
2. Python スクリプトで CP932 のまま編集する（Edit/Write ツール禁止）：
   ```python
   # 編集スクリプトの雛形
   path = r"作業ファイル\project\python_scripts\対象フォーム.bas"
   with open(path, 'r', encoding='cp932') as f:
       lines = f.readlines()
   # 行番号ベースで修正（多行文字列マッチは使わない）
   # lines[N] = 新しい行内容
   with open(path, 'w', encoding='cp932') as f:
       f.writelines(lines)
   ```
   **行番号は inspect スクリプトで事前確認する：**
   ```python
   with open(path, 'r', encoding='cp932') as f:
       lines = f.readlines()
   for i, line in enumerate(lines, 1):
       if "対象キーワード" in line:
           print(f"L{i}: {repr(line)}")
   ```
3. `py vba_manager.py replace-module <モジュール名> <basファイル>` で適用

## 標準作業フロー（標準モジュール全体の適用）

.bas ファイルを編集してモジュール全体を置換する場合：

1. `py vba_manager.py export-module <モジュール名>` でエクスポート
2. .bas ファイルを Python で CP932 のまま編集
3. `py vba_manager.py replace-module <モジュール名> <basファイル>` で適用

## UserForm の作成・修正

form_builder.py を使う。

```python
from form_builder import FormBuilder, add_btn, add_lbl, add_txt, add_lst, add_combo

with FormBuilder.connect() as fb:  # アクティブブック接続
    frm = fb.get_or_create("FormName", caption="タイトル", width=300, height=200)
    f = fb.clear_controls(frm)
    add_btn(f, "BtnOK", "OK", 80, 160, 60, 20)
    add_lbl(f, "Label1", "テキスト", 10, 10, 100, 18)
    add_txt(f, "TextBox1", 10, 30, 200, 22)
    add_lst(f, "ListBox1", 10, 60, 200, 90)
    fb.inject_vba(frm, "_form_code.vba")  # VBAコードを注入
    fb.save()
```

デフォルトフォントは全コントロール 12pt。

## 絶対に守るべきルール

### 勝手な変更の禁止
- 指示されていない機能を追加しない
- 指示されていないコードを変更しない
- 影響範囲を確認せずに変更しない

### エンコーディング（最重要 - 違反厳禁）
- .bas ファイルは **CP932**（Shift-JIS）で保存
- win32com 経由の文字列は Unicode
- _last_proc.vba は UTF-8

**⚠ .bas ファイルに Edit ツール・Write ツールを絶対に使うな ⚠**
Claude の Edit / Write ツールは UTF-8 で書き込むため、CP932 の .bas ファイルが破壊される。
VBA にインポートするとモジュール名・プロシージャ名・日本語文字列が全て文字化けする。

.bas ファイルを修正するときは必ず以下のいずれかを使うこと：
1. **Python で CP932 のまま読み書き**:
   ```python
   with open(path, 'r', encoding='cp932') as f:
       content = f.read()
   # 修正処理
   with open(path, 'w', encoding='cp932') as f:
       f.write(content)
   ```
2. **_last_proc.vba 経由で replace-procedure**（プロシージャ単位の修正）

### モジュール適用方式
- AddFromString は Attribute 行を正しく処理しないことがある
- **必ず Remove + Import 方式（replace-module）を使うこと**
- ショートカットキーは Attribute VB_ProcData.VB_Invoke_Func で定義される

### replace-module の副作用
- Remove+Import でモジュールが VBComponents の末尾に移動する
- アクティブマクロフォーム等でマクロの表示順が変わる場合がある
- 影響を受けたモジュール（shu003等）も replace-module して順番を揃える

### InsertLines の改行問題
- Python の \n では正しく複数行に分割されないことがある
- .bas ファイル直接編集方式を使うこと

## マウスホイール対応

ListBox を持つフォームではホイールスクロールを有効化する。
shu001 に以下の Public Sub が定義済み：

```vba
' フォームの UserForm_Initialize で呼ぶ
Public Sub ホイール有効化(対象リスト As Object)

' フォームの UserForm_Terminate で呼ぶ
Public Sub ホイール無効化()
```

フォーム側の実装例：
```vba
Private Sub UserForm_Initialize()
    ホイール有効化 ListBox1
End Sub

Private Sub UserForm_Terminate()
    ホイール無効化
End Sub
```

## 変更前の確認チェックリスト

コードを変更する前に必ず以下を確認：

1. 変更対象のフォーム/モジュールが他の機能から参照されていないか
2. AutoFilter の起点列（B1始まりかA1始まりか）
3. 既存のコントロール名と用途（別機能で使われていないか）
4. 変更がユーザーの指示の範囲内か
