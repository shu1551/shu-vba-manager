# 秀.xlsm VBAマネージャー — プロジェクトルール

このリポジトリは Excel VBA アドイン「秀.xlsm」を Claude Code で
会話しながら操作するためのツールキット。Claude はこの指示に必ず従うこと。

## 作業の前提

- スクリプトは `作業ファイル/project/python_scripts/` から実行する
- Python は `py` コマンドを使う（`python` ではない）
- Excel で 秀.xlsm を開いた状態で操作する
- 事前に Excel のトラストセンターで
  「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしておく

## 絶対禁止事項：.bas / .frm / .cls を Edit / Write ツールで編集するな

これらのファイルは CP932 (Shift-JIS) で保存されている。
Claude の Edit / Write ツールは UTF-8 で書き込むため、CP932 ファイルを
編集すると日本語（モジュール名・プロシージャ名・文字列）が全て文字化けし、
VBA モジュールが破壊される。

.bas ファイルを修正するときは必ず以下のいずれかを使うこと：

1. Python で CP932 のまま読み書き：
   ```python
   with open(path, 'r', encoding='cp932') as f:
       content = f.read()
   content = content.replace(old, new)
   with open(path, 'w', encoding='cp932') as f:
       f.write(content)
   ```
2. vba_manager.py の replace-procedure（_last_proc.vba 経由、プロシージャ単位）

なお `_last_proc.vba` は UTF-8 なので Write ツールで編集してよい。

## マクロ修正の標準フロー

1. `py vba_manager.py list` でマクロ一覧を確認
2. `py vba_manager.py get <Sub名>` で対象コードを取得（_last_proc.vba に出力）
3. `_last_proc.vba` を Read ツールで読み、修正内容を検討
4. 修正後のコードを `_last_proc.vba` に Write
5. `py vba_manager.py replace-procedure` で適用
6. ユーザーに動作確認を依頼

## 守ること

- 指示されていない機能を追加しない
- 指示されていないコードを変更しない
- 影響範囲を確認せずに変更しない

詳しい手順は `.claude/skills/shu-addin-manager/SKILL.md` を参照。
