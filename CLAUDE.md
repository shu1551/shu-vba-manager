# Excel VBA マネージャー — プロジェクトルール

このリポジトリは Excel の VBA マクロ・シート・テーブル・ピボット・パワークエリ等を
Claude Code で会話しながら操作するための **汎用ツールキット**。
**対象は常に「今アクティブに開いている Excel ブック（ActiveWorkbook）」** で、特定のブックに依存しない。
Excel VBA アドイン「秀.xlsm」は、その代表的な利用例として同梱しているにすぎない。
Claude はこの指示に必ず従うこと。

## 対象の原則（最重要・まずこれを当てる）

- `vba_manager.py` / `form_builder.py` は **アクティブな開いているブック** に対して動く。
  引数なし＝アクティブブック、第1引数にパスを渡せばそのファイルを対象にする。
- 特定ブック名・パス・モジュール名（`shu001`/`shu003` 等）に**限定／依存させない**。
- 既存コードに `XLSM_PATH` や `shu001` 等の固有名詞があっても「対象」と鵜呑みにしない。
  **固有名詞より「対象はアクティブブック」を先に当てる**。秀.xlsm を亡霊のように呼び戻さない。

## 作業の前提

- スクリプトは `作業ファイル/project/python_scripts/` から実行する
- Python は `py` コマンドを使う（`python` ではない）
- 操作したいブックを Excel で開いてアクティブにしておく（秀.xlsm はその一例）
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

**1〜数行の小修正は `py vba_manager.py code-replace "旧" "新" --module 名 -y` が最短**
（変更行だけ送る。プロシージャ全文の再送は下のフロー＝新規追加や大改造のときだけ）。

1. `py vba_manager.py list` でマクロ一覧を確認
2. `py vba_manager.py get <Sub名>` で対象コードを取得（_last_proc.vba に出力）
3. `_last_proc.vba` を Read ツールで読み、修正内容を検討
4. 修正後のコードを `_last_proc.vba` に Write
5. `py vba_manager.py replace-procedure -y` で適用（非対話実行では -y 必須）
6. ユーザーに動作確認を依頼

## 取り込み前の検査（check-bas）

手書きで `.bas` を作った場合など、取り込む前に機械的な事故を1コマンドで検査できる（COM不要）：

```bash
py vba_manager.py check-bas <file.bas>         # 文字コード/改行二重化/重複を検査
py vba_manager.py check-bas <file.bas> --fix   # 改行二重化(\r\r\n)だけ自動修正
```

## 守ること

- 指示されていない機能を追加しない
- 指示されていないコードを変更しない
- 影響範囲を確認せずに変更しない

詳しい手順は `.claude/skills/excel-vba-manager/SKILL.md`（汎用ツールの全コマンド・標準フロー）を参照。
秀.xlsm 固有の情報（モジュール構成・フォーム一覧・アドインの仕組み）は
`.claude/skills/shu-addin-manager/SKILL.md` を参照。
