# -*- coding: utf-8 -*-
"""COM 不要の純粋ロジックの自動テスト。

実行: このフォルダで `py -m pytest test_tools.py -v`
Excel には一切接続しない（COM 依存部は実機テストで担保）。
今日踏んだバグの回帰テストを含む:
  - End Sub 末尾コメントのブロック境界
  - frame の高さ計算（キャプション帯ぶんのクリップ）
  - multipage のコンテンツ幅計算漏れ
  - 改行二重化の多層ガード
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import vba_manager as vm
import form_layout as fl
import form_inspect as fi


# ================================================================
# vba_manager: 改行・エンコーディングの多層ガード
# ================================================================

def test_normalize_bas_newlines_idempotent(tmp_path):
    p = tmp_path / "a.bas"
    p.write_bytes("Sub A()\r\nEnd Sub\r\n".encode('cp932'))
    fixed, raw, was = vm.normalize_bas_newlines(str(p))
    assert not was
    assert fixed == raw


def test_normalize_bas_newlines_fixes_doubling(tmp_path):
    p = tmp_path / "a.bas"
    p.write_bytes("Sub A()\r\r\nEnd Sub\r\r\n".encode('cp932'))
    fixed, raw, was = vm.normalize_bas_newlines(str(p))
    assert was
    assert fixed == "Sub A()\r\nEnd Sub\r\n".encode('cp932')


def test_read_code_file_collapses_doubling(tmp_path):
    p = tmp_path / "a.vba"
    p.write_bytes('Sub A()\r\r\n    MsgBox "x"\r\r\nEnd Sub\r\r\n'.encode('cp932'))
    assert vm.read_code_file(str(p)) == 'Sub A()\n    MsgBox "x"\nEnd Sub\n'


def test_validate_bas_encoding_rejects_utf8_japanese(tmp_path):
    p = tmp_path / "a.bas"
    p.write_bytes("Sub あ()\r\nEnd Sub\r\n".encode('utf-8'))
    assert vm.validate_bas_encoding(str(p)) is False


def test_validate_bas_encoding_accepts_cp932(tmp_path):
    p = tmp_path / "a.bas"
    p.write_bytes("Sub あ()\r\nEnd Sub\r\n".encode('cp932'))
    assert vm.validate_bas_encoding(str(p)) is True


def test_duplicate_procedures_detected():
    code = "Sub A()\nEnd Sub\nSub A()\nEnd Sub\n"
    assert "A" in vm._find_duplicate_procedures(code)


def test_parse_module_blocks_end_sub_comment():
    # End Sub の末尾コメントでブロック境界がずれないこと（隣Sub消失バグの回帰）
    bas = ('Attribute VB_Name = "M"\r\n'
           "Sub A()\r\n"
           "End Sub ' comment\r\n"
           "Sub B()\r\n"
           "End Sub\r\n")
    header, blocks, trailing = vm._parse_module_blocks(bas)
    assert [b['name'] for b in blocks] == ['A', 'B']


def test_looks_like_xl_file():
    assert vm.looks_like_xl_file("a.xlsm")
    assert vm.looks_like_xl_file(r"C:\path\b.xlam")
    assert not vm.looks_like_xl_file("fix.vba")          # 位置引数の罠の回帰
    assert not vm.looks_like_xl_file(r"C:\tmp\fix.vba")


def test_coerce_cell():
    assert vm._coerce_cell("12") == 12
    assert vm._coerce_cell("1.5") == 1.5
    assert vm._coerce_cell("=SUM(A1)") == "=SUM(A1)"
    assert vm._coerce_cell("nan") == "nan"     # Excel でエラー値化するため文字列のまま
    assert vm._coerce_cell("inf") == "inf"
    assert vm._coerce_cell("") is None


# ================================================================
# form_layout: レイアウト計算の不変条件
# ================================================================

def _std_rows():
    return [
        fl.row(fl.lbl("名前"), fl.txt("txtName")),
        fl.row(fl.lbl("区分"), fl.combo("cmbKind")),
        fl.spacer(),
        fl.button_bar(fl.ok("btnOK"), fl.cancel("btnCancel")),
    ]


def test_layout_alignment():
    pl, cw, ch = fl.compute_layout(_std_rows())
    lefts = {l for e, l, t, w, h in pl if e['kind'] in ('txt', 'combo')}
    rights = {l + w for e, l, t, w, h in pl if e['kind'] in ('txt', 'combo')}
    assert len(lefts) == 1, "入力の左端は1本に揃う"
    assert len(rights) == 1, "入力の右端は1本に揃う"


def test_layout_button_bar_uniform_and_right_aligned():
    pl, cw, ch = fl.compute_layout(_std_rows())
    sizes = {(w, h) for e, l, t, w, h in pl if e['kind'] == 'btn'}
    assert len(sizes) == 1, "ボタンバーは同サイズ"
    right = max(l + w for e, l, t, w, h in pl if e['kind'] == 'btn')
    assert right == fl.STYLE['pad'] + cw, "ボタンバーはコンテンツ右端に揃う"


def test_frame_children_fit():
    # frame 内の最終行がクリップされない（frame_top 足し忘れバグの回帰）
    rows = [fl.frame("G",
                     fl.row(fl.lbl("A"), fl.txt("t1")),
                     fl.row(fl.chk("c1", "チェック")))]
    pl, cw, ch = fl.compute_layout(rows)
    fe, l, t, w, h = pl[0]
    for ce, cl, ct, cw2, ch2 in fe['children']:
        assert ct + ch2 <= h, f"{ce.get('name')} が frame からはみ出している"


def test_multipage_width_included():
    # multipage がコンテンツ幅の自動計算に入っている（計算漏れバグの回帰）
    rows = [fl.multipage("mp",
                         fl.page("P1", fl.row(fl.lbl("ラベル"), fl.txt("t", width=200))))]
    pl, cw, ch = fl.compute_layout(rows)
    assert cw >= 200


def test_required_star_on_label():
    rows = [fl.row(fl.lbl("名前"), fl.txt("txtName", required=True))]
    pl, cw, ch = fl.compute_layout(rows)
    labels = [e for e, *_ in pl if e['kind'] == 'lbl']
    assert any('＊' in (e.get('caption') or '') for e in labels)


def test_stub_required_spin_cancel(tmp_path):
    rows = [
        fl.row(fl.lbl("名前"), fl.txt("txtName", required=True)),
        fl.row(fl.lbl("数"), fl.spin_txt("txtQty")),
        fl.button_bar(fl.ok("btnGo", "実行"), fl.cancel("btnClose")),
    ]
    out = fl.generate_vba_stub(rows, str(tmp_path / "s.vba"))
    code = open(out, encoding='utf-8').read()
    assert 'If Trim(txtName.Value) = ""' in code, "必須チェックの雛形"
    assert "txtQtySpin_Change" in code, "スピン連動イベント"
    assert "Unload Me" in code, "キャンセルの雛形"


def test_refedit_is_composite_with_pick_stub(tmp_path):
    rows = [fl.row(fl.lbl("範囲"), fl.refedit("refX")),   # refedit は複合部品（rowが展開）
            fl.button_bar(fl.ok("btnGo"))]
    out = fl.generate_vba_stub(rows, str(tmp_path / "s.vba"))
    code = open(out, encoding='utf-8').read()
    assert "Application.InputBox" in code, "範囲選択ハンドラの雛形"


# ================================================================
# form_inspect: lint とリバースの機械判定
# ================================================================

def _ctl(name, type_, l, t, w, h, parent="F", **kw):
    d = dict(name=name, type=type_, left=l, top=t, width=w, height=h,
             caption=kw.pop('caption', None), font_size=kw.pop('font_size', 12.0),
             parent=parent, tab_index=kw.pop('tab_index', None),
             bold=kw.pop('bold', None))
    d.update(kw)
    return d


def test_lint_detects_overlap_and_out_of_bounds():
    info = {"inside_width": 200, "inside_height": 100}
    ctrls = [
        _ctl("a", "TextBox", 10, 10, 100, 22),
        _ctl("b", "TextBox", 50, 12, 100, 22),      # a と重なる
        _ctl("c", "TextBox", 150, 90, 100, 22),     # 右下にはみ出す
    ]
    findings = fi.lint_form("F", info, ctrls)
    assert any("重なり" in s for s in findings)
    assert any("はみ出し" in s for s in findings)


def test_lint_clean_form_has_no_findings():
    info = {"inside_width": 300, "inside_height": 200}
    ctrls = [
        _ctl("lbl1", "Label", 12, 14, 40, 18),
        _ctl("txtA", "TextBox", 60, 12, 200, 22),
        _ctl("lbl2", "Label", 12, 44, 40, 18),
        _ctl("txtB", "TextBox", 60, 42, 200, 22),
    ]
    assert fi.lint_form("F", info, ctrls) == []


def test_lint_orphan_handler():
    info = {"inside_width": 300, "inside_height": 200}
    ctrls = [_ctl("btnGo", "CommandButton", 10, 10, 72, 24,
                  default=True, cancel=True, accelerator=None)]
    code = ("Private Sub btnGone_Click()\nEnd Sub\n"
            "Private Sub btnGo_Click()\nEnd Sub\n")
    findings = fi.lint_form("F", info, ctrls, code=code)
    assert any("孤児ハンドラ" in s and "btnGone" in s for s in findings)
    assert not any("btnGo'" in s for s in findings)


def test_lint_missing_click_handler():
    info = {"inside_width": 300, "inside_height": 200}
    ctrls = [_ctl("btnGo", "CommandButton", 10, 10, 72, 24,
                  default=True, cancel=True, accelerator=None)]
    findings = fi.lint_form("F", info, ctrls, code="Private Sub UserForm_Initialize()\nEnd Sub\n")
    assert any("Click ハンドラ未実装" in s for s in findings)


def test_type_normalize():
    assert fi._normalize_type("IMdcText") == "TextBox"
    assert fi._normalize_type("ILabelControl") == "Label"
    assert fi._normalize_type("IMultiPage") == "MultiPage"
    assert fi._normalize_type("ICommandButton") == "CommandButton"
    assert fi._normalize_type("TextBox") == "TextBox"


def test_cluster_rows_groups_by_center():
    items = [
        _ctl("lbl1", "Label", 12, 14, 40, 18),      # 中央 23
        _ctl("txtA", "TextBox", 60, 12, 200, 22),   # 中央 23 → 同じ行
        _ctl("txtB", "TextBox", 60, 42, 200, 22),   # 別の行
    ]
    rows = fi._cluster_rows(items)
    assert len(rows) == 2
    assert [c["name"] for c in rows[0]] == ["lbl1", "txtA"]
