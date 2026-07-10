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


# ================================================================
# vba_manager: 健康診断（checkup）強化まわりの純粋ロジック
# ================================================================

def _mod(name, type_, code, procs=()):
    return {'name': name, 'type': type_, 'type_name': '',
            'total_lines': code.count('\r\n'), 'procs': list(procs), 'code': code}


def test_strip_vba_comment_keeps_quote_in_string():
    assert vm._strip_vba_comment('x = "a\'b" \' comment') == 'x = "a\'b" '
    assert vm._strip_vba_comment("' 全部コメント") == ""


def test_extra_scans_hardcoded_path_and_comment_excluded():
    code = ('Sub A()\r\n'
            '    p = "C:\\data\\in.csv"\r\n'
            '    \' 例: "C:\\old\\path"\r\n'
            '    On Error Resume Next\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 5}])]}
    res = vm._extra_code_scans(inv)
    assert [(m, p, path) for m, p, _, path in res['hardcoded_paths']] == \
        [('M1', 'A', 'C:\\data\\in.csv')]
    assert [(m, p) for m, p, _ in res['error_resume']] == [('M1', 'A')]
    assert res['no_option_explicit'] == ['M1']


def test_extra_scans_auto_exec_and_option_explicit():
    code = 'Option Explicit\r\nPrivate Sub Workbook_Open()\r\nEnd Sub\r\n'
    inv = {'modules': [_mod('ThisWorkbook', 100, code,
                            [{'name': 'Workbook_Open', 'lines': 3}]),
                       _mod('Module1', 1, 'Sub Auto_Open()\r\nEnd Sub\r\n',
                            [{'name': 'Auto_Open', 'lines': 2}])]}
    res = vm._extra_code_scans(inv)
    assert {n for _, n, _ in res['auto_exec']} == {'Workbook_Open', 'Auto_Open'}
    assert res['no_option_explicit'] == ['Module1']


def test_checkup_diff_line_shift_is_not_a_change():
    prev = {'time': '2026-07-03 23:00', 'keys': ['未解決Call: [M1] 親 → 子'],
            'procs': {'M1': ['親']}, 'total_lines': 100, 'sheets': ['S1'], 'forms': []}
    cur = {'time': '2026-07-04 09:00', 'keys': ['未解決Call: [M1] 親 → 子'],
           'procs': {'M1': ['親']}, 'total_lines': 100, 'sheets': ['S1'], 'forms': []}
    assert vm._checkup_diff(prev, cur)['changed'] is False


def test_checkup_diff_detects_new_resolved_and_growth():
    prev = {'time': 't0', 'keys': ['重複プロシージャ: [M1] A'],
            'procs': {'M1': ['A']}, 'total_lines': 50, 'sheets': ['S1'], 'forms': []}
    cur = {'time': 't1', 'keys': ['未解決Call: [M2] B → 消えた子'],
           'procs': {'M1': ['A'], 'M2': ['B']}, 'total_lines': 80,
           'sheets': ['S1', 'S2'], 'forms': ['F_New']}
    d = vm._checkup_diff(prev, cur)
    assert d['changed']
    assert d['new'] == ['未解決Call: [M2] B → 消えた子']
    assert d['resolved'] == ['重複プロシージャ: [M1] A']
    assert d['procs_added'] == ['[M2] B']
    assert d['lines_delta'] == 30
    assert d['sheets_added'] == ['S2'] and d['forms_added'] == ['F_New']


def test_checkup_diff_first_run_returns_none():
    assert vm._checkup_diff(None, {'keys': [], 'procs': {}, 'total_lines': 0,
                                   'sheets': [], 'forms': []}) is None


def test_extra_scans_destructive_and_no_restore():
    code = ('Sub 掃除()\r\n'
            '    Application.ScreenUpdating = False\r\n'
            '    Kill "C:\\tmp\\old.txt"\r\n'
            '    Worksheets("作業").Delete\r\n'
            '    ActiveSheet.Rows("2:" & lastRow).Delete\r\n'
            'End Sub\r\n'
            'Sub 正常()\r\n'
            '    Application.ScreenUpdating = False\r\n'
            '    Application.ScreenUpdating = True\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code,
                            [{'name': '掃除', 'lines': 6}, {'name': '正常', 'lines': 4}])]}
    res = vm._extra_code_scans(inv)
    labels = {(p, lab) for _, p, _, lab, _ in res['destructive']}
    assert ('掃除', 'ファイル/フォルダ削除') in labels
    assert ('掃除', 'シート削除') in labels
    assert ('掃除', '行/列の削除') in labels
    # 行削除の行が「シート削除」と誤ラベルされないこと
    row_line_labels = {lab for _, _, ln, lab, _ in res['destructive'] if ln == 5}
    assert row_line_labels == {'行/列の削除'}
    assert [(m, p) for m, p, _ in res['no_restore']] == [('M1', '掃除')]


def test_extra_scans_no_restore_ignores_comment_and_string():
    code = ('Sub A()\r\n'
            '    \' Application.ScreenUpdating = False と書いたコメント\r\n'
            '    s = "Kill されそうな文字列"\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 4}])]}
    res = vm._extra_code_scans(inv)
    assert res['no_restore'] == [] and res['destructive'] == []


def test_extra_scans_enableevents_no_restore():
    code = ('Sub 事故りがち()\r\n'
            '    Application.EnableEvents = False\r\n'
            'End Sub\r\n'
            'Sub 正しい()\r\n'
            '    Application.EnableEvents = False\r\n'
            '    Application.EnableEvents = True\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': '事故りがち', 'lines': 3},
                                            {'name': '正しい', 'lines': 4}])]}
    res = vm._extra_code_scans(inv)
    assert [(m, p) for m, p, _ in res['no_restore']] == [('M1', '事故りがち')]
    assert 'EnableEvents' in res['no_restore'][0][2]


def test_checkup_rating():
    assert vm._checkup_rating(0, 0) == "A（異常なし）"
    assert vm._checkup_rating(5, 0) == "B（軽度所見）"
    assert vm._checkup_rating(5, 2) == "C（要確認）"


def test_analyze_calls_declared_api_is_not_unresolved():
    code = ('Private Declare PtrSafe Sub MoveMemory Lib "kernel32" '
            'Alias "RtlMoveMemory" (d As LongPtr, s As LongPtr, ByVal n As LongPtr)\r\n'
            'Sub A()\r\n'
            '    Call MoveMemory(1, 2, 3)\r\n'
            '    Call 存在しない子\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 4}])]}
    res = vm._analyze_calls(inv)
    assert [u[2] for u in res['unresolved']] == ['存在しない子']


def test_analyze_calls_object_method_and_qualified_call():
    code = ('Sub A()\r\n'
            '    Call wCompo.Export(sFilePath)\r\n'          # オブジェクトのメソッド＝対象外
            '    Call M1.子マクロ\r\n'                        # モジュール修飾の実マクロ＝辺
            'End Sub\r\n'
            'Sub 子マクロ()\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 4},
                                            {'name': '子マクロ', 'lines': 2}])]}
    res = vm._analyze_calls(inv)
    assert res['unresolved'] == []
    assert res['edges'][('M1', 'A')] == {'子マクロ'}


def test_analyze_calls_dynamic_run_not_unresolved():
    code = ('Sub メニュー実行()\r\n'
            '    Application.Run "\'PERSONAL.XLSB\'!" & AAA\r\n'   # 動的＝未解決にしない
            '    Application.Run "\'" & ZZZ & "\'!" & AAA\r\n'     # 動的
            '    Application.Run "実在マクロ"\r\n'                  # 静的＝辺
            '    Application.Run "居ないマクロ"\r\n'                # 静的＝未解決
            'End Sub\r\n'
            'Sub 実在マクロ()\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'メニュー実行', 'lines': 6},
                                            {'name': '実在マクロ', 'lines': 2}])]}
    res = vm._analyze_calls(inv)
    assert [u[2] for u in res['unresolved']] == ['Run "居ないマクロ"']
    assert len(res['dynamic_runs']) == 2
    assert res['edges'][('M1', 'メニュー実行')] >= {'実在マクロ'}


def test_analyze_calls_commented_run_is_ignored():
    # コメントアウトされた Application.Run を未解決Callにしない（総点検で発見）
    code = ('Sub A()\r\n'
            "    ' Application.Run \"居ないマクロ\"  ←コメント行\r\n"
            '    x = "文字列の中の Application.Run も無視"\r\n'
            'End Sub\r\n')
    inv = {'modules': [_mod('M1', 1, code, [{'name': 'A', 'lines': 4}])]}
    res = vm._analyze_calls(inv)
    assert res['unresolved'] == []
    assert res['dynamic_runs'] == []


def test_checkup_diff_survives_missing_fields():
    # 履歴ファイルは外部データ＝フィールド欠損でも落ちない
    d = vm._checkup_diff({'time': 't0'}, {'time': 't1'})
    assert d['changed'] is False


def test_analyze_calls_onaction_macro_is_not_orphan():
    inv = {'modules': [_mod('M1', 1,
                            'Sub ボタン処理()\r\nEnd Sub\r\nSub 本当の孤立()\r\nEnd Sub\r\n',
                            [{'name': 'ボタン処理', 'lines': 2},
                             {'name': '本当の孤立', 'lines': 2}])],
           'onaction': [('メニュー', '角丸四角形 1', 'ボタン処理')]}
    res = vm._analyze_calls(inv)
    assert res['onaction'] == {'ボタン処理': ['メニュー/角丸四角形 1']}
    assert [n for _, n in res['orphans']] == ['本当の孤立']


def test_validate_vba_code_single_line_sub():
    # 1行書き Sub x(): End Sub を「End Sub 不足」と誤警告しない
    assert vm.validate_vba_code('Sub x(): End Sub\n') is True
    # 文字列内の ':' で誤分割しない
    assert vm.validate_vba_code('Sub y()\n    s = "a:b"\nEnd Sub\n') is True
    # 本当に対応が取れていないものは引き続き検出する
    assert vm.validate_vba_code('Sub z()\n', force=False) is False


# ================================================================
# 2026-07-10 総点検の回帰テスト
# ================================================================

def test_parse_module_blocks_one_liner_and_property():
    # 1行完結 Sub（＋随伴Attribute）が次のプロシージャを巻き込まないこと
    bas = ('Attribute VB_Name = "M"\r\n'
           'Sub S1(): Call Main: End Sub\r\n'
           'Attribute S1.VB_ProcData.VB_Invoke_Func = "q\\n14"\r\n'
           'Sub Main()\r\n'
           'End Sub\r\n'
           'Property Get V() As Long\r\n'
           '    V = 1\r\n'
           'End Property\r\n'
           'Sub Last()\r\n'
           'End Sub\r\n')
    header, blocks, trailing = vm._parse_module_blocks(bas)
    assert [b['name'] for b in blocks] == ['S1', 'Main', 'V', 'Last']
    assert [b['kind'] for b in blocks] == ['sub', 'sub', 'property get', 'sub']
    assert any('VB_Invoke_Func' in ln for ln in blocks[0]['lines'])
    assert vm._write_module(header, blocks, trailing) == bas


def test_analyze_calls_run_module_qualified_resolves():
    # Application.Run "モジュール名.マクロ名" が未解決（C判定）に落ちないこと
    inv = {'modules': [
        _mod('M1', 1, 'Sub A()\r\n    Application.Run "M2.B"\r\nEnd Sub\r\n',
             [{'name': 'A', 'lines': 3}]),
        _mod('M2', 1, 'Sub B()\r\nEnd Sub\r\n', [{'name': 'B', 'lines': 2}]),
    ], 'onaction': []}
    res = vm._analyze_calls(inv)
    assert not res['unresolved'], res['unresolved']
    assert 'B' in res['edges'][('M1', 'A')]


def test_analyze_calls_onaction_module_qualified_resolves():
    # OnAction の「モジュール名.マクロ名」形式も孤立扱いしないこと
    inv = {'modules': [_mod('M1', 1, 'Sub C()\r\nEnd Sub\r\n',
                            [{'name': 'C', 'lines': 2}])],
           'onaction': [('メニュー', '四角形 1', 'M1.C')]}
    res = vm._analyze_calls(inv)
    assert res['onaction'] == {'C': ['メニュー/四角形 1']}
    assert not res['orphans']


def test_layout_button_bar_fits_content_width():
    # 幅の違うボタン（OK/キャンセル）でもバーが左にあふれないこと
    rows = [fl.button_bar(fl.ok("btnOK", "OK"), fl.cancel("btnCancel", "キャンセル"))]
    pl, cw, ch = fl.compute_layout(rows)
    assert min(l for e, l, t, w, h in pl) >= fl.STYLE['pad']
    assert max(l + w for e, l, t, w, h in pl) <= fl.STYLE['pad'] + cw + 0.01


def test_frame_autoname_deterministic():
    # hash() 乱数化で再buildのたび frame 名が変わらないこと
    a = fl.frame("設定", fl.row(fl.lbl("A"), fl.txt("t1")))
    b = fl.frame("設定", fl.row(fl.lbl("B"), fl.txt("t2")))
    c = fl.frame("別枠", fl.row(fl.lbl("C"), fl.txt("t3")))
    assert a['name'] == b['name']
    assert a['name'] != c['name']


def test_multipage_frame_children_present_in_layout():
    # page 内 frame の子が配置計画に含まれること（build 消失バグの回帰）
    rows = [fl.multipage("mp",
                         fl.page("基本", fl.frame("G", fl.row(fl.lbl("N"), fl.txt("t")))),
                         fl.page("詳細", fl.row(fl.lbl("M"), fl.txt("u"))))]
    pl, cw, ch = fl.compute_layout(rows)
    mp_e = next(e for e, *_ in pl if e['kind'] == 'multipage')
    frames = [ce for ce, *_ in mp_e['pages_layout'][0]['children'] if ce['kind'] == 'frame']
    assert frames and len(frames[0]['children']) == 2


# ================================================================
# vba_manager: Remove+Import の名前衝突ガード（2026-07-11 shu005→shu0051 事故の回帰）
# フェイク VBE で「Remove 遅延完了中の Import は連番付き別名になる」挙動を模す。
# 実 Excel での正常系は E2E（実機テスト）で担保。
# ================================================================

class _FakeComp:
    def __init__(self, name):
        self.Name = name


class _FakeVBComponents:
    """VBE の Import 挙動を模す: 同名が既に居ると連番付き別名で取り込まれる。

    ghost_clears_after_iters: 列挙がその回数を超えたらゴースト（遅延 Remove 中の
    旧モジュール）を消す。None なら永遠に残る（回復不能ケース）。
    """
    def __init__(self, names, base_name, ghost=None, ghost_clears_after_iters=None):
        self.comps = [_FakeComp(n) for n in names]
        self._base = base_name
        self._ghost = ghost
        self._clear_after = ghost_clears_after_iters
        self._iters = 0

    def __iter__(self):
        self._iters += 1
        if (self._ghost is not None and self._clear_after is not None
                and self._iters > self._clear_after):
            self.comps = [c for c in self.comps if c.Name != self._ghost]
            self._ghost = None
        return iter(list(self.comps))

    def Import(self, path):
        name = self._base
        if any(c.Name.lower() == name.lower() for c in self.comps):
            name = name + "1"          # VBE の連番リネーム
        c = _FakeComp(name)
        self.comps.append(c)
        return c


class _FakeWB:
    def __init__(self, components):
        self.VBProject = type("VBP", (), {"VBComponents": components})()


def test_import_verified_normal_returns_expected_name():
    comps = _FakeVBComponents(["OtherMod"], "TestMod")
    wb = _FakeWB(comps)
    got = vm._import_module_verified(wb, "x.bas", "TestMod",
                                     ghost_timeout=0.05, rename_timeout=0.05, settle=0)
    assert got.Name == "TestMod"
    assert [c.Name for c in comps.comps] == ["OtherMod", "TestMod"]


def test_import_verified_waits_for_ghost_and_avoids_collision():
    # 遅延 Remove 中のゴーストが待機中に消える → 衝突せず期待名で取り込まれる
    comps = _FakeVBComponents(["TestMod"], "TestMod",
                              ghost="TestMod", ghost_clears_after_iters=2)
    wb = _FakeWB(comps)
    got = vm._import_module_verified(wb, "x.bas", "TestMod",
                                     ghost_timeout=3.0, rename_timeout=0.05, settle=0)
    assert got.Name == "TestMod"
    assert [c.Name for c in comps.comps] == ["TestMod"]


def test_import_verified_collision_recovers_by_rename():
    # ゴーストが Import 後まで残る → TestMod1 で衝突 → 消滅を待って改名回復
    comps = _FakeVBComponents(["TestMod"], "TestMod",
                              ghost="TestMod", ghost_clears_after_iters=2)
    wb = _FakeWB(comps)
    got = vm._import_module_verified(wb, "x.bas", "TestMod",
                                     ghost_timeout=0.01, rename_timeout=3.0, settle=0)
    assert got.Name == "TestMod"
    assert [c.Name for c in comps.comps] == ["TestMod"]


def test_import_verified_unrecoverable_raises_not_silent_success():
    # ゴーストが消えない → 黙って成功にせず ModuleNameCollisionError
    comps = _FakeVBComponents(["TestMod"], "TestMod", ghost="TestMod")
    wb = _FakeWB(comps)
    try:
        vm._import_module_verified(wb, "x.bas", "TestMod",
                                   ghost_timeout=0.01, rename_timeout=0.05, settle=0)
        assert False, "ModuleNameCollisionError が飛ぶべき"
    except vm.ModuleNameCollisionError as ex:
        assert ex.expected_name == "TestMod"
        assert ex.actual_name == "TestMod1"
    # 衝突した別名側にコードが残っている（バックアップ再Importさせないための前提）
    assert any(c.Name == "TestMod1" for c in comps.comps)


def test_trailing_spacer_counts_fully():
    # 末尾 spacer が gap_y ぶん目減りしないこと（中間の spacer と同じ意味論）
    base = [fl.row(fl.lbl("X"), fl.txt("tX"))]
    _, _, h_a = fl.compute_layout(base)
    _, _, h_b = fl.compute_layout(base + [fl.spacer(24)])
    assert abs((h_b - h_a) - (fl.STYLE['gap_y'] + 24)) < 0.01
