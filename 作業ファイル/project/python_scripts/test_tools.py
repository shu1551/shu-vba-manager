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


def test_backup_reimport_only_when_import_actually_failed():
    """Import 成功後の Save 失敗で、バックアップを再 Import してはいけない（回帰）。

    2026-07-14 発見の実害: Remove+Import 系3経路の except が removed フラグしか
    見ておらず、Import が成功した後の wb.Save() が失敗すると「モジュールが消えた」と
    誤認してバックアップを重ね Import していた。期待名のモジュールは既に正しく
    存在するので _wait_component_gone は空振りし（15秒）、VB_Name 衝突で連番別名として
    取り込まれ、改名待ちも空振りして（20秒）例外になる——つまりツール自身が
    「旧コード入りの連番モジュール」を生み、35秒待たせたうえで失敗していた。

    復旧（バックアップの再 Import）は Import 自体が失敗したときだけ通ること。
    構文木で「復旧呼び出しを囲む if の条件に imported が出てくるか」を機械照合する。
    """
    import ast as _ast
    base = os.path.dirname(os.path.abspath(__file__))
    targets = ['vbam_vba.py', 'optimize_vba_modules.py']

    def _func_name(node):
        f = node.func
        return getattr(f, 'id', None) or getattr(f, 'attr', None) or ''

    total = 0
    for fname in targets:
        with open(os.path.join(base, fname), encoding='utf-8') as f:
            tree = _ast.parse(f.read())
        parents = {}
        for parent in _ast.walk(tree):
            for child in _ast.iter_child_nodes(parent):
                parents[child] = parent

        recoveries = [
            n for n in _ast.walk(tree)
            if isinstance(n, _ast.Call)
            and _func_name(n) == '_import_module_verified'
            and len(n.args) >= 2
            and getattr(n.args[1], 'id', '') == 'module_backup'
        ]
        for call in recoveries:
            names = set()
            node = call
            while node in parents:
                node = parents[node]
                if isinstance(node, _ast.If):
                    names |= {n.id for n in _ast.walk(node.test)
                              if isinstance(n, _ast.Name)}
            assert 'imported' in names, (
                f"{fname}: バックアップの再 Import が imported フラグで守られていない。\n"
                "  Import 成功後の Save 失敗で、ツール自身が連番モジュール"
                "（旧コード入り）を作る経路が復活している。")
        total += len(recoveries)

    # replace-procedure(Attribute経路) / replace-module / reorder / optimize の4経路
    assert total == 4, f"復旧経路の数が想定と違う（台帳の更新が必要）: {total}"


# ================================================================
# ダイアログ自動解除の報告（マクロ発火中のMsgBoxを黙って握りつぶさない）
# 2026-07-11: write-range→Worksheet_Change→MsgBox で無言ハングした実害の対策。
# COM/win32 部は実機E2Eで担保、ここは「検出時に報告文を出す」純粋部分の回帰。
# ================================================================

class _FakeWatcher:
    def __init__(self, count, last=""):
        self.count = count
        self.last_text = last


def test_dialog_note_empty_when_no_dialog():
    assert vm._dialog_watcher_note(_FakeWatcher(0), None) == ""
    assert vm._dialog_watcher_note(None, None) == ""


def test_dialog_note_safe_mode_reports_body():
    note = vm._dialog_watcher_note(_FakeWatcher(1, "B1が変わりました"), None)
    assert "1件" in note
    assert "安全側" in note
    assert "B1が変わりました" in note


def test_dialog_note_explicit_mode_says_specified_button():
    note = vm._dialog_watcher_note(_FakeWatcher(2, "確認"), "ok")
    assert "2件" in note
    assert "指定ボタン" in note


# ================================================================
# _collect_shapes: OnAction のブック修飾を落とさない（wiring の外部ブック判定の素）
# ================================================================

class _FakeShape:
    def __init__(self, name, onaction):
        self.Name = name
        self.Type = 1
        self.OnAction = onaction
        self.Left = 1.0
        self.Top = 2.0


def test_collect_shapes_keeps_onaction_book_qualifier():
    out = []
    vm._collect_shapes([_FakeShape("B1", "'秀 テスト.xlam'!Macro1"),
                        _FakeShape("B2", "Macro2"),
                        _FakeShape("B3", "ファイル一覧.xlsm!全検索開始")], out)
    assert out[0]["onaction"] == "Macro1"
    assert out[0]["onaction_book"] == "秀 テスト.xlam"
    assert out[1]["onaction"] == "Macro2"
    assert "onaction_book" not in out[1]
    assert out[2]["onaction"] == "全検索開始"
    assert out[2]["onaction_book"] == "ファイル一覧.xlsm"


# ================================================================
# snapshot-diff: 結合未走査の snapshot に疑似差分（結合追加/解除）を出さない
# ================================================================

def _write_snap(path, sheets):
    import json
    with open(path, 'w', encoding='utf-8') as f:
        json.dump({"success": True, "book": "T.xlsm", "sheets": sheets},
                  f, ensure_ascii=False)


def _diff_args(old_path, new_path):
    import argparse
    return argparse.Namespace(posargs=[str(old_path), str(new_path)], max_opt=None)


def test_snapshot_diff_merged_unscanned_suppresses_pseudo_diff(tmp_path, capsys):
    # 旧=結合未走査 / 新=結合あり。セル差分でシートは表示されるが、
    # 「結合追加」の疑似差分は出さず、比較していない旨を注記する
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "1行 x 1列",
                           "cells": [{"r": 1, "c": {"A": "x"}}],
                           "merged_skipped_cells": 99999}})
    _write_snap(p2, {"S": {"dims": "1行 x 1列",
                           "cells": [{"r": 1, "c": {"A": "y"}}],
                           "merged": ["$A$5:$B$5"]}})
    assert vm.cmd_snapshot_diff(_diff_args(p1, p2)) is True
    out = capsys.readouterr().out
    assert "結合追加" not in out
    assert "結合の差分は比較していない" in out
    assert "セル変更: 1件" in out


def test_snapshot_diff_merged_unscanned_note_survives_hidden_sheet(tmp_path, capsys):
    # 結合以外に差分がなくシート自体が非表示でも、「比較できていない」事実は落とさない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "1行 x 1列", "cells": [],
                           "merged_skipped_cells": 99999}})
    _write_snap(p2, {"S": {"dims": "1行 x 1列", "cells": [],
                           "merged": ["$A$1:$B$1"]}})
    assert vm.cmd_snapshot_diff(_diff_args(p1, p2)) is True
    out = capsys.readouterr().out
    assert "結合の差分を比較できていません" in out
    assert "S" in out.split("比較できていません:")[-1]
    assert "差分なし" in out


def test_trailing_spacer_counts_fully():
    # 末尾 spacer が gap_y ぶん目減りしないこと（中間の spacer と同じ意味論）
    base = [fl.row(fl.lbl("X"), fl.txt("tX"))]
    _, _, h_a = fl.compute_layout(base)
    _, _, h_b = fl.compute_layout(base + [fl.spacer(24)])
    assert abs((h_b - h_a) - (fl.STYLE['gap_y'] + 24)) < 0.01


# ================================================================
# vba_manager: VBA 識別子ガード（先頭 _ の Sub 注入事故 `_tmp検証` の回帰）
# ================================================================

def test_check_vba_identifier_rejects_leading_underscore():
    assert vm.check_vba_identifier("_tmp検証") is not None


def test_check_vba_identifier_rejects_leading_digit():
    assert vm.check_vba_identifier("1テスト") is not None


def test_check_vba_identifier_rejects_symbols():
    assert vm.check_vba_identifier("foo-bar") is not None


def test_check_vba_identifier_accepts_normal_names():
    for name in ("tmp検証", "テスト検証", "Btn_Click", "UserForm_Initialize", "A1"):
        assert vm.check_vba_identifier(name) is None, name


def test_find_invalid_procedure_names_hits_declaration():
    code = "Sub _tmp検証()\nEnd Sub\n"
    hits = vm._find_invalid_procedure_names(code)
    assert len(hits) == 1
    assert hits[0][1] == "_tmp検証"


def test_find_invalid_procedure_names_ignores_comments_and_events():
    code = ("' Sub _コメントは対象外()\n"
            "Private Sub CommandButton1_Click()\n"
            "End Sub\n"
            "Property Get 値()\n"
            "End Property\n")
    assert vm._find_invalid_procedure_names(code) == []


def test_validate_vba_code_rejects_underscore_name():
    assert vm.validate_vba_code("Sub _tmp検証()\nEnd Sub\n") is False


def test_validate_vba_code_accepts_valid_japanese_name():
    assert vm.validate_vba_code("Sub tmp検証()\nEnd Sub\n") is True


def test_check_bas_rejects_invalid_identifier(tmp_path):
    p = tmp_path / "m.bas"
    p.write_bytes('Attribute VB_Name = "M"\r\nSub _tmp検証()\r\nEnd Sub\r\n'.encode('cp932'))
    assert vm._check_bas_one(str(p)) is False


# ================================================================
# 注入経路の台帳（機械検査）
# ================================================================

def test_injection_route_ledger():
    """VBA へコードを入れる注入プリミティブの台帳。

    2026-07-12 の識別子ガード配線で一番時間を食ったのは「注入経路の洗い出し」
    だった。この台帳が現物と一致する限り、次回の穴塞ぎで経路の再調査は不要。
    新しい注入点を足す手順: ①ガードを配線する（または安全な理由を確認する）
    ②下の EXPECTED に理由コメントつきで登録する。
    未登録の注入点が現れたらこのテストが落ちる＝ガード無しの新経路の検知器。
    """
    import re as _re
    base = os.path.dirname(os.path.abspath(__file__))
    PRIM = _re.compile(r'\.AddFromString\(|\.InsertLines\(|VBComponents\.Import\(')
    DEF = _re.compile(r'^\s*def\s+(\w+)')
    EXPECTED = {
        # (ファイル, 関数): ガードの所在／安全な理由
        ('form_builder.py', 'inject_vba'),          # 注入前に識別子検査（既存コード削除より前）
        ('form_inspect.py', 'render_form_png'),     # 機械固定名 tmpFormInspect* のみ＝安全
        ('form_layout.py', 'build_form'),           # 起動マクロ名を check_vba_identifier で検査
        ('form_tool.py', 'cmd_copy_form'),          # 新フォーム名を check_vba_identifier で検査
        ('live_sync_vba.py', 'update_module'),      # 既存コード往復＋固定名マクロ追記のみ＝新規名の流入なし
                                                    # （2026-07-14: 失敗時に旧コードを書き戻すため関数へ切り出し）
        # 2026-07-14: optimize_vba_modules.apply_module は素の VBComponents.Import をやめ、
        # vbam_core._import_module_verified（実名検証つき）に委譲したので注入経路ではなくなった
        # 2026-07-12 分割: vba_manager.py の実装は vbam_core/vbam_vba 等へ移動（入口は不変）
        ('vbam_core.py', '_import_module_verified'),  # 取込の中央関数（名前衝突ガード）。内容の識別子検査は呼び元
        ('vbam_vba.py', 'cmd_replace_procedure'),     # validate_vba_code で識別子検査
        ('vbam_vba.py', 'cmd_add_procedure'),         # validate_vba_code で識別子検査
        ('vbam_vba.py', 'cmd_test'),                  # 機械固定名 VMT_n ハーネス＝安全
        ('test_e2e_com.py', 'book'),                  # E2E の使い捨てブック生成。固定リテラルの
                                                      # テストコードのみ＝外から名前が流入しない
    }
    found = set()
    for fname in sorted(os.listdir(base)):
        if not fname.endswith('.py') or fname == os.path.basename(__file__):
            continue
        cur = '(module)'
        with open(os.path.join(base, fname), encoding='utf-8', errors='replace') as f:
            for line in f:
                m = DEF.match(line)
                if m:
                    cur = m.group(1)
                if PRIM.search(line):
                    found.add((fname, cur))
    new_routes = found - EXPECTED
    # 台帳のファイルが現物に無いのは「経路が消えた」ではなく「ファイルごと無い」
    # （例: live_sync_vba.py は 2026-07-16 にローカル専用へ退役＝公開リポに入らない。
    #   クローン先ではファイル自体が無いのが正常）。ファイルは在るのに注入が
    # 消えた場合だけ「台帳の掃除」として検知する
    gone = {e for e in (EXPECTED - found)
            if os.path.exists(os.path.join(base, e[0]))}
    assert not new_routes, (
        f"台帳に無い注入経路: {sorted(new_routes)}\n"
        "  ガード（check_vba_identifier / _find_invalid_procedure_names）を配線してから"
        "台帳に理由コメントつきで登録すること。")
    assert not gone, f"台帳にあるのに現物に無い注入経路（台帳の掃除が必要）: {sorted(gone)}"


def test_copy_form_rejects_leading_underscore():
    # 旧regexは先頭 _ を素通しした（\w−数字＝英字＋_）。COM接続前に止まることの回帰
    import argparse
    import pytest
    import form_tool
    with pytest.raises(SystemExit):
        form_tool.cmd_copy_form(argparse.Namespace(form="F_X", new="_F_X2"))


# ================================================================
# 2026-07-15 点検: 「1行完結 Sub」の残存経路（delete-procedure / call-graph）
# ================================================================

class _FakeCodeModule:
    """ProcStartLine/ProcCountLines の領域に次の宣言行が食い込む状況を再現する。"""
    def __init__(self, lines):
        self._lines = lines

    def Lines(self, start, count):
        return '\r\n'.join(self._lines[start - 1:start - 1 + count])


def test_narrow_proc_range_excludes_next_declaration():
    # 1行完結 Sub の直後のプロシージャ。ProcCountLines が次の宣言行まで含んでしまう
    # （count=3 が「Sub 次の処理()」まで食い込んだ状態）
    cm = _FakeCodeModule([
        'Sub 一行版(): Call Main: End Sub',   # 1行目
        '',                                    # 2行目
        'Sub 次の処理()',                      # 3行目 ← 食い込み
        '    MsgBox 1',
        'End Sub',
    ])
    start, count = vm._narrow_proc_range(cm, 1, 3)
    assert (start, count) == (1, 1), "次の宣言行を削除範囲に含めてはいけない"
    # 絞った範囲は1行完結 Sub 自身のみ
    assert cm.Lines(start, count) == 'Sub 一行版(): Call Main: End Sub'


def test_narrow_proc_range_keeps_normal_proc_intact():
    cm = _FakeCodeModule([
        'Sub 普通()',
        '    MsgBox 1',
        'End Sub',
        '',
    ])
    start, count = vm._narrow_proc_range(cm, 1, 4)
    assert (start, count) == (1, 3)


def test_inline_body_after_decl():
    f = vm._inline_body_after_decl
    # 1行完結 Sub → 本体を返す
    raw = 'Sub X(): Call Main: End Sub'
    assert f(raw, raw.index('X') + 1).strip() == 'Call Main: End Sub'
    # 普通の宣言行 → 空
    raw2 = 'Sub Y(a As String)'
    assert f(raw2, raw2.index('Y') + 1) == ''
    # 引数の既定値の文字列に ':' や ')' があっても誤らない
    raw3 = 'Sub Z(Optional s As String = "a:b)"): Call Main: End Sub'
    assert f(raw3, raw3.index('Z') + 1).strip() == 'Call Main: End Sub'
    # 戻り型指定つき Function の1行完結
    raw4 = 'Function F() As String: F = "x": End Function'
    assert 'F = ' in f(raw4, raw4.index('F') + 1)


def _inv(code, procs):
    return {'modules': [{'name': 'Mod1', 'type': 1, 'code': code,
                         'procs': [{'name': p} for p in procs]}],
            'onaction': []}


def test_analyze_calls_sees_body_of_one_line_sub():
    # 「Sub ボタン18_Click(): Call 印刷実行: End Sub」の Call を見落とすと、
    # 印刷実行が「どこからも呼ばれていない」に誤って載る
    code = '\r\n'.join([
        'Sub ボタン18_Click(): Call 印刷実行: End Sub',
        '',
        'Sub 印刷実行()',
        '    MsgBox 1',
        'End Sub',
    ])
    res = vm._analyze_calls(_inv(code, ['ボタン18_Click', '印刷実行']))
    assert ('Mod1', 'ボタン18_Click') in res['edges'], "1行完結Subの Call が計上されていない"
    assert '印刷実行' in res['edges'][('Mod1', 'ボタン18_Click')]
    assert '印刷実行' not in res['orphans'], "呼ばれているのに孤立扱いになっている"


def test_analyze_calls_detects_typo_in_one_line_sub():
    # 1行完結 Sub の中の誤記（印刷実効）が未解決Callとして検出されること
    code = '\r\n'.join([
        'Sub ボタン18_Click(): Call 印刷実効: End Sub',
        '',
        'Sub 印刷実行()',
        'End Sub',
    ])
    res = vm._analyze_calls(_inv(code, ['ボタン18_Click', '印刷実行']))
    names = [u[2] for u in res['unresolved']]
    assert '印刷実効' in names, "1行完結Subの中の誤記が検出されていない"


def test_analyze_calls_one_line_sub_does_not_capture_following_lines():
    # 1行完結 Sub は宣言行で閉じている。後続行の呼び出しをその Sub の名で
    # 計上し続けてはいけない（cur が閉じられているか）
    code = '\r\n'.join([
        'Sub 一行版(): Call A: End Sub',
        '',
        'Sub 別の処理()',
        '    Call B',
        'End Sub',
    ])
    res = vm._analyze_calls(_inv(code, ['一行版', '別の処理', 'A', 'B']))
    assert 'B' not in res['edges'].get(('Mod1', '一行版'), set()), \
        "1行完結Subの後続行が、その Sub の呼び出しとして計上されている"
    assert 'B' in res['edges'][('Mod1', '別の処理')]


# ================================================================
# 2026-07-15 点検: snapshot-diff が「読めなかった」を「消えた」と誤報する件
# （ヘルパーは既存の _write_snap(path, sheets) / _diff_args(old, new) を使う）
# ================================================================

def test_snapshot_diff_does_not_report_unreadable_shapes_as_deleted(tmp_path, capsys):
    # 旧: 図形3つを正常に読めた / 新: 図形の列挙に失敗（shapes_error だけ）
    # → 図形は1つも消えていない。「図形削除: 3件」と言ってはいけない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "10行 x 3列", "cells": [],
                           "shapes": [{"name": "ボタン1"}, {"name": "ボタン2"},
                                      {"name": "ボタン3"}]}})
    _write_snap(p2, {"S": {"dims": "10行 x 3列", "cells": [],
                           "shapes_error": "COM エラー: 図形を列挙できません"}})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "図形削除" not in out, "読めなかっただけの図形を「削除」と誤報している"
    assert "比較していません" in out or "比較できていません" in out


def test_snapshot_diff_never_claims_match_when_unreadable(tmp_path, capsys):
    # 両側とも図形が読めない → 差分ゼロに見えるが「一致」と断言してはいけない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    info = {"dims": "10行 x 3列", "cells": [], "shapes_error": "COM エラー"}
    _write_snap(p1, {"S": dict(info)})
    _write_snap(p2, {"S": dict(info)})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "一致" not in out, "読めていないのに「一致」と断言している"
    assert "比較できていません" in out


def test_snapshot_diff_does_not_report_unreadable_cells_as_deleted(tmp_path, capsys):
    # 旧: セルを読めた / 新: セルの読み取りに失敗 → 「セル削除」と言ってはいけない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "2行 x 1列",
                           "cells": [{"r": 1, "c": {"A": "あ"}},
                                     {"r": 2, "c": {"A": "い"}}]}})
    _write_snap(p2, {"S": {"dims": "2行 x 1列", "cells": [],
                           "cells_error": "COM エラー: 値を読めません"}})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "セル削除" not in out, "読めなかっただけのセルを「削除」と誤報している"


def test_snapshot_diff_detects_shape_resize(tmp_path, capsys):
    # ボタンを 1x1pt に潰す変更を「差分なし」と言ってはいけない（図形のサイズの目）
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": {"dims": "10行 x 3列", "cells": [],
                           "shapes": [{"name": "実行ボタン", "l": 10, "t": 10,
                                       "w": 96, "h": 24}]}})
    _write_snap(p2, {"S": {"dims": "10行 x 3列", "cells": [],
                           "shapes": [{"name": "実行ボタン", "l": 10, "t": 10,
                                       "w": 1, "h": 1}]}})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "図形変更" in out, "図形のサイズ変更を見落としている"
    assert "大きさ" in out


def test_snapshot_diff_match_message_does_not_overclaim(tmp_path, capsys):
    # 本当に一致しているときも、見ていない書式まで一致したかのように言わない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    info = {"dims": "1行 x 1列", "cells": [{"r": 1, "c": {"A": "x"}}]}
    _write_snap(p1, {"S": dict(info)})
    _write_snap(p2, {"S": dict(info)})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "差分なし" in out
    assert "書式" in out, "書式が比較対象外であることを明示していない"


# ================================================================
# 2026-07-15: 書式の目（四つ目の目）。クリーン化で書式が飛んでも
# 「差分なし（一致）」と言っていたのを塞いだ回帰テスト
# ================================================================

def _fmt_sheet(sheet=None, cols=None, rows=None):
    """format ブロックつきのシート情報を作る"""
    return {"dims": "11行 x 4列", "cells": [{"r": 1, "c": {"A": "商品"}}],
            "format": {"sheet": sheet or {}, "cols": cols or {}, "rows": rows or {}}}


def test_snapshot_diff_detects_wiped_formatting(tmp_path, capsys):
    # クリーン化マクロが書式だけを吹き飛ばした（値・結合・図形は無傷）。
    # 従来はこれを「差分なし（一致）」と報告していた＝ポスター.xlsm 破壊の再演
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"売上表": _fmt_sheet(
        sheet={"bold": None, "border": 1, "fill": 0, "numfmt": None},
        cols={"A": {"bold": None, "border": 1, "w": 18.0}},
        rows={"1": {"bold": True, "border": 1, "h": 24.0}})})
    _write_snap(p2, {"売上表": _fmt_sheet(
        sheet={"bold": False, "border": -4142, "fill": 16777215, "numfmt": "G/標準"},
        cols={"A": {"bold": False, "border": -4142, "w": 8.42}},
        rows={"1": {"bold": False, "border": -4142, "h": 18.0}})})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "書式変更" in out, "書式が吹き飛んだのに検出していない"
    assert "罫線" in out, "罫線の消失を検出していない"
    assert "一致" not in out
    # シート全体の一撃を列・行ぶん繰り返して騒がしくしない（読めない検分は役に立たない）
    assert "列幅" in out, "局所的な変化（列幅）は出すこと"


def test_snapshot_diff_format_noise_is_suppressed(tmp_path, capsys):
    # シート全体で起きた変化を、各列でも繰り返さない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    cols_old = {c: {"border": 1, "w": 8.38} for c in "ABCD"}
    cols_new = {c: {"border": -4142, "w": 8.38} for c in "ABCD"}
    _write_snap(p1, {"S": _fmt_sheet(sheet={"border": 1}, cols=cols_old)})
    _write_snap(p2, {"S": _fmt_sheet(sheet={"border": -4142}, cols=cols_new)})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "シート全体: 罫線" in out
    # A列〜D列で同じ内容を繰り返さない
    assert "A列: 罫線" not in out, "シート全体の変化を列ぶん繰り返している（騒がしい）"


def test_snapshot_diff_format_local_change_survives(tmp_path, capsys):
    # 局所的な変化（その列だけ違う）は、抑制に巻き込まれず出ること
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": _fmt_sheet(sheet={"border": 1},
                                     cols={"A": {"border": 1}, "B": {"border": 1}})})
    _write_snap(p2, {"S": _fmt_sheet(sheet={"border": -4142},
                                     cols={"A": {"border": -4142}, "B": {"border": 9}})})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "B列: 罫線" in out, "列だけの局所的な変化が消えている"
    assert "A列: 罫線" not in out


def test_snapshot_diff_does_not_report_unreadable_format(tmp_path, capsys):
    # 書式を読めなかっただけの snapshot を「書式が変わった」と誤報しない
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": _fmt_sheet(sheet={"bold": True})})
    s2 = {"dims": "11行 x 4列", "cells": [{"r": 1, "c": {"A": "商品"}}],
          "format_error": "COM エラー: 書式を読めません"}
    _write_snap(p2, {"S": s2})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "書式変更" not in out, "読めなかっただけの書式を「変更」と誤報している"
    assert ("比較していません" in out or "比較できていません" in out)
    assert "一致" not in out, "書式を読めていないのに「一致」と言っている"


def test_snapshot_diff_format_missing_on_one_side(tmp_path, capsys):
    # 片方が --no-format で採られている → 書式は比較できないと明示する
    p1, p2 = tmp_path / "old.json", tmp_path / "new.json"
    _write_snap(p1, {"S": _fmt_sheet(sheet={"bold": True})})
    _write_snap(p2, {"S": {"dims": "11行 x 4列",
                           "cells": [{"r": 1, "c": {"A": "商品"}}]}})
    vm.cmd_snapshot_diff(_diff_args(p1, p2))
    out = capsys.readouterr().out
    assert "書式変更" not in out
    assert ("比較していません" in out or "比較できていません" in out)


# ================================================================
# open / close / rehearse（2026-07-15 追加）: COM に触る前の門番
# ================================================================

import argparse as _ap


def test_open_requires_path(capsys):
    assert vm.cmd_open(_ap.Namespace(posargs=[])) is False
    assert "使い方" in capsys.readouterr().out


def test_open_rejects_multiple_paths(capsys):
    assert vm.cmd_open(_ap.Namespace(posargs=["a.xlsm", "b.xlsm"])) is False
    assert "1つだけ" in capsys.readouterr().out


def test_open_rejects_missing_file(capsys):
    assert vm.cmd_open(_ap.Namespace(posargs=["__vbam_unit_no_such__.xlsm"])) is False
    assert "見つかりません" in capsys.readouterr().out


def test_close_requires_name(capsys):
    ns = _ap.Namespace(posargs=[], save_flag=True, no_save_flag=False, yes=True)
    assert vm.cmd_close(ns) is False
    assert "使い方" in capsys.readouterr().out


def test_close_requires_save_policy(capsys):
    # 保存方針の明示は鎧の一部。無指定も両指定も門前払いする
    ns = _ap.Namespace(posargs=["a.xlsm"], save_flag=False, no_save_flag=False, yes=True)
    assert vm.cmd_close(ns) is False
    out = capsys.readouterr().out
    assert "--save" in out and "--no-save" in out

    ns = _ap.Namespace(posargs=["a.xlsm"], save_flag=True, no_save_flag=True, yes=True)
    assert vm.cmd_close(ns) is False


def test_close_not_open_book_fails(capsys):
    # 開いていないブック名は「開いていません」で止まる（保存方針が正しくても）
    ns = _ap.Namespace(posargs=["__vbam_unit_no_such__.xlsm"],
                       save_flag=False, no_save_flag=True, yes=True)
    assert vm.cmd_close(ns) is False
    assert "開いていません" in capsys.readouterr().out


def test_rehearse_requires_macro_name(capsys):
    assert vm.cmd_rehearse(_ap.Namespace(posargs=[])) is False
    assert "使い方" in capsys.readouterr().out


def test_new_commands_are_wired():
    # コマンド表と argparse の両方に配線されていること（表だけ・パーサだけの片肺を防ぐ）
    table = vm._command_table()
    for name in ("open", "close", "rehearse", "予行演習"):
        assert name in table
    parser = vm.build_parser()
    ns, unknown = parser.parse_known_args(
        ["close", "a.xlsm", "--no-save", "-y"])
    assert ns.command == "close" and ns.no_save_flag and ns.yes and not unknown
    ns, unknown = parser.parse_known_args(
        ["rehearse", "マクロA", "--addins", "--discard", "--max", "5"])
    assert ns.command == "rehearse" and ns.addins and ns.discard
    assert ns.max_opt == "5" and not unknown
    ns, unknown = parser.parse_known_args(["open", "a.xlsm"])
    assert ns.command == "open" and not unknown
