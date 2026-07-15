"""pytest 共通設定。

テストは2本立て:
  純ロジック（既定）  py -m pytest -q
  実機E2E（明示）     py -m pytest --run-e2e -q

E2E は本物の Excel / VBE を起こすため、既定では走らせない（オプトイン）。
CI や普段の確認は純ロジックだけで 1〜2 秒で終わる。
"""
import pytest

# sync_tools.py が退避した旧版コピー（_sync_backup/～/test_tools.py 等）を
# pytest が再帰収集すると、本物の test_tools.py と import 名が衝突して
# collection error になり、テストが1件も走らない（公開リポ側で実測）。
# 退避フォルダは収集対象から外す
collect_ignore_glob = ["_sync_backup*"]


def pytest_addoption(parser):
    parser.addoption(
        "--run-e2e", action="store_true", default=False,
        help="実機の Excel/VBE を使う E2E テストも実行する（既定はスキップ）")


def pytest_configure(config):
    config.addinivalue_line(
        "markers", "e2e: 実機の Excel/VBE を起こすテスト（--run-e2e で有効）")


def pytest_collection_modifyitems(config, items):
    if config.getoption("--run-e2e"):
        return
    skip_e2e = pytest.mark.skip(reason="実機E2E（走らせるには --run-e2e）")
    for item in items:
        if "e2e" in item.keywords:
            item.add_marker(skip_e2e)
