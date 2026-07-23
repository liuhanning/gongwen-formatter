"""
基础 pytest 测试 —— 验证 Python 辅助脚本的核心函数。

运行方式:
    .venv\\Scripts\\python.exe -m pytest tests/ -v
"""

import importlib.util
import sys
import types
from pathlib import Path
from unittest.mock import patch, mock_open

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent


# ---------------------------------------------------------------------------
# 辅助：安全导入含顶层副作用代码的脚本
# ---------------------------------------------------------------------------

def _import_module_safe(module_name: str, file_path: Path):
    """导入脚本模块，阻止其顶层文件 I/O 执行。"""
    fake_file = mock_open(read_data="")
    with patch("builtins.open", fake_file):
        spec = importlib.util.spec_from_file_location(module_name, str(file_path))
        mod = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = mod
        spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="session")
def fix_encoding_mod():
    return _import_module_safe("fix_encoding", PROJECT_ROOT / "fix_encoding.py")


@pytest.fixture(scope="session")
def repair_vba_mod():
    return _import_module_safe("repair_vba", PROJECT_ROOT / "repair_vba.py")


@pytest.fixture(scope="session")
def refactor_msgbox_mod():
    return _import_module_safe("refactor_msgbox", PROJECT_ROOT / "refactor_msgbox.py")


# ===========================================================================
# fix_encoding.py 测试
# ===========================================================================

class TestToChrw:
    """测试 to_chrw：将字符串转换为 VBA ChrW 表达式。"""

    def test_pure_ascii(self, fix_encoding_mod):
        result = fix_encoding_mod.to_chrw("Hello")
        assert result == '"Hello"'

    def test_pure_chinese(self, fix_encoding_mod):
        # "公文" -> ChrW(&H516C) & ChrW(&H6587)
        result = fix_encoding_mod.to_chrw("公文")
        assert "ChrW(&H516C)" in result
        assert "ChrW(&H6587)" in result

    def test_mixed_ascii_and_chinese(self, fix_encoding_mod):
        result = fix_encoding_mod.to_chrw("A公B")
        # 应包含 ASCII 片段和 ChrW 片段
        assert '"A"' in result
        assert "ChrW(&H516C)" in result
        assert '"B"' in result
        assert " & " in result

    def test_quote_escaping(self, fix_encoding_mod):
        # 双引号在 VBA 字符串中需要转义为 ""
        result = fix_encoding_mod.to_chrw('say "hi"')
        assert '""' in result

    def test_empty_string(self, fix_encoding_mod):
        result = fix_encoding_mod.to_chrw("")
        assert result == ""


class TestProcessLineContent:
    """测试 process_line_content：替换行内非 ASCII 字符串字面量为 ChrW。"""

    def test_ascii_only_unchanged(self, fix_encoding_mod):
        line = 'x = "Hello World"'
        result = fix_encoding_mod.process_line_content(line)
        assert result == line

    def test_chinese_in_string_replaced(self, fix_encoding_mod):
        line = 'MsgBox "你好"'
        result = fix_encoding_mod.process_line_content(line)
        assert "ChrW(" in result
        assert "你好" not in result

    def test_multiple_strings(self, fix_encoding_mod):
        line = 'x = "abc" & "中文" & "def"'
        result = fix_encoding_mod.process_line_content(line)
        # ASCII 部分保留
        assert '"abc"' in result
        assert '"def"' in result
        # 中文部分被替换
        assert "ChrW(" in result


# ===========================================================================
# repair_vba.py 测试
# ===========================================================================

class TestRepairContent:
    """测试 repair_content：修复被拆散的 MsgBox 代码块。"""

    def test_non_msgbox_code_unchanged(self, repair_vba_mod):
        code = "Sub Foo()\n    x = 1\nEnd Sub"
        result = repair_vba_mod.repair_content(code)
        # 输出使用 \r\n
        assert "x = 1" in result
        assert "Sub Foo()" in result

    def test_msgbox_block_reassembled(self, repair_vba_mod):
        code = (
            "    Dim msg As String\n"
            '    msg = ChrW(&H4F60)\n'
            '    msg = msg & ChrW(&H597D)\n'
            "    MsgBox msg, vbOKOnly\n"
        )
        result = repair_vba_mod.repair_content(code)
        # 修复后应合并 msg 赋值行
        assert "Dim msg As String" in result
        assert "MsgBox msg" in result

    def test_fixes_broken_chrw(self, repair_vba_mod):
        # 模拟 ( & H 被错误拆开的情况
        code = (
            "    Dim msg As String\n"
            '    msg = ChrW( & H4F60)\n'
            "    MsgBox msg, vbOKOnly\n"
        )
        result = repair_vba_mod.repair_content(code)
        # 修复后 (&H 应合并
        assert "(&H4F60)" in result


# ===========================================================================
# refactor_msgbox.py 测试
# ===========================================================================

class TestParseVbMsgbox:
    """测试 parse_vb_msgbox：解析 MsgBox 语句的内容和参数。"""

    def test_simple_msgbox(self, refactor_msgbox_mod):
        stmt = 'MsgBox "Hello", vbOKOnly, "Title"'
        content, rest = refactor_msgbox_mod.parse_vb_msgbox(stmt)
        assert content == '"Hello"'
        assert "vbOKOnly" in rest

    def test_msgbox_with_chrw(self, refactor_msgbox_mod):
        stmt = 'MsgBox ChrW(&H4F60) & ChrW(&H597D), vbYesNo, "Test"'
        content, rest = refactor_msgbox_mod.parse_vb_msgbox(stmt)
        assert "ChrW(&H4F60)" in content
        assert "vbYesNo" in rest

    def test_no_comma_returns_none(self, refactor_msgbox_mod):
        stmt = 'MsgBox "JustContent"'
        content, rest = refactor_msgbox_mod.parse_vb_msgbox(stmt)
        assert content is None
        assert rest is None
