"""模板引擎测试：_subst / _eval_truthy / _find_block_end / _render_lines /
render_template / build_switch_config。

全部使用虚构数据（SW-1-1F / 192.0.2.x）。GUI/网络相关逻辑不在本套件范围内。
"""

import pytest

import nadt as N

SMALL_TPL = (
    "sysname {{hostname}}\n"
    "vlan batch {{vlan_batch}}\n"
    "interface Vlanif{{mgmt_vlan}}\n"
    " ip address {{mgmt_ip}} {{mask}}\n"
    " ip route-static 0.0.0.0 {{gateway}}\n"
)


# ────────────────────────── _subst ──────────────────────────

def test_subst_replaces_defined_var():
    assert N._subst("sysname {{hostname}}", {"hostname": "SW-1-1F"}) == "sysname SW-1-1F"


def test_subst_keeps_undefined_var_untouched():
    """未定义变量保留原样，绝不静默清空。"""
    assert N._subst("a {{nope}} b", {}) == "a {{nope}} b"


def test_subst_repeats_and_non_string_values():
    assert N._subst("{{a}}/{{b}}/{{a}}", {"a": 1, "b": 2}) == "1/2/1"


@pytest.mark.parametrize("line,expected", [
    ("{{hostname}}-{{mgmt_vlan}}", "SW-1-1F-11"),
    ("no vars here", "no vars here"),
    ("{{mgmt_ip}}", "192.0.2.11"),
])
def test_subst_cases(line, expected):
    assert N._subst(line, {"hostname": "SW-1-1F", "mgmt_vlan": "11",
                           "mgmt_ip": "192.0.2.11"}) == expected


# ────────────────────────── _eval_truthy ──────────────────────────

@pytest.mark.parametrize("expr,vars_map,expected", [
    ("x", {"x": "1"}, True),
    ("x", {"x": "  "}, False),
    ("x", {"x": ""}, False),
    ("x", {}, False),
    ("!x", {"x": ""}, True),
    ("!x", {}, True),
    ("!x", {"x": "1"}, False),
    ("a=1", {"a": "1"}, True),
    ("a=1", {"a": " 1 "}, True),
    ("a=1", {"a": "2"}, False),
    ("a=1", {}, False),
    ("!a=1", {"a": "2"}, True),
    ("hostname=SW-1-1F", {"hostname": "SW-1-1F"}, True),
    ("hostname=SW-1-2F", {"hostname": "SW-1-1F"}, False),
])
def test_eval_truthy(expr, vars_map, expected):
    assert N._eval_truthy(vars_map, expr) is expected


def test_eval_truthy_strips_expression():
    assert N._eval_truthy({"a": "1"}, "  a = 1  ") is True


# ────────────────────────── _find_block_end ──────────────────────────

def test_find_block_end_if_else():
    lines = ["{{#if a}}", "A", "{{#else}}", "B", "{{#endif}}", "C"]
    block, els, nxt = N._find_block_end(lines, 1)
    assert block == ["A"]
    assert els == ["B"]
    assert nxt == 5


def test_find_block_end_if_without_else():
    lines = ["{{#if a}}", "A", "B", "{{#endif}}", "C"]
    block, els, nxt = N._find_block_end(lines, 1)
    assert block == ["A", "B"]
    assert els is None
    assert nxt == 4


def test_find_block_end_for():
    lines = ["{{#for n from 1 to 2}}", "port {{n}}", "{{#endfor}}", "tail"]
    block, els, nxt = N._find_block_end(lines, 1)
    assert block == ["port {{n}}"]
    assert els is None
    assert nxt == 3


def test_find_block_end_nested_blocks():
    lines = ["{{#if a}}", "A", "{{#if b}}", "B", "{{#endif}}", "{{#endif}}", "tail"]
    block, els, nxt = N._find_block_end(lines, 1)
    assert block == ["A", "{{#if b}}", "B", "{{#endif}}"]
    assert nxt == 6


def test_find_block_end_ignores_inline_child():
    """同行内联的子块不参与深度计数。"""
    lines = ["{{#if a}}", "{{#if b}}x{{#endif}}", "{{#endif}}"]
    block, _, nxt = N._find_block_end(lines, 1)
    assert block == ["{{#if b}}x{{#endif}}"]
    assert nxt == 3


def test_find_block_end_missing_terminator_raises():
    with pytest.raises(ValueError, match="缺少块结束符"):
        N._find_block_end(["A", "B"], 0)


# ────────────────────────── _render_lines / render_template ──────────────────────────

def test_render_lines_direct():
    out = []
    N._render_lines(["a{{x}}", "b"], {"x": "1"}, out)
    assert out == ["a1", "b"]


def test_render_template_plain_substitution():
    assert N.render_template("sysname {{hostname}}\n", {"hostname": "SW-1-1F"}) == \
        "sysname SW-1-1F\n"


def test_render_template_undefined_var_kept():
    assert N.render_template("sysname {{hostname}}\n", {}) == "sysname {{hostname}}\n"


def test_render_template_newline_handling():
    assert N.render_template("{{a}}", {"a": "1"}) == "1"      # 无尾换行 → 不补
    assert N.render_template("", {}) == ""
    assert N.render_template("{{a}}\n", {"a": "1"}) == "1\n"  # 有尾换行 → 保留


def test_render_template_if_block_multiline():
    tpl = "line1\n{{#if a}}\n  inner {{hostname}}\n{{#endif}}\nline2\n"
    assert N.render_template(tpl, {"a": "1", "hostname": "SW-1-1F"}) == \
        "line1\n  inner SW-1-1F\nline2\n"
    assert N.render_template(tpl, {"a": "", "hostname": "SW-1-1F"}) == "line1\nline2\n"


def test_render_template_if_else_block_multiline():
    tpl = "{{#if a}}\nA\n{{#else}}\nB\n{{#endif}}\n"
    assert N.render_template(tpl, {"a": "1"}) == "A\n"
    assert N.render_template(tpl, {"a": ""}) == "B\n"


def test_render_template_if_variants():
    tpl = "{{#if a=1}}one{{#else}}other{{#endif}}\n"
    assert N.render_template(tpl, {"a": "1"}) == "one\n"
    assert N.render_template(tpl, {"a": "2"}) == "other\n"
    tpl_neg = "{{#if !a}}empty{{#else}}set{{#endif}}\n"
    assert N.render_template(tpl_neg, {"a": ""}) == "empty\n"
    assert N.render_template(tpl_neg, {"a": "1"}) == "set\n"


def test_render_template_inline_if_without_else_is_empty_when_false():
    assert N.render_template("{{#if a}}T{{#endif}}\n", {"a": ""}) == "\n"
    assert N.render_template("{{#if a}}T{{#endif}}\n", {"a": "1"}) == "T\n"


def test_render_template_inline_directive_allows_spaces():
    assert N.render_template("{{#if  a = 1 }}T{{#else}}F{{#endif}}\n", {"a": "1"}) == "T\n"


def test_render_template_for_literal_range():
    assert N.render_template("{{#for n from 1 to 3}}\nvlan {{n}}\n{{#endfor}}\n", {}) == \
        "vlan 1\nvlan 2\nvlan 3\n"


def test_render_template_for_variable_range():
    """for 范围写成 {{#for n from 1 to {{port_count}}}}。"""
    tpl = "{{#for n from 1 to {{port_count}}}}\nport {{n}}\n{{#endfor}}\n"
    assert N.render_template(tpl, {"port_count": "2"}) == "port 1\nport 2\n"


def test_render_template_for_single_value_range():
    assert N.render_template("{{#for n from 3 to 3}}\n{{n}}\n{{#endfor}}\n", {}) == "3\n"


def test_render_template_inline_for():
    """单行内联 for（指令必须在行首，见 test_known_issues）。"""
    assert N.render_template("{{#for n from 1 to 3}}{{n}} {{#endfor}}\n", {}) == "1 2 3 \n"


def test_render_template_inline_for_with_variable_range():
    assert N.render_template("{{#for n from 1 to {{pc}}}}v{{n}} {{#endfor}}\n",
                             {"pc": "2"}) == "v1 v2 \n"


def test_render_template_nested_for_in_if():
    tpl = ("{{#if a}}\n"
           "{{#for n from 1 to 2}}\n"
           "port {{n}}\n"
           "{{#endfor}}\n"
           "{{#endif}}\n")
    assert N.render_template(tpl, {"a": "1"}) == "port 1\nport 2\n"
    assert N.render_template(tpl, {"a": ""}) == "\n"     # 整块跳过，只剩模板尾换行


def test_render_template_nested_if_in_for():
    tpl = ("{{#for i from 1 to 2}}\n"
           "port {{i}}\n"
           "{{#if a=1}}\nyes\n{{#else}}\nno\n{{#endif}}\n"
           "{{#endfor}}\n")
    assert N.render_template(tpl, {"a": "2"}) == "port 1\nno\nport 2\nno\n"
    assert N.render_template(tpl, {"a": "1"}) == "port 1\nyes\nport 2\nyes\n"


def test_render_template_keeps_indentation_of_body_lines():
    tpl = "{{#if a}}\n  vlan {{mgmt_vlan}}\n{{#else}}\n  vlan 999\n{{#endif}}\n"
    assert N.render_template(tpl, {"a": "1", "mgmt_vlan": "11"}) == "  vlan 11\n"
    assert N.render_template(tpl, {"a": "", "mgmt_vlan": "11"}) == "  vlan 999\n"


# ── 报错分支 ──

def test_directive_line_with_body_raises_if():
    with pytest.raises(ValueError, match="指令行不能带内容"):
        N.render_template("{{#if a}} body\nx\n{{#endif}}\n", {"a": "1"})


def test_directive_line_with_body_raises_for():
    with pytest.raises(ValueError, match="指令行不能带内容"):
        N.render_template("{{#for n from 1 to 2}} body\nx\n{{#endfor}}\n", {})


def test_missing_endif_raises():
    with pytest.raises(ValueError, match="缺少块结束符"):
        N.render_template("{{#if a}}\nA\n", {"a": "1"})


def test_missing_endfor_raises():
    with pytest.raises(ValueError, match="缺少块结束符"):
        N.render_template("{{#for n from 1 to 2}}\nA\n", {})


def test_for_range_not_number_raises():
    with pytest.raises(ValueError, match="范围不是数字"):
        N.render_template("{{#for n from 1 to {{port_count}}}}\nx\n{{#endfor}}\n",
                          {"port_count": "abc"})


def test_for_reversed_range_raises():
    with pytest.raises(ValueError, match="范围无效"):
        N.render_template("{{#for n from 5 to 2}}\nx\n{{#endfor}}\n", {})


def test_single_line_block_with_nested_directive_raises():
    with pytest.raises(ValueError, match="单行块不支持嵌套指令"):
        N.render_template("{{#if a}}{{#if b}}x{{#endif}}{{#endif}}\n", {"a": "1", "b": "1"})


# ────────────────────────── build_switch_config ──────────────────────────

def test_build_switch_config_renders_standard_fields(make_dev):
    dev = make_dev(1, mgmt_vlan="11", access_vlan="12")
    assert N.build_switch_config(SMALL_TPL, dev) == (
        "sysname SW-1-1F\n"
        "vlan batch 11 12\n"
        "interface Vlanif11\n"
        " ip address 192.0.2.11 255.255.255.0\n"
        " ip route-static 0.0.0.0 192.0.2.254\n"
    )


def test_build_switch_config_uses_extras(make_dev):
    dev = make_dev(1, _extras={"port_count": "24"})
    assert N.build_switch_config("port {{port_count}}\n", dev) == "port 24\n"


def test_build_switch_config_for_loop_with_variable_range(make_dev):
    tpl = ("port_count {{port_count}}\n"
           "{{#for n from 1 to {{port_count}}}}\n"
           "interface GigabitEthernet0/0/{{n}}\n"
           "{{#endfor}}\n")
    dev = make_dev(1, _extras={"port_count": "2"})
    assert N.build_switch_config(tpl, dev) == (
        "port_count 2\n"
        "interface GigabitEthernet0/0/1\n"
        "interface GigabitEthernet0/0/2\n"
    )


def test_build_switch_config_leftover_placeholder_raises(make_dev):
    with pytest.raises(ValueError, match="模板缺少变量: nope"):
        N.build_switch_config("x {{nope}}\n", make_dev(1))


def test_build_switch_config_leftover_reports_all_missing(make_dev):
    with pytest.raises(ValueError, match="a, b"):
        N.build_switch_config("{{a}} {{b}}\n", make_dev(1))


def test_build_switch_config_vlan_batch_required():
    with pytest.raises(ValueError, match="vlan_batch"):
        N.build_switch_config("vlan {{vlan_batch}}\n",
                              {"hostname": "SW-1-1F", "mgmt_ip": "192.0.2.11"})


def test_build_switch_config_does_not_mutate_device(make_dev):
    dev = make_dev(1)
    snapshot = dict(dev)
    N.build_switch_config(SMALL_TPL, dev)
    assert dev == snapshot
