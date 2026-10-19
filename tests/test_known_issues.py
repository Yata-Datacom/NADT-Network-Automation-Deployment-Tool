"""已修缺陷的回归用例（原「已知问题」清单已清空）。

这些用例原先用 characterization 手法锁定「当前有问题的行为」；随着修复逐条转成
「验证修复后的行为」。将来若发现同类**静默出错**缺陷，按同样手法：先加 xfail 锁定 → 修 → 转正。
"""

import time

import pytest

import nadt as N


def test_fixed_spaced_placeholder_substituted_and_leftover_raises():
    """已修：占位符带空格（{{ hostname }}）曾既不替换、又不被残留检测发现。"""
    cfg = N.build_switch_config("sysname {{ hostname }}\nvlan {{mgmt_vlan}}\n",
                                {"hostname": "SW-1-1F", "mgmt_vlan": "11",
                                 "mgmt_ip": "192.0.2.11"})
    assert cfg == "sysname SW-1-1F\nvlan 11\n"
    with pytest.raises(ValueError, match="模板缺少变量"):
        N.build_switch_config("sysname {{ not_defined }}\n", {"hostname": "SW-1-1F"})


def test_fixed_unprocessed_inline_directive_raises():
    """已修：指令不在行首时曾整行原样输出（残留检测只认 {{word}} 形式）。"""
    with pytest.raises(ValueError, match="模板缺少变量"):
        N.build_switch_config("sysname X{{#if a}}Y{{#endif}}Z\n", {"hostname": "SW-1-1F"})


def test_fixed_else_and_endif_with_trailing_text_raise():
    """已修：{{#else}} / {{#endif}} 行带内容会被静默丢弃（与 {{#if}}/{{#for}} 行为不一致）。"""
    with pytest.raises(ValueError, match="不能带其他内容"):
        N.render_template("{{#if a}}\nA\n{{#else}} 需要保留的文本\nB\n{{#endif}}\n", {"a": ""})
    with pytest.raises(ValueError, match="不能带其他内容"):
        N.render_template("{{#if a}}\nA\n{{#endif}} 行尾注释\n", {"a": "1"})
    # 规范写法不受影响
    assert N.render_template("{{#if a}}\nA\n{{#else}}\nB\n{{#endif}}\n", {"a": ""}) == "B\n"
    assert N.render_template("{{#if a}}\nA\n{{#endif}}\n", {"a": "1"}) == "A\n"


def test_fixed_mapping_sheet_gets_the_right_warning():
    """已修：映射表识别曾排在「无参数列」早退之后 → 真实宏工作簿拿到误导性警告。"""
    w = []
    t, d = N.parse_sheet_rows([["【lswnet】"], ["esn=0000-1111-2222;cfgfile=SW-1-1F.cfg;"]],
                              "lswnet", w)
    assert (t, d) == ({}, [])
    assert any("ESN 映射表" in x for x in w), w
    assert not any("无参数列" in x for x in w), w


def test_fixed_count_expansions_is_true_lightweight():
    """已修：count_expansions 号称「轻量预估」却会物化区间（{{1-1000000}} 峰值约 63MB）。"""
    started = time.time()
    assert N.count_expansions({"hostname": "SW-{1-1000000}"}) == 1_000_000
    # 天文区间也不该 OOM/卡死（旧实现会尝试物化一亿个字符串）
    assert N.count_expansions({"hostname": "SW-{1-100000000}"}) == 100_000_000
    assert time.time() - started < 2
    # 与真实展开结果交叉验证（小范围）
    assert N.count_expansions({"hostname": "SW-{1-5}"}) == len(N.expand_device({"hostname": "SW-{1-5}"}))


def test_fixed_invalid_mac_no_empty_mac_line():
    """已修：MAC 非空但格式非法时曾写出空的 MAC= 行（设备可能忽略或解析异常）。"""
    ini = N.build_usb_ini([{"hostname": "SW-1-1F", "esn": "0000-1111-2222", "mac": "not-a-mac"}])
    assert "MAC=" not in ini
    ok = N.build_usb_ini([{"hostname": "SW-1-1F", "esn": "0000-1111-2222", "mac": "0000-1111-2222"}])
    assert "MAC=0000-1111-2222" in ok
