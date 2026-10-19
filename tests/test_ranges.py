"""批量编号（{1-10} / {01-10} / {1,3}）测试。

_parse_range_values / expand_ranges / expand_device / count_expansions / check_range_lengths
"""

import time

import pytest

import nadt as N


# ────────────────────────── _parse_range_values ──────────────────────────

@pytest.mark.parametrize("expr,expected", [
    ("1-3", ["1", "2", "3"]),
    ("1", ["1"]),
    ("1,10", ["1", "10"]),
    ("1-2,9", ["1", "2", "9"]),
    ("5-3", ["3", "4", "5"]),                                  # 反向自动纠正
    ("01-10", [f"{i:02d}" for i in range(1, 11)]),              # 零填充
    ("1-005", [f"{i:03d}" for i in range(1, 6)]),
    (" 1 - 2 ", ["1", "2"]),                                   # 允许空格
    (",1,,3,", ["1", "3"]),                                    # 空片段忽略
    ("", []),
    ("a-b", []),                                               # 非法 → 空列表
])
def test_parse_range_values(expr, expected):
    assert N._parse_range_values(expr) == expected


# ────────────────────────── expand_ranges ──────────────────────────

def test_expand_ranges_single():
    assert N.expand_ranges("192.0.2.{1-3}") == ["192.0.2.1", "192.0.2.2", "192.0.2.3"]


def test_expand_ranges_cartesian_product():
    assert N.expand_ranges("SW-{1,2}-{a,b}") == ["SW-1-a", "SW-1-b", "SW-2-a", "SW-2-b"]


def test_expand_ranges_zero_padded_two_groups():
    assert N.expand_ranges("SW-{1-2}-{01-02}") == ["SW-1-01", "SW-1-02", "SW-2-01", "SW-2-02"]


@pytest.mark.parametrize("text", ["SW-1-1F", "SW-{1", "SW-{}", "SW-{a-b}", ""])
def test_expand_ranges_without_valid_range_returns_input(text):
    """无范围语法 / 非法范围 → 原样返回单个元素（绝不丢数据）。"""
    assert N.expand_ranges(text) == [text]


# ────────────────────────── expand_device ──────────────────────────

def test_expand_device_no_range_returns_single():
    dev = {"hostname": "SW-1-1F", "mgmt_ip": "192.0.2.11"}
    out = N.expand_device(dev)
    assert out == [dev]


def test_expand_device_parallel_fields():
    dev = {"hostname": "SW-{1-3}", "mgmt_ip": "192.0.2.{1-3}", "mgmt_vlan": "11"}
    out = N.expand_device(dev)
    assert [d["hostname"] for d in out] == ["SW-1", "SW-2", "SW-3"]
    assert [d["mgmt_ip"] for d in out] == ["192.0.2.1", "192.0.2.2", "192.0.2.3"]
    assert all(d["mgmt_vlan"] == "11" for d in out)


def test_expand_device_uses_shortest_field_count():
    """字段范围长度不一致时按最短对齐（配合 check_range_lengths 的告警）。"""
    dev = {"hostname": "SW-{1-5}", "_extras": {"port": "{10-12}"}}
    out = N.expand_device(dev)
    assert len(out) == 3
    assert [d["hostname"] for d in out] == ["SW-1", "SW-2", "SW-3"]
    assert [d["_extras"]["port"] for d in out] == ["10", "11", "12"]


def test_expand_device_expands_extras_fields():
    dev = {"hostname": "SW-1-1F", "_extras": {"mgmt_ip": "192.0.2.{1-2}"}}
    out = N.expand_device(dev)
    assert [d["_extras"]["mgmt_ip"] for d in out] == ["192.0.2.1", "192.0.2.2"]
    assert all(d["hostname"] == "SW-1-1F" for d in out)


def test_expand_device_does_not_mutate_source():
    dev = {"hostname": "SW-{1-2}", "_extras": {"port": "{7-8}"}}
    out = N.expand_device(dev)
    assert dev["hostname"] == "SW-{1-2}"
    assert dev["_extras"] == {"port": "{7-8}"}
    assert len(out) == 2
    assert out[0] is not out[1]
    out[0]["_extras"]["port"] = "999"
    assert out[1]["_extras"]["port"] == "8"      # 每台设备的 _extras 独立副本


# ────────────────────────── count_expansions ──────────────────────────

@pytest.mark.parametrize("dev,expected", [
    ({"hostname": "SW-1-1F"}, 1),
    ({"hostname": "SW-{1-2}"}, 2),
    ({"hostname": "SW-{1-2}", "_extras": {"port": "{1-5}"}}, 2),          # 取最短
    ({"hostname": "SW-{1-1000}", "mgmt_ip": "192.0.2.{1-1000}"}, 1000),
    ({"hostname": "SW-{a-b}"}, 1),                                        # 非法范围按 1 计
    ({"_extras": {"port": "{1-10}"}}, 10),
])
def test_count_expansions(dev, expected):
    assert N.count_expansions(dev) == expected


def test_count_expansions_large_range_is_estimated_not_expanded():
    """{1-1000000} 必须能轻量预估（不实际展开成 100 万台设备）。"""
    started = time.time()
    assert N.count_expansions({"hostname": "SW-{1-1000000}"}) == 1_000_000
    assert time.time() - started < 10      # 不设死上限，只要能算完即可


def test_count_expansions_multi_range_multiplies_within_field():
    assert N.count_expansions({"hostname": "SW-{1-1000}-{1-1000}"}) == 1_000_000


# ────────────────────────── check_range_lengths ──────────────────────────

def test_check_range_lengths_consistent_returns_empty():
    assert N.check_range_lengths({"hostname": "SW-{1-2}", "mgmt_ip": "192.0.2.{1-2}"}) == ""
    assert N.check_range_lengths({"hostname": "SW-1-1F"}) == ""


def test_check_range_lengths_mismatch_warns():
    msg = N.check_range_lengths({"hostname": "SW-{1-2}", "mgmt_ip": "192.0.2.{1-5}"})
    assert "范围长度不一致" in msg
    assert "按最短 2 台生成" in msg


def test_check_range_lengths_covers_extras_with_prefix():
    msg = N.check_range_lengths({"hostname": "SW-{1-3}", "_extras": {"port": "{1-2}"}})
    assert "扩展.port" in msg
