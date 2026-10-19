"""清单解析 / 导出测试：Excel 宏结构表、CSV（含中文别名）、列映射、
扩展字段、模板库路径、设备库（devices.db，已重定向到 tmp_path）。
"""

import os

import pytest

import nadt as N


# ────────────────────────── parse_sheet_rows ──────────────────────────

# 宏工作流结构：A1=【模板名】，第 1 行 C 列起=参数名，A 列第 2 行起=模板文本，
# 第 2 行起 C 列起=设备各参数值（模板行与设备行共用行号）
SHEET_ROWS = [
    ["【AP-31U】", None, "{{sysname}}", "{{mgmt_ip}}", "{{mgmt_vlan}}"],
    ["sysname {{sysname}}", None, "SW-1-1F", "192.0.2.11", "11"],
    ["vlan {{mgmt_vlan}}", None, "SW-1-2F", "192.0.2.12", "11"],
]


def test_parse_sheet_rows_template_and_devices():
    warnings = []
    templates, devices = N.parse_sheet_rows(SHEET_ROWS, "AP-31U", warnings)
    assert list(templates) == ["AP-31U"]
    assert templates["AP-31U"] == "sysname {{sysname}}\nvlan {{mgmt_vlan}}"
    assert warnings == []
    assert [d["hostname"] for d in devices] == ["SW-1-1F", "SW-1-2F"]
    assert all(d["type"] == "AP-31U" for d in devices)
    assert devices[0]["_extras"] == {"sysname": "SW-1-1F", "mgmt_ip": "192.0.2.11",
                                     "mgmt_vlan": "11"}
    for k in N.CSV_HEADER:
        assert k in devices[0]


def test_parse_sheet_rows_none_cells_do_not_create_phantom_devices():
    """空单元格（openpyxl 的 None 填充行）不能变成幽灵设备（历史 965 台问题）。"""
    rows = [["【T】", None, "{{sysname}}"],
            ["sysname {{sysname}}", None, "SW-1-1F"]]
    rows += [[None] * 3 for _ in range(500)]
    templates, devices = N.parse_sheet_rows(rows, "T", [])
    assert len(devices) == 1
    assert devices[0]["hostname"] == "SW-1-1F"
    assert templates["T"] == "sysname {{sysname}}"


def test_parse_sheet_rows_empty_row_list():
    assert N.parse_sheet_rows([], "X", []) == ({}, [])


def test_parse_sheet_rows_non_template_sheet_ignored():
    assert N.parse_sheet_rows([["随便", "内容"], ["a", "b"]], "Sheet1", []) == ({}, [])


def test_parse_sheet_rows_placeholder_only_detection():
    """A 列出现 {{变量}}（没有【】标题）也算模板 sheet。"""
    rows = [["设备1", None, "{{sysname}}"], ["sysname {{sysname}}", None, "SW-1-1F"]]
    templates, devices = N.parse_sheet_rows(rows, "T2", [])
    assert "T2" in templates
    assert devices[0]["hostname"] == "SW-1-1F"


def test_parse_sheet_rows_duplicate_param_column_warns():
    warnings = []
    N.parse_sheet_rows([["【T】", None, "{{a}}", "{{a}}"], ["x", None, "1", "2"]], "T", warnings)
    assert any("参数名 a 重复" in w for w in warnings)


def test_parse_sheet_rows_without_param_columns_warns():
    warnings = []
    templates, devices = N.parse_sheet_rows([["【T】", None, None], ["x"]], "T", warnings)
    assert (templates, devices) == ({}, [])
    assert any("无参数列" in w for w in warnings)


def test_parse_sheet_rows_without_template_text_warns():
    warnings = []
    templates, devices = N.parse_sheet_rows([["【T】", None, "{{a}}"], [None, None, "1"]],
                                            "T", warnings)
    assert (templates, devices) == ({}, [])
    assert any("A 列无模板文本" in w for w in warnings)


def test_parse_sheet_rows_esn_mapping_sheet_is_skipped():
    """lswnet 这类 ESN 映射表 sheet 不应被当成模板导入。"""
    warnings = []
    rows = [["【lswnet】", None, "{{x}}"],
            ["esn=0000-1111-2222;cfgfile=SW-1-1F.cfg;", None, "1"],
            ["esn=0000-1111-3333;cfgfile=SW-1-2F.cfg;", None, "2"]]
    templates, devices = N.parse_sheet_rows(rows, "lswnet", warnings)
    assert (templates, devices) == ({}, [])
    assert any("lswnet" in w for w in warnings)


def test_parse_sheet_rows_generates_hostname_when_absent():
    rows = [["【AP-31U】", None, "{{port}}"], ["# 模板行", None, "24"]]
    _templates, devices = N.parse_sheet_rows(rows, "AP-31U", [])
    assert devices[0]["hostname"] == "AP-31U-1"
    assert devices[0]["_extras"]["port"] == "24"


def test_parse_sheet_rows_esn_param_promoted_to_standard_field():
    rows = [["【T】", None, "{{esn}}", "{{sysname}}"],
            ["x", None, "0000-1111-2222", "SW-1-1F"]]
    _templates, devices = N.parse_sheet_rows(rows, "T", [])
    assert devices[0]["esn"] == "0000-1111-2222"
    assert devices[0]["hostname"] == "SW-1-1F"


# ────────────────────────── parse_device_csv ──────────────────────────

CSV_TEXT = ("序列号,设备名,管理IP,管理VLAN,接入VLAN,备注,端口数\n"
            "0000-1111-2222,SW-1-1F,192.0.2.11,11,12,1F弱电间,24\n")


def test_parse_device_csv_chinese_aliases():
    devices = N.parse_device_csv(CSV_TEXT)
    assert len(devices) == 1
    d = devices[0]
    assert d["esn"] == "0000-1111-2222"
    assert d["hostname"] == "SW-1-1F"
    assert d["mgmt_ip"] == "192.0.2.11"
    assert d["mgmt_vlan"] == "11"
    assert d["access_vlan"] == "12"
    assert d["note"] == "1F弱电间"
    assert d["_extras"] == {"端口数": "24"}
    for k in N.CSV_HEADER:
        assert k in d


def test_parse_device_csv_missing_columns_filled_with_empty():
    devices = N.parse_device_csv("esn,hostname\n0000-1111-2222,SW-1-1F\n")
    assert devices[0]["mgmt_ip"] == ""
    assert devices[0]["mask"] == ""


def test_parse_device_csv_empty_and_blank_lines():
    assert N.parse_device_csv("") == []
    assert N.parse_device_csv("\n\n") == []
    # 全空单元格的行（逗号行）直接丢弃，不产生幽灵设备
    assert N.parse_device_csv("esn,hostname\n,\n") == []


def test_parse_device_csv_quoted_field_with_comma():
    devices = N.parse_device_csv('esn,hostname,note\n"0000-1111-2222","SW-1-1F","A栋,1楼"\n')
    assert devices[0]["note"] == "A栋,1楼"


def test_load_devices_csv_handles_utf8_bom(tmp_path):
    p = tmp_path / "devices.csv"
    p.write_bytes(CSV_TEXT.encode("utf-8-sig"))
    devices = N.load_devices_csv(str(p))
    assert devices[0]["esn"] == "0000-1111-2222"      # BOM 不污染首列表头
    assert devices[0]["hostname"] == "SW-1-1F"


# ────────────────────────── parse_device_mapped ──────────────────────────

MAPPED_TEXT = ("SN,NAME,IP,PORT_COUNT,X\n"
               "0000-1111-2222,SW-1-1F,192.0.2.11,24,ignore-me\n"
               ",,,\n")


def test_parse_device_mapped_standard_extras_and_ignore():
    mapping = {0: "esn", 1: "hostname", 2: "mgmt_ip", 3: "扩展字段", 4: "忽略"}
    devices = N.parse_device_mapped(MAPPED_TEXT, mapping)
    assert len(devices) == 1                       # 全空行被过滤
    d = devices[0]
    assert (d["esn"], d["hostname"], d["mgmt_ip"]) == ("0000-1111-2222", "SW-1-1F",
                                                       "192.0.2.11")
    assert d["_extras"] == {"port_count": "24"}    # 扩展字段 key = 原表头小写
    assert "x" not in d["_extras"]                 # 忽略列不进结果


def test_parse_device_mapped_ignores_row_without_identity():
    mapping = {0: "note", 1: "扩展字段"}
    devices = N.parse_device_mapped("备注,DATA\n,only-extra\n", mapping)
    assert devices == []


def test_parse_device_mapped_short_row_and_empty_text():
    mapping = {0: "esn", 1: "hostname", 2: "mgmt_ip"}
    devices = N.parse_device_mapped("a,b,c\n0000-1111-2222\n", mapping)
    assert devices[0]["hostname"] == ""
    assert devices[0]["mgmt_ip"] == ""
    assert N.parse_device_mapped("", mapping) == []


def test_parse_device_mapped_unknown_key_treated_as_standard_field():
    """映射里给的是标准字段名时直接落到 dev（含 type/note 等）。"""
    mapping = {0: "hostname", 1: "note"}
    devices = N.parse_device_mapped("H,N\nSW-1-1F,1F\n", mapping)
    assert devices[0]["note"] == "1F"
    assert devices[0]["_extras"] == {}


# ────────────────────────── devices_to_csv ──────────────────────────

def test_devices_to_csv_header_and_values():
    devices = N.parse_device_csv(CSV_TEXT)
    out = N.devices_to_csv(devices)
    lines = out.strip().splitlines()
    assert lines[0] == "esn,hostname,type,mgmt_ip,mgmt_vlan,access_vlan,mask,gateway,note,端口数"
    assert lines[1] == "0000-1111-2222,SW-1-1F,,192.0.2.11,11,12,,,1F弱电间,24"


def test_devices_to_csv_without_extras_uses_standard_header():
    out = N.devices_to_csv([{"esn": "0000-1111-2222", "hostname": "SW-1-1F"}])
    assert out.strip().splitlines()[0] == ",".join(N.CSV_HEADER)


def test_devices_to_csv_roundtrip():
    devices = N.parse_device_csv(CSV_TEXT)
    again = N.parse_device_csv(N.devices_to_csv(devices))
    assert again == devices


# ────────────────────────── 扩展字段文本 ──────────────────────────

def test_parse_extras_text_separators_and_comments():
    extras = N.parse_extras_text("port_count=24\n# 注释\nmac: AABB.CCDD.EEFF\n备注：1F\nbare\n")
    assert extras == {"port_count": "24", "mac": "AABB.CCDD.EEFF", "备注": "1F", "bare": ""}


def test_parse_extras_text_value_may_contain_colon():
    assert N.parse_extras_text("url=http://192.0.2.5/x") == {"url": "http://192.0.2.5/x"}


def test_format_extras_text_roundtrip():
    extras = {"port_count": "24", "vrp_file": "vrp.cc"}
    assert N.format_extras_text(extras) == "port_count=24\nvrp_file=vrp.cc"
    assert N.parse_extras_text(N.format_extras_text(extras)) == extras
    assert N.format_extras_text({}) == ""
    assert N.format_extras_text(None) == ""


# ────────────────────────── 模板库路径（TEMPLATE_DIR 已重定向） ──────────────────────────

def test_template_dir_is_redirected_into_tmp(tmp_path):
    assert os.path.abspath(N.TEMPLATE_DIR).startswith(os.path.abspath(str(tmp_path)))
    assert os.path.abspath(N.DB_PATH).startswith(os.path.abspath(str(tmp_path)))


def test_list_templates_only_cfg_sorted(tpl_dir):
    (tpl_dir / "b.cfg").write_text("b", encoding="utf-8")
    (tpl_dir / "a.cfg").write_text("a", encoding="utf-8")
    (tpl_dir / "note.txt").write_text("x", encoding="utf-8")
    assert N.list_templates() == ["a.cfg", "b.cfg"]


def test_list_templates_missing_dir_returns_empty(monkeypatch, tmp_path):
    monkeypatch.setattr(N, "TEMPLATE_DIR", str(tmp_path / "nope"))
    assert N.list_templates() == []


def test_get_template_path_strips_directory_traversal(tpl_dir):
    out = N.get_template_path("../../evil.cfg")
    assert os.path.dirname(os.path.abspath(out)) == os.path.abspath(str(tpl_dir))
    assert os.path.basename(out) == "evil.cfg"


def test_template_mapping_save_load(tpl_dir):
    assert N.load_template_mapping() == {}
    N.save_template_mapping({"AP-31U": "a.cfg"})
    assert N.load_template_mapping() == {"AP-31U": "a.cfg"}
    assert (tpl_dir / "mapping.json").exists()


def test_load_template_mapping_broken_json_returns_empty(tpl_dir):
    (tpl_dir / "mapping.json").write_text("{not json", encoding="utf-8")
    assert N.load_template_mapping() == {}


@pytest.mark.parametrize("devtype,mapping,expected", [
    ("AP-31U", {"AP-31U": "a.cfg"}, "a.cfg"),
    ("ap-31u", {"AP-31U": "b.cfg"}, "b.cfg"),      # 大小写不敏感
    ("SC-35U", {"AP-31U": "a.cfg"}, "fallback.cfg"),
    ("", {"AP-31U": "a.cfg"}, "fallback.cfg"),
    (None, {}, "fallback.cfg"),
])
def test_find_template_for_type(devtype, mapping, expected):
    assert N.find_template_for_type(devtype, mapping, "fallback.cfg") == expected


# ────────────────────────── 设备库 devices.db（重定向到 tmp_path） ──────────────────────────

def test_devices_db_roundtrip_in_tmp(tmp_path, make_dev):
    devices = [make_dev(1, _extras={"port_count": "24"}), make_dev(2)]
    N.save_devices_db(devices)
    db = tmp_path / "devices.db"
    assert db.exists()
    loaded = N.load_devices_db()
    assert [d["hostname"] for d in loaded] == ["SW-1-1F", "SW-1-2F"]
    assert loaded[0]["_extras"] == {"port_count": "24"}
    assert loaded[0]["esn"] == "0000-1111-2201"


def test_db_path_is_redirected_away_from_repo_root():
    """DB_PATH 必须指向 tmp_path；仓库根的 devices.db（真实设备库）由 conftest 的
    no_repo_writes 夹具全量保护，任何写入都会让用例失败。"""
    repo_db = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "devices.db")
    assert os.path.abspath(N.DB_PATH) != os.path.abspath(repo_db)
