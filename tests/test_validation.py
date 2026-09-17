"""设备字段兜底（derive_defaults）与清单校验（validate_devices）测试。"""


import nadt as N


# ────────────────────────── derive_defaults ──────────────────────────

def test_derive_defaults_fills_all_standard_keys():
    d = N.derive_defaults({"hostname": "SW-1-1F"})
    for k in N.CSV_HEADER:
        assert k in d
    assert d["hostname"] == "SW-1-1F"
    assert d["mask"] == "255.255.255.0"
    assert d["gateway"] == ""            # 无管理 IP → 不猜网关
    assert d["vlan_batch"] == ""


def test_derive_defaults_gateway_from_mgmt_ip():
    d = N.derive_defaults({"hostname": "SW-1-1F", "mgmt_ip": "192.0.2.11"})
    assert d["gateway"] == "192.0.2.254"


def test_derive_defaults_keeps_explicit_gateway():
    d = N.derive_defaults({"mgmt_ip": "192.0.2.11", "gateway": "192.0.2.1"})
    assert d["gateway"] == "192.0.2.1"


def test_derive_defaults_keeps_explicit_mask():
    d = N.derive_defaults({"mgmt_ip": "192.0.2.11", "mask": "255.255.0.0"})
    assert d["mask"] == "255.255.0.0"


def test_derive_defaults_vlan_batch():
    assert N.derive_defaults({"mgmt_vlan": "11", "access_vlan": "12"})["vlan_batch"] == "11 12"
    assert N.derive_defaults({"mgmt_vlan": "11"})["vlan_batch"] == "11"
    assert N.derive_defaults({"access_vlan": "12"})["vlan_batch"] == "12"


def test_derive_defaults_strips_whitespace():
    d = N.derive_defaults({"hostname": "  SW-1-1F  ", "mgmt_ip": " 192.0.2.11 ",
                           "mgmt_vlan": " 11 "})
    assert d["hostname"] == "SW-1-1F"
    assert d["vlan_batch"] == "11"      # access_vlan 缺失 → 只有管理 VLAN
    assert d["mask"] == "255.255.255.0"


def test_derive_defaults_gateway_ignores_invalid_ip():
    d = N.derive_defaults({"mgmt_ip": "192.0.2.{1-100}"})
    assert d["gateway"] == ""            # 批量编号尚未展开 → 不派生病网关


def test_derive_defaults_does_not_mutate_input():
    src = {"hostname": "SW-1-1F"}
    out = N.derive_defaults(src)
    assert src == {"hostname": "SW-1-1F"}
    assert out is not src


def test_derive_defaults_keeps_extras_and_custom_keys():
    d = N.derive_defaults({"hostname": "SW-1-1F", "_extras": {"port_count": "24"}})
    assert d["_extras"] == {"port_count": "24"}


# ────────────────────────── validate_devices ──────────────────────────

def _by_severity(problems, sev):
    return [p for p in problems if p[0] == sev]


def test_validate_clean_list_has_no_problems(make_dev):
    devices = [make_dev(1), make_dev(2)]
    assert N.validate_devices(devices) == []


def test_validate_detects_duplicate_esn():
    devices = [{"esn": "0000-1111-2222", "hostname": "SW-1-1F", "mgmt_ip": "192.0.2.11"},
               {"esn": "0000-1111-2222", "hostname": "SW-1-2F", "mgmt_ip": "192.0.2.12"}]
    problems = N.validate_devices(devices)
    msgs = [m for sev, _, m in problems if sev == "error"]
    assert any("ESN 重复: 0000-1111-2222" in m for m in msgs)


def test_validate_detects_duplicate_ip_and_hostname():
    devices = [{"esn": "0000-1111-2222", "hostname": "SW-1-1F", "mgmt_ip": "192.0.2.11"},
               {"esn": "0000-1111-3333", "hostname": "SW-1-1F", "mgmt_ip": "192.0.2.11"}]
    msgs = [m for _, _, m in N.validate_devices(devices)]
    assert any("主机名重复: SW-1-1F" in m for m in msgs)
    assert any("管理 IP 重复: 192.0.2.11" in m for m in msgs)


def test_validate_missing_hostname_and_ip_are_errors():
    problems = N.validate_devices([{"esn": "0000-1111-2222"}])
    errors = [m for _, _, m in _by_severity(problems, "error")]
    assert "缺主机名" in errors
    assert "缺管理 IP" in errors


def test_validate_bad_ip_format_is_error():
    problems = N.validate_devices([{"esn": "0000-1111-2222", "hostname": "SW-1-1F",
                                    "mgmt_ip": "192.0.2"}])
    assert ("error", "SW-1-1F", "管理 IP 格式错误: 192.0.2") in problems


def test_validate_range_ip_gets_hint():
    problems = N.validate_devices([{"esn": "0000-1111-2222", "hostname": "SW-1-1F",
                                    "mgmt_ip": "192.0.2.{1-100}"}])
    msg = [m for _, _, m in problems if "格式错误" in m][0]
    assert "花括号格式" in msg


def test_validate_missing_esn_is_warning_only():
    problems = N.validate_devices([{"hostname": "SW-1-1F", "mgmt_ip": "192.0.2.11"}])
    assert len(problems) == 1
    sev, tag, msg = problems[0]
    assert sev == "warning"
    assert tag == "SW-1-1F"
    assert "缺 ESN" in msg


def test_validate_severity_values_are_known():
    devices = [{"esn": "", "hostname": "", "mgmt_ip": "not-an-ip"},
               {"esn": "0000-1111-2222", "hostname": "SW-1-1F", "mgmt_ip": "192.0.2.11"}]
    for sev, _tag, _msg in N.validate_devices(devices):
        assert sev in ("error", "warning")


def test_validate_empty_list():
    assert N.validate_devices([]) == []
