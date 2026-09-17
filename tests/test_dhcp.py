"""DHCP 配置片段 / MAC 规范化 / dnsmasq / ISC dhcpd 生成测试。"""

import pytest

import nadt as N

HOST1, HOST2 = "SW-1-1F", "SW-1-2F"
ESN1, ESN2 = "0000-1111-2222", "0000-1111-3333"


# ────────────────────────── dhcp_snippet ──────────────────────────

def test_dhcp_snippet_mode_67_pxe():
    out = N.dhcp_snippet("192.0.2.254", mode="67")
    assert "option 66 ip-address 192.0.2.254" in out
    assert "option 67 ascii lswnet.cfg" in out
    assert "option 66 + option 67（通用 PXE）" in out
    assert "option 146" not in out.split("# 通用说明")[0]


def test_dhcp_snippet_mode_146_easydeploy():
    out = N.dhcp_snippet("192.0.2.254", mode="146")
    assert "option 146 ascii opervalue=1;delaytime=0;netfile=lswnet.cfg;" in out
    assert "option 66 ascii tftp://192.0.2.254" in out
    assert "option 66 + option 146（华为 EasyDeploy）" in out


def test_dhcp_snippet_mode_146_custom_fileserver_url():
    out = N.dhcp_snippet("192.0.2.254", mode="146",
                         fileserver_url="sftp://user@192.0.2.5:22/")
    assert "option 66 ascii sftp://user@192.0.2.5:22/" in out
    assert "tftp://192.0.2.254" not in out


def test_dhcp_snippet_mode_146_blank_fileserver_falls_back_to_tftp():
    out = N.dhcp_snippet("192.0.2.254", mode="146", fileserver_url="   ")
    assert "option 66 ascii tftp://192.0.2.254" in out


def test_dhcp_snippet_unknown_mode_behaves_like_67():
    """当前实现：非 '146' 一律走 option 67 分支。"""
    out = N.dhcp_snippet("192.0.2.254", mode="weird")
    assert "option 67 ascii lswnet.cfg" in out


def test_dhcp_snippet_custom_parameters():
    out = N.dhcp_snippet("192.0.2.254", bootfile="startup.cfg", mgmt_vlan="9",
                         network="192.0.2.0", gateway="192.0.2.254")
    assert "option 67 ascii startup.cfg" in out
    assert "interface Vlanif 9" in out
    assert "network 192.0.2.0 mask 255.255.255.0" in out
    assert "gateway-list 192.0.2.254" in out
    assert 'dhcp-option=67,"startup.cfg"' in out          # 通用说明段也同步


def test_dhcp_snippet_defaults_use_module_constants():
    out = N.dhcp_snippet("192.0.2.254")
    assert N.BOOTFILE in out
    assert f"option 67 ascii {N.BOOTFILE}" in out


def test_dhcp_snippet_mentions_other_servers():
    out = N.dhcp_snippet("192.0.2.254")
    assert "ISC dhcpd" in out and "dnsmasq" in out and "Windows DHCP" in out


# ────────────────────────── normalize_mac ──────────────────────────

@pytest.mark.parametrize("raw", ["aa:bb:cc:dd:ee:ff", "AA:BB:CC:DD:EE:FF", "aabb.ccdd.eeff",
                                 "aa-bb-cc-dd-ee-ff", "aabbccddeeff", "AA-BB-CC-DD-EE-FF"])
def test_normalize_mac_colon_accepts_any_format(raw):
    assert N.normalize_mac(raw, "colon") == "aa:bb:cc:dd:ee:ff"


@pytest.mark.parametrize("raw", ["aa:bb:cc:dd:ee:ff", "AABB.CCDD.EEFF", "aabbccddeeff"])
def test_normalize_mac_huawei(raw):
    assert N.normalize_mac(raw, "huawei") == "aabb-ccdd-eeff"


def test_normalize_mac_default_style_is_colon():
    assert N.normalize_mac("aabb.ccdd.eeff") == "aa:bb:cc:dd:ee:ff"


@pytest.mark.parametrize("raw", ["", None, "zz:zz:zz:zz:zz:zz", "aa:bb:cc:dd:ee", "aabbccddeeff00",
                                 "not-a-mac"])
def test_normalize_mac_invalid_returns_empty(raw):
    assert N.normalize_mac(raw) == ""
    assert N.normalize_mac(raw, "huawei") == ""


# ────────────────────────── build_dhcp_dnsmasq ──────────────────────────

def test_build_dhcp_dnsmasq_basic():
    out = N.build_dhcp_dnsmasq([], "192.0.2.254")
    assert "dhcp-option=66,192.0.2.254" in out
    assert "dhcp-option=67" not in out              # 未指定 bootfile → 不输出
    assert "dhcp-range" not in out
    assert out.endswith("\n")


def test_build_dhcp_dnsmasq_with_bootfile_and_network():
    out = N.build_dhcp_dnsmasq([], "192.0.2.254", bootfile="lswnet.cfg",
                               network="192.0.2.0", gateway="192.0.2.254",
                               range_start="192.0.2.100", range_end="192.0.2.200")
    assert 'dhcp-option=67,"lswnet.cfg"' in out
    assert "dhcp-range=192.0.2.100,192.0.2.200,12h" in out
    assert "dhcp-option=3,192.0.2.254" in out


def test_build_dhcp_dnsmasq_range_requires_all_params():
    out = N.build_dhcp_dnsmasq([], "192.0.2.254", network="192.0.2.0",
                               gateway="192.0.2.254", range_start="192.0.2.100")
    assert "dhcp-range" not in out


def test_build_dhcp_dnsmasq_mac_bindings_and_counts():
    devices = [{"hostname": HOST1, "mgmt_ip": "192.0.2.11",
                "_extras": {"mac": "AABB.CCDD.EEFF"}},
               {"hostname": HOST2, "mgmt_ip": "192.0.2.12"}]       # 缺 MAC
    out = N.build_dhcp_dnsmasq(devices, "192.0.2.254")
    assert "dhcp-host=aa:bb:cc:dd:ee:ff,set:SW_1_1F,192.0.2.11" in out
    assert 'dhcp-option=tag:SW_1_1F,67,"SW-1-1F.cfg"' in out
    assert "已生成 1 条 MAC 绑定；1 台设备缺 MAC" in out


def test_build_dhcp_dnsmasq_bound_without_ip():
    devices = [{"hostname": HOST1, "_extras": {"mac": "aabbccddeeff"}}]
    out = N.build_dhcp_dnsmasq(devices, "192.0.2.254")
    assert "dhcp-host=aa:bb:cc:dd:ee:ff,set:SW_1_1F\n" in out


# ────────────────────────── build_dhcp_isc ──────────────────────────

def test_build_dhcp_isc_global_options():
    out = N.build_dhcp_isc([], "192.0.2.254")
    assert 'option tftp-server-name "192.0.2.254";' in out
    assert "bootfile-name" not in out


def test_build_dhcp_isc_with_bootfile_and_subnet():
    out = N.build_dhcp_isc([], "192.0.2.254", bootfile="lswnet.cfg", network="192.0.2.0",
                           gateway="192.0.2.254", range_start="192.0.2.100",
                           range_end="192.0.2.200", dns="192.0.2.53", domain="lab.test")
    assert 'option bootfile-name "lswnet.cfg";' in out
    assert "subnet 192.0.2.0 netmask 255.255.255.0 {" in out
    assert "  range 192.0.2.100 192.0.2.200;" in out
    assert "  option routers 192.0.2.254;" in out
    assert "  option domain-name-servers 192.0.2.53;" in out
    assert '  option domain-name "lab.test";' in out


def test_build_dhcp_isc_host_bindings():
    devices = [{"hostname": HOST1, "mgmt_ip": "192.0.2.11",
                "_extras": {"mac": "AABB.CCDD.EEFF"}},
               {"hostname": HOST2, "mgmt_ip": "192.0.2.12"}]
    out = N.build_dhcp_isc(devices, "192.0.2.254", bootfile="lswnet.cfg")
    assert f"host {HOST1} {{" in out
    assert "  hardware ethernet aa:bb:cc:dd:ee:ff;" in out
    assert "  fixed-address 192.0.2.11;" in out
    assert '  option bootfile-name "SW-1-1F.cfg";' in out
    assert f"host {HOST2} {{" not in out               # 缺 MAC → 不生成 host 段
    assert "已生成 1 条 MAC 绑定；1 台设备缺 MAC" in out


def test_build_dhcp_isc_bound_device_without_ip_omits_fixed_address():
    devices = [{"hostname": HOST1, "_extras": {"mac": "aabbccddeeff"}}]
    out = N.build_dhcp_isc(devices, "192.0.2.254")
    assert "host SW-1-1F {" in out
    assert "fixed-address" not in out
    assert "已生成 1 条 MAC 绑定；0 台设备缺 MAC" in out


def test_build_dhcp_isc_mac_from_standard_field():
    devices = [{"hostname": HOST1, "mgmt_ip": "192.0.2.11", "mac": "aabb.ccdd.eeff"}]
    out = N.build_dhcp_isc(devices, "192.0.2.254")
    assert "hardware ethernet aa:bb:cc:dd:ee:ff;" in out


def test_build_dhcp_isc_empty_devices_counts():
    out = N.build_dhcp_isc([], "192.0.2.254")
    assert "已生成 0 条 MAC 绑定；0 台设备缺 MAC" in out


@pytest.mark.parametrize("idx", [1, 2])
def test_build_dhcp_builders_are_deterministic(idx, make_dev):
    devices = [make_dev(idx, _extras={"mac": "aabbccddeeff"})]
    assert N.build_dhcp_isc(devices, "192.0.2.254") == N.build_dhcp_isc(devices, "192.0.2.254")
    assert N.build_dhcp_dnsmasq(devices, "192.0.2.254") == \
        N.build_dhcp_dnsmasq(devices, "192.0.2.254")
    dev = devices[0]
    assert f"esn={dev['esn']};cfgfile={dev['hostname']}.cfg;" in N.build_lswnet(devices)
