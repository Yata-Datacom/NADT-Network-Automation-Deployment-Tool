"""开局引导文件测试：build_lswnet（ESN 映射表）与 build_usb_ini（U盘 EasyDeploy ini）。"""

import pytest

import nadt as N

ESN1, ESN2 = "0000-1111-2222", "0000-1111-3333"
HOST1, HOST2 = "SW-1-1F", "SW-1-2F"
IP1 = "192.0.2.11"


# ────────────────────────── build_lswnet ──────────────────────────

def test_build_lswnet_format():
    devices = [{"esn": ESN1, "hostname": HOST1}, {"esn": ESN2, "hostname": HOST2}]
    assert N.build_lswnet(devices) == (
        f"esn={ESN1};cfgfile={HOST1}.cfg;\n"
        f"esn={ESN2};cfgfile={HOST2}.cfg;\n"
    )


def test_build_lswnet_empty_returns_empty_string():
    assert N.build_lswnet([]) == ""


def test_build_lswnet_skips_incomplete_devices():
    devices = [{"esn": ESN1, "hostname": HOST1},
               {"esn": "", "hostname": HOST2},        # 无 ESN → 无法匹配
               {"esn": ESN2, "hostname": ""},         # 无主机名 → 无 cfgfile
               {"esn": ESN1, "hostname": HOST2}]      # 完整
    out = N.build_lswnet(devices)
    assert out == f"esn={ESN1};cfgfile={HOST1}.cfg;\nesn={ESN1};cfgfile={HOST2}.cfg;\n"


def test_build_lswnet_global_vrp_and_patch():
    devices = [{"esn": ESN1, "hostname": HOST1}]
    out = N.build_lswnet(devices, vrpfile="vrp.cc", patchfile="patch.pat")
    assert out == f"esn={ESN1};vrpfile=vrp.cc;patchfile=patch.pat;cfgfile={HOST1}.cfg;\n"


def test_build_lswnet_device_extra_overrides_global():
    devices = [{"esn": ESN1, "hostname": HOST1, "_extras": {"vrp_file": "dev.cc",
                                                            "patch_file": "dev.pat"}}]
    out = N.build_lswnet(devices, vrpfile="global.cc", patchfile="global.pat")
    assert out == f"esn={ESN1};vrpfile=dev.cc;patchfile=dev.pat;cfgfile={HOST1}.cfg;\n"


def test_build_lswnet_standard_field_vrp_file():
    devices = [{"esn": ESN1, "hostname": HOST1, "vrp_file": "col.cc"}]
    assert N.build_lswnet(devices) == f"esn={ESN1};vrpfile=col.cc;cfgfile={HOST1}.cfg;\n"


def test_build_lswnet_strips_whitespace():
    devices = [{"esn": f"  {ESN1}  ", "hostname": f" {HOST1} "}]
    assert N.build_lswnet(devices) == f"esn={ESN1};cfgfile={HOST1}.cfg;\n"


def test_build_lswnet_agnostic_to_extra_fields(make_dev):
    dev = make_dev(1, _extras={"port_count": "24", "vrp_file": ""})
    assert N.build_lswnet([dev]) == f"esn={dev['esn']};cfgfile={dev['hostname']}.cfg;\n"


# ────────────────────────── build_usb_ini ──────────────────────────

def test_build_usb_ini_single_device_exact_output():
    """usb_config.ini 内容固定：;BEGIN DC / [GLOBAL CONFIG] / FILESERVER / [DEVICEn] / ;END DC。"""
    ini = N.build_usb_ini([{"esn": ESN1, "hostname": HOST1}])
    assert ini == (
        ";BEGIN DC\r\n"
        "[GLOBAL CONFIG]\r\n"
        "FILESERVER=file:/usb:\r\n"
        "\r\n"
        "[DEVICE0 DESCRIPTION]\r\n"
        f"ESN={ESN1}\r\n"
        "DEVICETYPE=DEFAULT\r\n"
        f"SYSTEM-CONFIG={HOST1}.cfg\r\n"
        ";END DC\r\n"
    )


def test_build_usb_ini_is_strict_crlf():
    """华为 EasyDeploy 要求 CRLF —— 任何裸 LF 都会让 U盘开局失败。"""
    ini = N.build_usb_ini([{"esn": ESN1, "hostname": HOST1},
                           {"esn": ESN2, "hostname": HOST2}])
    assert "\n" not in ini.replace("\r\n", "")
    assert ini.endswith("\r\n")
    assert ini.count("\r\n") == len(ini.split("\r\n")) - 1


def test_build_usb_ini_empty_device_list():
    ini = N.build_usb_ini([])
    assert ini == ";BEGIN DC\r\n[GLOBAL CONFIG]\r\nFILESERVER=file:/usb:\r\n;END DC\r\n"


def test_build_usb_ini_mac_only_falls_back_to_default_esn():
    """无 ESN 时 ESN=DEFAULT + 华为格式 MAC=xxxx-xxxx-xxxx。"""
    ini = N.build_usb_ini([{"hostname": HOST2, "mac": "AABB.CCDD.EEFF"}])
    assert "ESN=DEFAULT\r\n" in ini
    assert "MAC=aabb-ccdd-eeff\r\n" in ini
    assert f"SYSTEM-CONFIG={HOST2}.cfg\r\n" in ini
    assert "[DEVICE0 DESCRIPTION]" in ini


def test_build_usb_ini_mac_from_extras():
    ini = N.build_usb_ini([{"esn": ESN1, "hostname": HOST1,
                            "_extras": {"mac": "aa:bb:cc:dd:ee:ff"}}])
    assert "MAC=aabb-ccdd-eeff\r\n" in ini
    assert f"ESN={ESN1}\r\n" in ini


def test_build_usb_ini_without_esn_generates_default_block():
    ini = N.build_usb_ini([{"hostname": HOST1, "mac": "not-a-mac"}])
    assert "ESN=DEFAULT\r\n" in ini
    assert f"SYSTEM-CONFIG={HOST1}.cfg\r\n" in ini


def test_build_usb_ini_blocks_are_numbered_sequentially():
    ini = N.build_usb_ini([{"esn": ESN1, "hostname": HOST1},
                           {"esn": ESN2, "hostname": HOST2},
                           {"esn": "0000-1111-4444", "hostname": "SW-1-4F"}])
    assert "[DEVICE0 DESCRIPTION]" in ini
    assert "[DEVICE1 DESCRIPTION]" in ini
    assert "[DEVICE2 DESCRIPTION]" in ini
    assert ini.index("[DEVICE0 DESCRIPTION]") < ini.index("[DEVICE1 DESCRIPTION]")


def test_build_usb_ini_skips_devices_without_esn_and_mac_or_hostname():
    ini = N.build_usb_ini([{"hostname": HOST1},                # 无 ESN 无 MAC
                           {"esn": ESN2, "hostname": ""},      # 无主机名
                           {"esn": ESN1, "hostname": HOST1}])
    assert ini.count("[DEVICE") == 1
    assert "[DEVICE0 DESCRIPTION]" in ini
    assert f"ESN={ESN1}\r\n" in ini


def test_build_usb_ini_custom_fileserver():
    ini = N.build_usb_ini([], fileserver="file:/flash:")
    assert "FILESERVER=file:/flash:\r\n" in ini


def test_build_usb_ini_software_and_pat_from_extras():
    ini = N.build_usb_ini([{"esn": ESN1, "hostname": HOST1,
                            "_extras": {"system_software": "vrp.cc",
                                        "system_pat": "patch.pat",
                                        "stack_member_id": "2"}}])
    for line in ("SYSTEM-SOFTWARE=vrp.cc\r\n", "SYSTEM-PAT=patch.pat\r\n",
                 "STACK-MEMBER-ID=2\r\n"):
        assert line in ini
    # 顺序：ESN → DEVICETYPE → SYSTEM-SOFTWARE → SYSTEM-CONFIG → SYSTEM-PAT → STACK-MEMBER-ID
    order = [ini.index(k) for k in ("ESN=", "DEVICETYPE=", "SYSTEM-SOFTWARE=",
                                    "SYSTEM-CONFIG=", "SYSTEM-PAT=", "STACK-MEMBER-ID=")]
    assert order == sorted(order)


def test_build_usb_ini_global_software_defaults():
    ini = N.build_usb_ini([{"esn": ESN1, "hostname": HOST1}],
                          system_software="global.cc", system_pat="global.pat")
    assert "SYSTEM-SOFTWARE=global.cc\r\n" in ini
    assert "SYSTEM-PAT=global.pat\r\n" in ini


def test_build_usb_ini_without_software_omits_lines():
    ini = N.build_usb_ini([{"esn": ESN1, "hostname": HOST1}])
    assert "SYSTEM-SOFTWARE" not in ini
    assert "SYSTEM-PAT" not in ini
    assert "STACK-MEMBER-ID" not in ini


def test_build_usb_ini_stack_member_id_standard_field():
    ini = N.build_usb_ini([{"esn": ESN1, "hostname": HOST1, "stack_member_id": "3"}])
    assert "STACK-MEMBER-ID=3\r\n" in ini


@pytest.mark.parametrize("idx", [1, 2, 3])
def test_build_usb_ini_is_a_pure_function(idx, make_dev):
    """build_usb_ini 只吃内存数据、不读写磁盘（路径常量已被重定向到 tmp_path）。"""
    dev = make_dev(idx)
    ini = N.build_usb_ini([dev])
    assert f"SYSTEM-CONFIG={dev['hostname']}.cfg\r\n" in ini
    assert ini == N.build_usb_ini([dev])          # 可重复、无副作用
