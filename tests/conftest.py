"""NADT 测试公共夹具 / shared fixtures.

两条硬规则（每次跑测试都自动强制）:
  1. 只用虚构数据：主机名 SW-1-1F、ESN 0000-1111-2222、网段 192.0.2.x。
  2. 绝不往仓库目录写文件：
     - 所有写盘路径常量（BASE_DIR/TEMPLATE_DIR/DEFAULT_TEMPLATE/OUTPUT_DIR/DB_PATH）
       在每个用例里 monkeypatch 到 tmp_path；
     - 每个用例前后自动快照仓库真实数据目录（devices.db / templates / output /
       devices / backups），有任何新增/删除/改动立刻失败。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import nadt  # noqa: E402  必须在 sys.path 补好之后再导入 / import after sys.path setup

# 仓库里放真实数据或生成物的路径：测试绝不能碰
GUARDED = [
    REPO_ROOT / "devices.db",
    REPO_ROOT / "templates",
    REPO_ROOT / "output",
    REPO_ROOT / "devices",
    REPO_ROOT / "backups",
]


def _snapshot() -> dict:
    """(路径 → (大小, mtime_ns))，用于用例前后比对。"""
    snap: dict = {}
    for p in GUARDED:
        if p.is_dir():
            for f in p.rglob("*"):
                if f.is_file():
                    st = f.stat()
                    snap[str(f)] = (st.st_size, st.st_mtime_ns)
        elif p.exists():
            st = p.stat()
            snap[str(p)] = (st.st_size, st.st_mtime_ns)
    return snap


@pytest.fixture(autouse=True)
def no_repo_writes():
    """用例前后比对仓库真实文件：任何写盘都让用例失败。"""
    before = _snapshot()
    yield
    after = _snapshot()
    added = sorted(set(after) - set(before))
    removed = sorted(set(before) - set(after))
    changed = sorted(k for k in set(before) & set(after) if before[k] != after[k])
    assert not (added or removed or changed), (
        f"测试改动了仓库文件 → 新增={added} 删除={removed} 改动={changed}")


@pytest.fixture(autouse=True)
def redirect_paths(tmp_path, monkeypatch):
    """把 nadt 的写盘路径常量全部重定向到 tmp_path（templates/ output/ devices.db）。"""
    tpl = tmp_path / "templates"
    out = tmp_path / "output"
    tpl.mkdir()
    out.mkdir()
    monkeypatch.setattr(nadt, "BASE_DIR", str(tmp_path), raising=False)
    monkeypatch.setattr(nadt, "TEMPLATE_DIR", str(tpl))
    monkeypatch.setattr(nadt, "DEFAULT_TEMPLATE", str(tpl / "access-switch.cfg"))
    monkeypatch.setattr(nadt, "OUTPUT_DIR", str(out))
    monkeypatch.setattr(nadt, "DB_PATH", str(tmp_path / "devices.db"))
    return tmp_path


@pytest.fixture
def nd():
    """被 redirect_paths 处理过的 nadt 模块。"""
    return nadt


@pytest.fixture
def tpl_dir(redirect_paths):
    """重定向后的模板目录（tmp_path/templates）。"""
    return redirect_paths / "templates"


@pytest.fixture
def make_dev():
    """虚构设备工厂：/tmp 级别的假值，index 决定序号（1 → SW-1-1F / 0000-1111-2201）。"""
    def _make(idx: int = 1, **kw) -> dict:
        dev = {
            "esn": f"0000-1111-22{idx:02d}",
            "hostname": f"SW-1-{idx}F",
            "type": "AP-31U",
            "mgmt_ip": f"192.0.2.{10 + idx}",
            "mgmt_vlan": "11",
            "access_vlan": "12",
            "mask": "",
            "gateway": "",
            "note": "",
        }
        dev.update(kw)
        return dev
    return _make
