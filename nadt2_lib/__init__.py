"""NADT 2.0 — Network Automation Deployment Tool.

模块划分：
  model.py     项目文件（.nadt）、设备、模板、凭据、设置、recent
  secrets.py   Windows DPAPI 凭据保护（不可用时降级为本地明文标记）
  services.py  校验、生成计划、Diff、批量操作
  ui/          Tkinter 界面（方案 / 部署 / 项目）
"""

__version__ = "2.0.1"
APP_NAME = "NADT 2.0"
