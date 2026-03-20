# wenxing-web

地基基础分析及选型报告生成系统。

本项目提供两套入口：

1. Web 入口：基于 FastAPI + Uvicorn + IIS 反向代理，支持上传勘察 Excel、自动分析并生成报告。
2. 桌面入口：保留本地桌面分析界面，便于单机使用和调试。

## 主要功能

- 解析建筑物、钻孔、地层和各孔分层等多工作表勘察数据
- 自动识别持力层、软弱下卧层和特殊地层工况
- 执行承载力、当量模量、沉降和变形相关分析
- 自动输出标准化地基基础分析报告
- 提供 Web 上传、任务执行、结果下载和基础运行状态接口

## 运行环境

- Python 3.11+
- Windows Server / Windows
- IIS（仅 Web 部署时需要）

## 安装依赖

```powershell
pip install -r requirements.txt
```

## 启动方式

### Web 服务

```powershell
python run_web.py
```

默认监听：

```text
http://0.0.0.0:8086
```

### 服务器启动脚本

```powershell
.\start_server.ps1
```

或直接双击：

```text
start_server.bat
```

### 桌面界面

```powershell
python main.py
```

## 关键目录

- `api/`：Web API 与静态页面
- `wenxing2.py`：核心分析与报告生成逻辑
- `wenxing/`：桌面入口及部分模块化代码
- `Template_File/`：输入模板文件
- `tools/`：辅助脚本

## 部署说明

部署和运维文档见：

- `DEPLOY.md`
- `OPS.md`
