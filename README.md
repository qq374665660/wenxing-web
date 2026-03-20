# wenxing-web

地基基础分析及选型报告生成系统。

项目当前同时保留两套入口：

1. Web 入口：基于 FastAPI + Uvicorn，适合服务器部署和内网访问。
2. 桌面入口：适合本地单机分析和调试。

## 功能概览

- 解析建筑物、钻孔、地层和各孔分层等多工作表勘察数据
- 自动识别持力层、软弱下卧层和特殊地层工况
- 执行承载力、等效模量、沉降和变形相关分析
- 自动生成标准化地基基础分析报告
- 提供文件上传、任务执行、结果下载和健康检查接口

## 项目结构

```text
api/                Web API、静态页面和运行时相关代码
Template_File/      输入模板文件
tools/              辅助脚本
wenxing/            桌面端与模块化代码
wenxing2.py         核心分析与报告生成逻辑
main.py             桌面端启动入口
run_web.py          Web 服务启动入口
start_server.ps1    Windows 服务器启动脚本
```

## 运行环境

- Python 3.11+
- Windows / Windows Server
- IIS
  仅在反向代理部署 Web 服务时需要

## 安装依赖

```powershell
pip install -r requirements.txt
```

## 启动方式

### Web 服务

```powershell
python run_web.py
```

默认端口：

```text
http://127.0.0.1:8086
```

### Windows 服务器启动脚本

```powershell
.\start_server.ps1
```

或直接运行：

```text
start_server.bat
```

### 桌面端

```powershell
python main.py
```

## 说明

- `api/static/index.html` 为当前正式前端页面
- `api/main.py` 为 Web 入口
- `api/main_runtime.py` 为实际运行时主逻辑
- `wenxing2.py` 为核心分析算法入口

## 文档

- [部署文档](docs/deploy.md)
- [运维文档](docs/operations.md)
