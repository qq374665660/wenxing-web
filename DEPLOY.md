# Windows Server 部署教程

## 成都地区地基基础分析系统 - Web 版部署指南

本文档详细介绍如何在 Windows Server 上部署本系统。

---

## 目录

1. [环境准备](#1-环境准备)
2. [安装 Python](#2-安装-python)
3. [部署项目](#3-部署项目)
4. [配置 IIS 反向代理](#4-配置-iis-反向代理)
5. [配置 Windows 服务](#5-配置-windows-服务)
6. [防火墙设置](#6-防火墙设置)
7. [常见问题](#7-常见问题)

---

## 1. 环境准备

### 1.1 服务器要求

- **操作系统**: Windows Server 2016/2019/2022
- **内存**: 最低 4GB，推荐 8GB+
- **磁盘**: 至少 10GB 可用空间
- **网络**: 开放 80 端口（HTTP）或 443 端口（HTTPS）

### 1.2 需要安装的软件

- Python 3.9+
- IIS（可选，用于反向代理）
- NSSM（用于创建 Windows 服务）

---

## 2. 安装 Python

### 2.1 下载 Python

1. 访问 [Python 官网](https://www.python.org/downloads/windows/)
2. 下载 Python 3.11.x 或更高版本的 Windows installer (64-bit)

### 2.2 安装 Python

1. 运行安装程序
2. **重要**: 勾选 "Add Python to PATH"
3. 选择 "Customize installation"
4. 勾选所有可选功能
5. 安装位置建议: `C:\Python311`
6. 完成安装

### 2.3 验证安装

打开 PowerShell，执行：

```powershell
python --version
pip --version
```

---

## 3. 部署项目

### 3.1 上传项目文件

将项目文件夹上传到服务器，例如：`C:\WebApps\wenxing-web`

### 3.2 创建虚拟环境

```powershell
cd C:\WebApps\wenxing-web
python -m venv venv
```

### 3.3 激活虚拟环境

```powershell
.\venv\Scripts\Activate.ps1
```

如果遇到执行策略错误，先执行：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3.4 安装依赖

```powershell
# 使用国内镜像加速（推荐）
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 或使用阿里云镜像
pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple
```

**配置永久镜像（可选）：**
```powershell
pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
```

### 3.5 测试运行

```powershell
python run_web.py
```

访问 `http://localhost:8000` 确认服务正常运行。

---

## 4. 配置 IIS 反向代理

### 4.1 安装 IIS

1. 打开 "服务器管理器"
2. 点击 "添加角色和功能"
3. 选择 "Web 服务器 (IIS)"
4. 完成安装

### 4.2 安装 URL Rewrite 和 ARR

1. 下载并安装 [URL Rewrite](https://www.iis.net/downloads/microsoft/url-rewrite)
2. 下载并安装 [Application Request Routing (ARR)](https://www.iis.net/downloads/microsoft/application-request-routing)

### 4.3 启用 ARR 代理

1. 打开 IIS 管理器
2. 选择服务器节点
3. 双击 "Application Request Routing Cache"
4. 点击右侧 "Server Proxy Settings"
5. 勾选 "Enable proxy"
6. 点击 "应用"

### 4.4 创建网站

1. 在 IIS 管理器中，右键 "网站" → "添加网站"
2. 网站名称: `WenxingWeb`
3. 物理路径: `C:\WebApps\wenxing-web\api\static`
4. 端口: `80`（或其他端口）
5. 主机名: 填写您的域名（可选）

### 4.5 配置 URL 重写规则

在网站目录下创建 `web.config` 文件：

```xml
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <rewrite>
            <rules>
                <rule name="ReverseProxyToUvicorn" stopProcessing="true">
                    <match url="(.*)" />
                    <conditions>
                        <add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" />
                    </conditions>
                    <action type="Rewrite" url="http://localhost:8000/{R:1}" />
                </rule>
            </rules>
        </rewrite>
    </system.webServer>
</configuration>
```

---

## 5. 配置 Windows 服务

使用 NSSM (Non-Sucking Service Manager) 将应用注册为 Windows 服务。

### 5.1 下载 NSSM

1. 访问 [NSSM 官网](https://nssm.cc/download)
2. 下载最新版本
3. 解压到 `C:\Tools\nssm`

### 5.2 创建服务

以管理员身份打开 PowerShell：

```powershell
C:\Tools\nssm\win64\nssm.exe install WenxingWebService
```

在弹出的界面中配置：

- **Path**: `C:\WebApps\wenxing-web\venv\Scripts\python.exe`
- **Startup directory**: `C:\WebApps\wenxing-web`
- **Arguments**: `run_web.py`

切换到 "Details" 标签：
- **Display name**: 成都地区地基基础分析系统
- **Description**: Foundation Analysis Web Service

切换到 "I/O" 标签（可选，用于日志）：
- **Output (stdout)**: `C:\WebApps\wenxing-web\logs\service.log`
- **Error (stderr)**: `C:\WebApps\wenxing-web\logs\error.log`

点击 "Install service"。

### 5.3 启动服务

```powershell
# 创建日志目录
New-Item -ItemType Directory -Path "C:\WebApps\wenxing-web\logs" -Force

# 启动服务
Start-Service WenxingWebService

# 设置开机自启
Set-Service WenxingWebService -StartupType Automatic
```

### 5.4 服务管理命令

```powershell
# 查看服务状态
Get-Service WenxingWebService

# 停止服务
Stop-Service WenxingWebService

# 重启服务
Restart-Service WenxingWebService

# 卸载服务
C:\Tools\nssm\win64\nssm.exe remove WenxingWebService confirm
```

---

## 6. 防火墙设置

### 6.1 开放端口

以管理员身份打开 PowerShell：

```powershell
# 开放 80 端口（HTTP）
New-NetFirewallRule -DisplayName "HTTP Port 80" -Direction Inbound -Protocol TCP -LocalPort 80 -Action Allow

# 开放 443 端口（HTTPS，如果需要）
New-NetFirewallRule -DisplayName "HTTPS Port 443" -Direction Inbound -Protocol TCP -LocalPort 443 -Action Allow

# 开放 8000 端口（如果直接访问 uvicorn）
New-NetFirewallRule -DisplayName "Uvicorn Port 8000" -Direction Inbound -Protocol TCP -LocalPort 8000 -Action Allow
```

---

## 7. 常见问题

### Q1: 服务启动失败

**检查步骤：**
1. 查看日志文件 `C:\WebApps\wenxing-web\logs\error.log`
2. 确认 Python 路径正确
3. 确认所有依赖已安装

### Q2: 上传文件失败

**解决方案：**
1. 确保 `api/uploads` 和 `api/outputs` 目录有写入权限
2. 检查 IIS 应用程序池的标识是否有权限

```powershell
# 授予 IIS 用户写入权限
icacls "C:\WebApps\wenxing-web\api\uploads" /grant "IIS_IUSRS:(OI)(CI)F"
icacls "C:\WebApps\wenxing-web\api\outputs" /grant "IIS_IUSRS:(OI)(CI)F"
```

### Q3: 502 Bad Gateway 错误

**解决方案：**
1. 确认后端服务正在运行: `Get-Service WenxingWebService`
2. 确认端口 8000 没有被占用
3. 检查 IIS ARR 代理是否启用

### Q4: 如何配置 HTTPS？

**步骤：**
1. 获取 SSL 证书（可使用 Let's Encrypt 免费证书）
2. 在 IIS 中导入证书
3. 修改网站绑定，添加 HTTPS 绑定
4. 配置 HTTP 重定向到 HTTPS

---

## 快速部署脚本

将以下内容保存为 `deploy.ps1`，以管理员身份运行：

```powershell
# 快速部署脚本
$AppPath = "C:\WebApps\wenxing-web"

# 1. 创建虚拟环境
Set-Location $AppPath
python -m venv venv

# 2. 激活并安装依赖
& "$AppPath\venv\Scripts\pip.exe" install -r requirements.txt

# 3. 创建必要目录
New-Item -ItemType Directory -Path "$AppPath\api\uploads" -Force
New-Item -ItemType Directory -Path "$AppPath\api\outputs" -Force
New-Item -ItemType Directory -Path "$AppPath\logs" -Force

# 4. 设置权限
icacls "$AppPath\api\uploads" /grant "Everyone:(OI)(CI)F"
icacls "$AppPath\api\outputs" /grant "Everyone:(OI)(CI)F"
icacls "$AppPath\logs" /grant "Everyone:(OI)(CI)F"

Write-Host "部署完成！请手动配置 NSSM 服务。" -ForegroundColor Green
```

---

## 联系方式

如有问题，请联系：中建西勘院 文兴

---

*最后更新: 2024年*
