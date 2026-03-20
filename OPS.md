# 成都地区地基基础分析系统 - 运维手册

> 适用环境：Windows Server 2016/2019/2022 + IIS + Python 3.9+
>
> 本手册面向日常维护人员，重点说明“怎么启动 / 监控 / 排错 / 调整参数”，不讲开发细节。

---

## 1. 系统总体概览

### 1.1 架构说明

- **前端**：静态页面
  - 目录：`api/static/index.html`
  - 由 IIS 直接提供静态文件访问
- **Web API**：FastAPI + Uvicorn
  - 入口：`api/main.py`（应用对象 `app`）
  - 启动脚本：`run_web.py`（默认监听 `0.0.0.0:8086`）
- **反向代理**：IIS
  - 站点根目录：`api/static`
  - 通过 `web.config` 将非静态请求转发到 `http://127.0.0.1:8086`
- **分析引擎**：`wenxing2.run_analysis`
  - 每个任务在**独立子进程**中执行，避免单个任务拖垮整个服务
- **持久数据**：
  - 上传文件：`api/uploads/`
  - 分析结果：`api/outputs/`
  - 日志：`logs/`
  - 使用统计：`api/stats/usage_stats.json`

### 1.2 主要端口与服务

- Uvicorn 服务端口：**8086**（仅本机访问，一般通过 IIS 反代）
- IIS 网站端口：通常为 **80 / 443**（对外暴露）
- 建议后台以 Windows 服务运行：
  - 服务名示例：`WenxingWebService`
  - 可通过 NSSM 或任务计划等方式创建

---

## 2. 目录结构说明（运维相关）

根目录示例：`C:\WebApps\wenxing-web`（以下路径均以此为根）

- `run_web.py`：Web 服务启动脚本
- `requirements.txt`：Python 依赖列表
- `logs/`
  - `service.log`：标准输出日志
  - `error.log`：错误输出 / 异常日志
- `api/`
  - `main.py`：FastAPI 入口，API 路由及任务管理参数
  - `task_executor.py`：子进程任务执行器（超时控制、强制终止）
  - `static/`：前端页面与 `web.config`
  - `uploads/`：上传的 Excel 文件（会被后台定时清理）
  - `outputs/`：生成的 Word 报告文件（会被后台定时清理）
  - `stats/usage_stats.json`：使用统计数据
- `Template_File/template_file.xlsx`：输入数据模板
- `venv/` 或 `.venv/`：Python 虚拟环境（可重建，一般不做备份）

---

## 3. 启动 / 停止 / 重启

### 3.1 作为 Windows 服务运行（推荐）

假设已使用 NSSM 创建服务名为 `WenxingWebService`：

```powershell
# 启动服务
Start-Service WenxingWebService

# 停止服务
Stop-Service WenxingWebService

# 重启服务
Restart-Service WenxingWebService

# 查看服务状态
Get-Service WenxingWebService
```

> 建议将服务启动类型设置为 **Automatic**，保证服务器重启后自动拉起：
>
> ```powershell
> Set-Service WenxingWebService -StartupType Automatic
> ```

### 3.2 临时手工运行（调试 / 单机使用）

```powershell
cd C:\WebApps\wenxing-web

# 激活虚拟环境
.\venv\Scripts\Activate.ps1

# 启动 Web 服务
python run_web.py
```

- 终端会显示访问地址：`http://localhost:8086`
- 按 `Ctrl + C` 可手动停止

> **注意**：生产环境不要同时运行“服务版”和“手工版”，否则端口冲突或出现两个实例。

---

## 4. 健康检查与监控

### 4.1 健康检查接口

- 地址：`GET /api/health`
- 示例返回：

```json
{
  "status": "ok",
  "message": "服务运行正常",
  "active_tasks": 0,
  "total_tasks": 12
}
```

关键字段说明：

- `status`：`ok` 为正常，其它视为异常
- `active_tasks`：当前正在处理或排队的任务数
- `total_tasks`：自服务启动以来创建的任务总数（24 小时内，过期任务会被清理）

可以在监控系统中每隔 1～5 分钟请求一次 `/api/health`：

- 若 **连续多次请求失败** 或 `status != "ok"`，可自动告警并尝试重启服务

### 4.2 使用统计接口（可选监控）

- 今日统计：`GET /api/stats/today`
- 近 N 天统计：`GET /api/stats/summary?days=7`（`days` 最大 90）

示例：

```json
{
  "date": "2024-12-11",
  "analysis": 5,
  "download": 3,
  "precheck": 7,
  "total": 15
}
```

可用于：

- 观察每日使用量
- 估算未来容量需求

---

## 5. 日志查看与排错

### 5.1 日志位置

- 目录：`C:\WebApps\wenxing-web\logs`
  - `service.log`：常规运行日志、访问时的一些信息
  - `error.log`：异常堆栈、错误信息

> 若通过 NSSM 配置了其它日志路径，请以 NSSM 配置为准。

### 5.2 常用查看命令

```powershell
cd C:\WebApps\wenxing-web

# 实时查看服务日志（类似 tail -f）
Get-Content .\logs\service.log -Wait

# 查看最近 200 行错误日志
Get-Content .\logs\error.log -Tail 200
```

### 5.3 日志轮转建议

当前项目未内置自动日志切割，建议运维侧：

- 使用计划任务定期（例如每天/每周）执行：
  - 将 `service.log` / `error.log` 复制为带日期的归档文件
  - 清空原日志文件
- 或使用第三方日志采集工具（如 Filebeat）上传到集中日志平台

示例（简单按日期归档）：

```powershell
$LogDir = "C:\WebApps\wenxing-web\logs"
$Date = Get-Date -Format "yyyyMMdd"

Copy-Item "$LogDir\service.log" "$LogDir\service_$Date.log" -ErrorAction SilentlyContinue
Copy-Item "$LogDir\error.log"   "$LogDir\error_$Date.log"   -ErrorAction SilentlyContinue

Clear-Content "$LogDir\service.log" -ErrorAction SilentlyContinue
Clear-Content "$LogDir\error.log"   -ErrorAction SilentlyContinue
```

---

## 6. 任务与资源管理配置

核心配置在 `api/main.py`：

```python
# 任务管理
MAX_ACTIVE_TASKS = 30          # 最大同时处理的任务数
MAX_QUEUE_SIZE = 50            # 最大排队任务数
TASK_TTL_SECONDS = 24 * 3600   # 任务记录保留时间 (24 小时)

# 文件限制
MAX_FILE_SIZE_MB = 30          # 最大上传文件大小 (MB)

# 超时设置
ANALYSIS_TIMEOUT_SECONDS = 300  # 单个任务最大执行时间 (5 分钟)
SOFT_TIMEOUT_SECONDS = 240      # 软超时警告时间 (4 分钟)

# 清理设置
CLEANUP_INTERVAL_SECONDS = 1800  # 定时清理间隔 (30 分钟)
ZOMBIE_CHECK_INTERVAL = 60       # 僵尸任务检查间隔 (1 分钟)
```

### 6.1 并发与性能

- **MAX_ACTIVE_TASKS**：
  - 当前设置为 **30**，表示同时最多 30 个分析任务在执行/排队
  - 若达到上限，新请求会返回 `503 当前分析任务过多，请稍后再试`
  - 调整方法：修改 `api/main.py` 中该值，重启服务生效
  - 增加该值前，请评估服务器 CPU / 内存

- **任务执行超时**：
  - 在 `api/task_executor.py` 中，子进程执行默认超时与 `ANALYSIS_TIMEOUT_SECONDS` 一致（300 秒）
  - 超时会强制终止子进程，返回“分析超时”错误，**不会卡死整个服务**

### 6.2 上传与磁盘空间

- 上传文件大小由 `MAX_FILE_SIZE_MB` 控制，超出直接拒绝
- 上传目录：`api/uploads/`
- 输出目录：`api/outputs/`
- 系统会定期根据 `TASK_TTL_SECONDS` + `CLEANUP_INTERVAL_SECONDS` 自动删除：
  - 超过保留时间的历史任务记录
  - 没有对应任务的“孤立”上传 / 输出文件

> 运维仍需定期关注磁盘空间：
>
> - 检查 `uploads` / `outputs` 目录是否异常增大
> - 必要时可以手动删除过旧文件（建议先停服务或确认无用户在用）

### 6.3 速率限制（防止恶意刷接口）

在 `api/main.py` 中使用 `slowapi` 对部分接口做了限流：

- `/api/upload`：`10/minute`
- `/api/precheck`：`120/minute`
- `/api/health`：`60/minute`
- `/api/stats/today`：`60/minute`
- `/api/stats/summary`：`30/minute`

如需调整，需要开发修改对应装饰器中的限流策略，然后重启服务。

---

## 7. 存储与备份

### 7.1 建议备份范围

定期（如每日 / 每周）备份以下内容：

- **代码与配置**：
  - 整个 `wenxing-web` 目录（排除 `venv/`、`.venv/` 等可重建环境）
- **模板与静态资源**：
  - `Template_File/`
  - `api/static/`
- **日志（可选）**：
  - `logs/` 下的归档日志
- **使用统计（可选）**：
  - `api/stats/usage_stats.json`

### 7.2 备份注意事项

- 备份前，尽量在**业务低谷时段**执行
- 无需特别停服务，但如果对一致性要求高，可以先暂时停止服务再备份
- 虚拟环境 `venv/` 可以不备份，恢复时重新执行 `pip install -r requirements.txt` 即可

---

## 8. 升级与回滚流程

### 8.1 升级步骤（建议）

1. **提前备份** 当前版本（参考第 7 章）
2. 停止服务：
   ```powershell
   Stop-Service WenxingWebService
   ```
3. 替换应用代码：
   - 从新版本包覆盖旧版本代码
   - 如有 `requirements.txt` 更新，则重新安装依赖：
     ```powershell
     .\venv\Scripts\Activate.ps1
     pip install -r requirements.txt
     ```
4. 启动服务：
   ```powershell
   Start-Service WenxingWebService
   ```
5. 验证：
   - 访问前端页面
   - 调用 `/api/health`，确认 `status = ok`
   - 做一次完整上传 → 分析 → 下载流程回归测试

### 8.2 回滚步骤

若升级后出现严重问题，可按以下方式回滚：

1. 停止服务：
   ```powershell
   Stop-Service WenxingWebService
   ```
2. 用备份包覆盖当前目录（代码 / 静态文件等）
3. 如依赖版本有变更，可重新安装旧版本依赖
4. 启动服务并验证

---

## 9. 常见问题排查

### 9.1 前端访问 502 / 500 / 打不开

1. 检查后端服务：
   ```powershell
   Get-Service WenxingWebService
   ```
   - 若未运行：`Start-Service WenxingWebService`
2. 在服务器本机测试：
   - 打开浏览器访问 `http://127.0.0.1:8086/api/health`
   - 若无法访问，查看 `logs/error.log`
3. 检查 IIS：
   - 确认网站已启动
   - 检查 `api/static/web.config` 中反向代理地址是否为 `http://127.0.0.1:8086/{R:1}`

### 9.2 上传 Excel 报错

- **提示不支持的文件类型**：
  - 只允许 `.xlsx` / `.xls`
- **提示文件过大**：
  - 检查报错信息中显示的文件大小
  - 如确需上传大文件，可适度调高 `MAX_FILE_SIZE_MB`，并注意磁盘和内存
- **提示“当前分析任务过多，请稍后再试”**：
  - 已达到 `MAX_ACTIVE_TASKS` 上限
  - 可等待或扩容服务器 / 降低单次使用文件规模

### 9.3 分析时间过长或提示超时

- 单个任务最大执行时间为 `ANALYSIS_TIMEOUT_SECONDS`（默认 300 秒）
- 若频繁超时：
  - 检查上传文件是否异常大或数据异常
  - 适当提高超时时间前，需评估服务器资源

### 9.4 使用统计弹窗一直显示“加载中”

1. 检查 `api/stats/usage_stats.json` 是否存在、权限是否正常
2. 查看浏览器控制台 / 网络请求：
   - `/api/stats/today`、`/api/stats/summary` 是否返回 200
3. 检查服务端日志中是否有统计相关错误

---

## 10. 安全与权限建议

- 仅开放必要端口（80/443），屏蔽 8086 对外访问
- 为应用目录授予 **最小权限**：
  - IIS 或运行服务的账户需对 `api/uploads`、`api/outputs`、`logs` 具有写权限
  - 其余目录只需读权限
- 如部署 HTTPS：
  - 在 IIS 中绑定证书，并配置 HTTP → HTTPS 重定向
- 不在代码中硬编码敏感信息（本系统目前无数据库 / 无密钥配置）

---

## 11. 快速检查清单（运维日常）

- [ ] 服务 `WenxingWebService` 处于 Running 状态
- [ ] `/api/health` 返回 `status = ok`
- [ ] 近期无大量 `error.log` 异常堆栈
- [ ] 磁盘空间充足，`uploads` / `outputs` 没有异常膨胀
- [ ] 日志有定期归档
- [ ] IIS 网站运行正常，证书未过期（如使用 HTTPS）

如需扩展更多运维规范（如接入 Zabbix / Prometheus / 日志平台等），可以在此文档基础上继续补充。
