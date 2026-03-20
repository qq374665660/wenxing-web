# -*- coding: utf-8 -*-
"""
并发压测脚本：模拟 N 个用户同时使用完整流程
- POST /api/upload 发送 Excel
- 轮询 /api/status/{task_id}
- GET /api/download/{task_id}

用法示例：
  D:\wenxing-web\wenxing-web\venv\Scripts\python.exe tools\load_test.py \
    --base-url http://your-domain-or-ip:port \
    --users 30 \
    --file "C:\\path\\to\\sample.xlsx" \
    --timeout 600

注意：该脚本需要 httpx
  D:\wenxing-web\wenxing-web\venv\Scripts\python.exe -m pip install httpx
"""

import asyncio
import argparse
import os
import sys
import time
import random
import statistics as stats
from pathlib import Path
from typing import Dict, Any, List

try:
    import httpx
except Exception as e:
    print("缺少 httpx，请先安装：venv Python 执行 'python -m pip install httpx'", file=sys.stderr)
    raise

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_UPLOADS_DIR = PROJECT_ROOT / "api" / "uploads"


def pick_default_file() -> Path:
    """优先从 api/uploads 中选择最近的 xlsx 作为样例。"""
    if DEFAULT_UPLOADS_DIR.exists():
        cands = sorted(DEFAULT_UPLOADS_DIR.glob("*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True)
        if cands:
            return cands[0]
    return None


async def run_user(idx: int, client: httpx.AsyncClient, base_url: str, file_bytes: bytes, file_name: str, 
                   timeout_s: int) -> Dict[str, Any]:
    t0 = time.perf_counter()
    result = {"user": idx, "status": "ok", "error": None, "duration": None}
    try:
        # 随机参数
        water = random.choice(["yes", "no"])
        silt = random.choice(["yes", "no"])
        silty = random.choice(["yes", "no"])

        files = {
            "file": (file_name, file_bytes, 
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        }
        data = {
            "water_level_above": "true" if water == "yes" else "false",
            "silt_clay_content_ge_10": "true" if silt == "yes" else "false",
            "silty_clay_e_il_ge_085": "true" if silty == "yes" else "false",
        }

        # 上传
        r = await client.post(f"{base_url}/api/upload", files=files, data=data, timeout=timeout_s)
        if r.status_code != 200:
            result["status"] = "upload_error"
            result["error"] = f"HTTP {r.status_code}: {r.text[:200]}"
            return result
        task_id = r.json().get("task_id")
        if not task_id:
            result["status"] = "upload_no_task"
            result["error"] = f"no task_id in response: {r.text[:200]}"
            return result

        # 轮询
        start_poll = time.perf_counter()
        while True:
            s = await client.get(f"{base_url}/api/status/{task_id}", timeout=timeout_s)
            if s.status_code != 200:
                result["status"] = "status_error"
                result["error"] = f"HTTP {s.status_code}: {s.text[:200]}"
                return result
            js = s.json()
            st = js.get("status")
            if st == "completed":
                # 下载验证
                d = await client.get(f"{base_url}/api/download/{task_id}", timeout=timeout_s)
                if d.status_code != 200:
                    result["status"] = "download_error"
                    result["error"] = f"HTTP {d.status_code}"
                else:
                    result["status"] = "ok"
                break
            elif st == "failed":
                result["status"] = "task_failed"
                result["error"] = js.get("message")
                break
            else:
                await asyncio.sleep(1)
                # 简单超时保护
                if time.perf_counter() - start_poll > timeout_s:
                    result["status"] = "timeout"
                    result["error"] = f"wait status timeout {timeout_s}s"
                    break

    except Exception as e:
        result["status"] = "exception"
        result["error"] = repr(e)
    finally:
        result["duration"] = round(time.perf_counter() - t0, 3)
    return result


async def main():
    parser = argparse.ArgumentParser(description="并发压测：上传->轮询->下载 全流程")
    parser.add_argument("--base-url", required=True, help="对外访问地址，例如 http://your-domain:1026")
    parser.add_argument("--users", type=int, default=30, help="并发用户数")
    parser.add_argument("--file", type=str, default="", help="Excel 样例文件路径（可留空，自动在 api/uploads 挑选最近一个）")
    parser.add_argument("--timeout", type=int, default=600, help="单用户最长等待秒数")
    parser.add_argument("--stagger", type=float, default=0.2, help="虚拟用户启动间隔秒，0 表示同时发起")
    args = parser.parse_args()

    base_url = args.base_url.rstrip("/")

    # 读取文件
    fpath = Path(args.file) if args.file else pick_default_file()
    if not fpath or not fpath.exists():
        print("未找到样例 Excel，请通过 --file 指定一个有效路径", file=sys.stderr)
        print(f"可参考目录: {DEFAULT_UPLOADS_DIR}", file=sys.stderr)
        sys.exit(2)

    file_bytes = fpath.read_bytes()
    file_name = fpath.name
    print(f"使用样例文件: {fpath}")

    limits = httpx.Limits(max_keepalive_connections=args.users, max_connections=args.users)
    async with httpx.AsyncClient(limits=limits, verify=False) as client:
        tasks: List[asyncio.Task] = []
        t0 = time.perf_counter()
        for i in range(args.users):
            tasks.append(asyncio.create_task(run_user(i + 1, client, base_url, file_bytes, file_name, args.timeout)))
            if args.stagger > 0 and i < args.users - 1:
                await asyncio.sleep(args.stagger)
        results = await asyncio.gather(*tasks)
        total = len(results)
        ok = sum(1 for r in results if r["status"] == "ok")
        errors = {}
        durations = []
        for r in results:
            durations.append(r["duration"]) if r.get("duration") else None
            if r["status"] != "ok":
                errors[r["status"]] = errors.get(r["status"], 0) + 1
        elapsed = round(time.perf_counter() - t0, 3)

        print("\n====== 压测结果 ======")
        print(f"总用户: {total}, 成功: {ok}, 失败: {total - ok}")
        if errors:
            print("错误分布:")
            for k, v in errors.items():
                print(f"  - {k}: {v}")
        if durations:
            durations.sort()
            p50 = durations[int(0.50 * len(durations))]
            p95 = durations[int(0.95 * len(durations)) - 1]
            p99 = durations[int(0.99 * len(durations)) - 1] if len(durations) >= 100 else durations[-1]
            print(f"整体耗时: {elapsed}s  | 每用户耗时(P50/P95/P99): {p50}/{p95}/{p99} s")

        # 输出前 5 个失败样本
        bad = [r for r in results if r["status"] != "ok"][:5]
        if bad:
            print("\n失败示例(最多5条):")
            for r in bad:
                print(f"  user#{r['user']}: {r['status']} | {r['error']}")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        pass
