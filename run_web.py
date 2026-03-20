# -*- coding: utf-8 -*-
"""
Web 服务启动脚本
"""

import uvicorn

if __name__ == "__main__":
    print("=" * 50)
    print("成都地区地基基础分析系统 - Web 服务")
    print("=" * 50)
    print("访问地址: http://localhost:8086")
    print("按 Ctrl+C 停止服务")
    print("=" * 50)
    
    uvicorn.run(
        "api.main:app",
        host="0.0.0.0",
        port=8086,
        reload=False  # 生产环境关闭自动重载
    )
