# -*- coding: utf-8 -*-
"""项目内置 Web 入口。"""

from api.main_runtime import app


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("api.main_runtime:app", host="0.0.0.0", port=8086)
