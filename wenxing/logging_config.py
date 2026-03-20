# -*- coding: utf-8 -*-
"""
日志配置模块
"""

import logging

# 日志配置
LOG_FILE = 'foundation_analysis.log'
LOG_LEVEL = logging.DEBUG
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'


def setup_logging():
    """设置日志配置"""
    logging.basicConfig(
        filename=LOG_FILE,
        level=LOG_LEVEL,
        format=LOG_FORMAT,
        encoding='utf-8'
    )
    return logging.getLogger(__name__)


# 获取全局 logger
logger = logging.getLogger('wenxing')
