# -*- coding: utf-8 -*-
"""
使用统计模块 - 记录和查询每日使用次数
"""

import json
import logging
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

# 统计数据存储文件
STATS_DIR = Path(__file__).parent / "stats"
STATS_FILE = STATS_DIR / "usage_stats.json"

# 线程锁，确保并发安全（使用可重入锁避免嵌套调用死锁）
_stats_lock = threading.RLock()


def _ensure_stats_dir():
    """确保统计目录存在"""
    STATS_DIR.mkdir(exist_ok=True)


def _load_stats() -> Dict:
    """加载统计数据"""
    _ensure_stats_dir()
    if STATS_FILE.exists():
        try:
            with open(STATS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            logger.error(f"加载统计数据失败: {e}")
            return {"daily": {}, "total": 0}
    return {"daily": {}, "total": 0}


def _save_stats(stats: Dict):
    """保存统计数据"""
    _ensure_stats_dir()
    try:
        with open(STATS_FILE, 'w', encoding='utf-8') as f:
            json.dump(stats, f, ensure_ascii=False, indent=2)
    except IOError as e:
        logger.error(f"保存统计数据失败: {e}")


def record_usage(event_type: str = "analysis"):
    """
    记录一次使用
    
    Args:
        event_type: 事件类型，默认为 'analysis'（分析任务）
    """
    with _stats_lock:
        stats = _load_stats()
        today = datetime.now().strftime("%Y-%m-%d")
        
        # 初始化每日统计
        if today not in stats["daily"]:
            stats["daily"][today] = {"analysis": 0, "download": 0, "precheck": 0}
        
        # 兼容旧数据格式
        if isinstance(stats["daily"][today], int):
            old_count = stats["daily"][today]
            stats["daily"][today] = {"analysis": old_count, "download": 0, "precheck": 0}
        
        # 确保事件类型存在
        if event_type not in stats["daily"][today]:
            stats["daily"][today][event_type] = 0
        
        # 增加计数
        stats["daily"][today][event_type] += 1
        stats["total"] = stats.get("total", 0) + 1
        
        _save_stats(stats)
        logger.info(f"记录使用: {event_type}, 今日总计: {stats['daily'][today]}")


def get_today_stats() -> Dict:
    """
    获取今日使用统计
    
    Returns:
        包含今日各类事件计数的字典
    """
    with _stats_lock:
        stats = _load_stats()
        today = datetime.now().strftime("%Y-%m-%d")
        
        if today in stats["daily"]:
            daily_data = stats["daily"][today]
            # 兼容旧数据格式
            if isinstance(daily_data, int):
                return {
                    "date": today,
                    "analysis": daily_data,
                    "download": 0,
                    "precheck": 0,
                    "total": daily_data
                }
            total = sum(daily_data.values())
            return {
                "date": today,
                **daily_data,
                "total": total
            }
        
        return {
            "date": today,
            "analysis": 0,
            "download": 0,
            "precheck": 0,
            "total": 0
        }


def get_stats_summary(days: int = 7) -> Dict:
    """
    获取统计摘要
    
    Args:
        days: 查询最近多少天的数据，默认7天
        
    Returns:
        包含每日统计和总计的字典
    """
    with _stats_lock:
        stats = _load_stats()
        
        # 生成最近N天的日期列表
        date_list = []
        for i in range(days):
            date = (datetime.now() - timedelta(days=i)).strftime("%Y-%m-%d")
            date_list.append(date)
        
        # 收集每日数据
        daily_stats = []
        for date in reversed(date_list):  # 从旧到新排序
            if date in stats["daily"]:
                daily_data = stats["daily"][date]
                # 兼容旧数据格式
                if isinstance(daily_data, int):
                    daily_stats.append({
                        "date": date,
                        "analysis": daily_data,
                        "download": 0,
                        "precheck": 0,
                        "total": daily_data
                    })
                else:
                    total = sum(daily_data.values())
                    daily_stats.append({
                        "date": date,
                        **daily_data,
                        "total": total
                    })
            else:
                daily_stats.append({
                    "date": date,
                    "analysis": 0,
                    "download": 0,
                    "precheck": 0,
                    "total": 0
                })
        
        # 计算期间总计
        period_total = sum(d["total"] for d in daily_stats)
        period_analysis = sum(d.get("analysis", 0) for d in daily_stats)
        
        return {
            "period_days": days,
            "period_total": period_total,
            "period_analysis": period_analysis,
            "all_time_total": stats.get("total", 0),
            "daily": daily_stats,
            "today": get_today_stats()
        }


def cleanup_old_stats(keep_days: int = 90):
    """
    清理旧的统计数据
    
    Args:
        keep_days: 保留最近多少天的数据，默认90天
    """
    with _stats_lock:
        stats = _load_stats()
        cutoff_date = (datetime.now() - timedelta(days=keep_days)).strftime("%Y-%m-%d")
        
        old_dates = [date for date in stats["daily"].keys() if date < cutoff_date]
        for date in old_dates:
            del stats["daily"][date]
        
        if old_dates:
            _save_stats(stats)
            logger.info(f"清理了 {len(old_dates)} 天的旧统计数据")
