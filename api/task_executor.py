# -*- coding: utf-8 -*-
"""
任务执行器 - 进程隔离版本

核心特性：
1. 每个分析任务在独立子进程中运行
2. 支持强制终止卡死的任务
3. 严格超时控制
4. 不会因单个任务卡死影响其他用户
"""

import multiprocessing
import logging
import time
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, Tuple

# 添加项目根目录到路径
_project_root = Path(__file__).parent.parent
sys.path.insert(0, str(_project_root))

logger = logging.getLogger(__name__)


def _run_analysis_in_process(
    input_path: str,
    output_path: str,
    params: dict,
    result_queue: multiprocessing.Queue
):
    """
    在子进程中执行分析任务
    这个函数会在独立进程中运行，与主进程完全隔离
    """
    try:
        # 在子进程中导入，避免主进程加载时的问题
        from wenxing2 import run_analysis
        
        # 创建参数回答函数
        def ask_yes_no(title, message):
            if "地下水位" in message or "持力层位置" in title:
                return params.get('water_level_above', True)
            elif "粉土" in title and "黏粒含量" in message:
                return params.get('silt_clay_content_ge_10', True)
            elif "粉质黏土" in title or "粉质粘土" in title:
                return params.get('silty_clay_e_il_ge_085', True)
            return True
        
        # 执行分析
        run_analysis(None, input_path, output_path, ask_yes_no)
        
        # 成功完成
        result_queue.put({"success": True, "error": None})
        
    except Exception as e:
        # 捕获所有异常
        result_queue.put({"success": False, "error": str(e)})


class TaskExecutor:
    """
    任务执行器
    
    使用独立进程执行每个任务，支持：
    - 强制终止超时任务
    - 任务状态跟踪
    - 资源清理
    """
    
    def __init__(self, timeout_seconds: int = 300):
        self.timeout_seconds = timeout_seconds
        self._active_processes: Dict[str, multiprocessing.Process] = {}
        self._process_start_times: Dict[str, datetime] = {}
        
    def execute(
        self,
        task_id: str,
        input_path: str,
        output_path: str,
        params: dict
    ) -> Tuple[bool, Optional[str]]:
        """
        执行分析任务
        
        Returns:
            (success: bool, error_message: Optional[str])
        """
        result_queue = multiprocessing.Queue()
        
        # 创建子进程
        process = multiprocessing.Process(
            target=_run_analysis_in_process,
            args=(input_path, output_path, params, result_queue),
            daemon=True  # 守护进程，主进程退出时自动终止
        )
        
        # 记录进程信息
        self._active_processes[task_id] = process
        self._process_start_times[task_id] = datetime.now()
        
        try:
            # 启动子进程
            process.start()
            logger.info(f"任务 {task_id} 已启动，进程 PID: {process.pid}")
            
            # 等待完成或超时
            process.join(timeout=self.timeout_seconds)
            
            if process.is_alive():
                # 超时，强制终止
                logger.warning(f"任务 {task_id} 超时 ({self.timeout_seconds}秒)，强制终止进程 {process.pid}")
                process.terminate()
                process.join(timeout=5)  # 等待最多5秒
                
                if process.is_alive():
                    # 如果还活着，强制杀死
                    logger.error(f"任务 {task_id} 无法正常终止，强制杀死")
                    process.kill()
                    process.join(timeout=2)
                
                return False, f"分析超时（超过{self.timeout_seconds}秒），请检查文件是否过大或格式是否正确"
            
            # 进程已结束，获取结果
            try:
                result = result_queue.get_nowait()
                if result["success"]:
                    logger.info(f"任务 {task_id} 成功完成")
                    return True, None
                else:
                    logger.error(f"任务 {task_id} 执行失败: {result['error']}")
                    return False, result["error"]
            except:
                # 进程退出但没有结果，可能是崩溃
                exit_code = process.exitcode
                if exit_code != 0:
                    return False, f"分析进程异常退出 (退出码: {exit_code})"
                return False, "分析进程异常退出"
                
        finally:
            # 清理进程记录
            self._active_processes.pop(task_id, None)
            self._process_start_times.pop(task_id, None)
            
            # 确保进程已终止
            if process.is_alive():
                process.kill()
    
    def get_active_count(self) -> int:
        """获取当前活跃任务数"""
        # 清理已结束的进程
        for task_id in list(self._active_processes.keys()):
            if not self._active_processes[task_id].is_alive():
                self._active_processes.pop(task_id, None)
                self._process_start_times.pop(task_id, None)
        return len(self._active_processes)
    
    def kill_task(self, task_id: str) -> bool:
        """强制终止指定任务"""
        process = self._active_processes.get(task_id)
        if process and process.is_alive():
            logger.warning(f"手动终止任务 {task_id}")
            process.terminate()
            process.join(timeout=5)
            if process.is_alive():
                process.kill()
            self._active_processes.pop(task_id, None)
            self._process_start_times.pop(task_id, None)
            return True
        return False
    
    def cleanup_zombie_tasks(self, max_age_seconds: int = 600) -> int:
        """
        清理僵尸任务（运行时间过长但未被正常处理的任务）
        
        Returns:
            清理的任务数量
        """
        cleaned = 0
        now = datetime.now()
        
        for task_id in list(self._active_processes.keys()):
            start_time = self._process_start_times.get(task_id)
            if start_time:
                age = (now - start_time).total_seconds()
                if age > max_age_seconds:
                    logger.warning(f"清理僵尸任务 {task_id}，已运行 {age:.0f} 秒")
                    self.kill_task(task_id)
                    cleaned += 1
        
        return cleaned


# 全局执行器实例
_executor: Optional[TaskExecutor] = None


def get_executor(timeout_seconds: int = 300) -> TaskExecutor:
    """获取全局任务执行器"""
    global _executor
    if _executor is None:
        _executor = TaskExecutor(timeout_seconds=timeout_seconds)
    return _executor


def execute_analysis(
    task_id: str,
    input_path: str,
    output_path: str,
    params: dict,
    timeout_seconds: int = 300
) -> Tuple[bool, Optional[str]]:
    """
    执行分析任务的便捷函数
    
    这是推荐的调用方式，会自动使用进程隔离
    """
    executor = get_executor(timeout_seconds)
    return executor.execute(task_id, input_path, output_path, params)
