# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

import time
import warnings
from typing import Any, Dict, List, Optional, Tuple, Union
from dataclasses import dataclass, field

import pyautogui
import pywinauto
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.win32structures import RECT

from ufo.config.config import Config
from ufo.utils import print_with_color

configs = Config.get_instance().config_data


@dataclass
class ClickAttempt:
    """记录单次点击尝试的信息"""
    x: int
    y: int
    timestamp: float
    success: bool = False
    error_message: str = ""


@dataclass
class SmartClickConfig:
    """智能点击重试配置"""
    max_retries: int = 24  # 8个方向 * 3个距离层级
    retry_offsets: List[int] = field(default_factory=lambda: [5, 10, 15])
    retry_directions: List[Tuple[int, int]] = field(default_factory=lambda: [
        (0, -1),   # 上
        (1, -1),   # 右上
        (1, 0),    # 右
        (1, 1),    # 右下
        (0, 1),    # 下
        (-1, 1),   # 左下
        (-1, 0),   # 左
        (-1, -1),  # 左上
    ])
    wait_between_clicks: float = 0.2  # 点击间隔时间
    focus_change_timeout: float = 1.0  # 焦点变化检测超时时间
    element_stability_check_count: int = 3  # 连续失败检查次数


class FocusChangeDetector:
    """焦点变化检测器"""
    
    def __init__(self):
        self.initial_window_handle = None
        self.initial_focus_element = None
    
    def capture_initial_state(self, application_window: UIAWrapper) -> None:
        """捕获初始焦点状态"""
        try:
            self.initial_window_handle = application_window.handle
            # 尝试获取当前焦点元素
            try:
                from pywinauto import Desktop
                desktop = Desktop(backend="uia")
                self.initial_focus_element = desktop.get_focus()
            except Exception:
                self.initial_focus_element = None
        except Exception as e:
            print_with_color(f"Failed to capture initial focus state: {e}", "yellow")
    
    def has_focus_changed(self, application_window: UIAWrapper) -> bool:
        """检测焦点是否发生变化"""
        try:
            # 检查窗口句柄是否变化
            current_handle = application_window.handle
            if current_handle != self.initial_window_handle:
                return True
            
            # 检查焦点元素是否变化
            try:
                from pywinauto import Desktop
                desktop = Desktop(backend="uia")
                current_focus = desktop.get_focus()
                
                if self.initial_focus_element is None and current_focus is not None:
                    return True
                elif self.initial_focus_element is not None and current_focus is None:
                    return True
                elif (self.initial_focus_element is not None and 
                    current_focus is not None and 
                    self.initial_focus_element != current_focus):
                    return True
            except Exception:
                # 如果无法获取焦点信息，假设没有变化
                pass
                
            return False
        except Exception as e:
            print_with_color(f"Error detecting focus change: {e}", "yellow")
            return False


class ElementStabilityChecker:
    """UI元素稳定性检查器"""
    
    def __init__(self, control: UIAWrapper):
        self.control = control
        self.original_rect = None
        if control:
            try:
                self.original_rect = control.rectangle()
            except Exception:
                self.original_rect = None
    
    def is_element_stable(self) -> bool:
        """检查元素是否仍然稳定（在UI树中且矩形位置未变）"""
        if not self.control or not self.original_rect:
            return False
        
        try:
            # 检查元素是否仍然可访问
            current_rect = self.control.rectangle()
            
            # 检查矩形是否相同
            if (current_rect.left == self.original_rect.left and
                current_rect.top == self.original_rect.top and
                current_rect.right == self.original_rect.right and
                current_rect.bottom == self.original_rect.bottom):
                return True
            
            return False
        except Exception:
            # 如果无法访问元素，说明元素已不稳定
            return False


class SmartClickController:
    """智能点击控制器 - 实现网格化重试点击机制"""
    
    def __init__(self, config: Optional[SmartClickConfig] = None):
        self.config = config or SmartClickConfig()
        self.click_attempts: List[ClickAttempt] = []
        self.focus_detector = FocusChangeDetector()
        self.element_checker: Optional[ElementStabilityChecker] = None
    
    def smart_click_with_retry(
        self,
        original_click_func,
        params: Dict[str, Any],
        control: Optional[UIAWrapper],
        application_window: UIAWrapper,
        target_coordinates: Optional[Tuple[int, int]] = None
    ) -> str:
        """
        执行智能点击重试
        
        :param original_click_func: 原始点击函数
        :param params: 点击参数
        :param control: 目标控件
        :param application_window: 应用程序窗口
        :param target_coordinates: 目标坐标 (绝对坐标)
        :return: 点击结果
        """
        # 重置状态
        self.click_attempts.clear()
        self.focus_detector.capture_initial_state(application_window)
        self.element_checker = ElementStabilityChecker(control) if control else None
        
        print_with_color("Starting smart click with retry mechanism...", "cyan")
        
        # 首次尝试原始点击
        result = self._attempt_click(original_click_func, params, 0, 0)
        if result["success"]:
            print_with_color("Original click succeeded", "green")
            return result["message"]
        
        # 如果原始点击失败，开始网格重试
        if not target_coordinates:
            # 尝试从参数中提取坐标
            target_coordinates = self._extract_coordinates_from_params(params, application_window)
        
        if not target_coordinates:
            print_with_color("Cannot extract target coordinates for retry", "red")
            return result["message"]
        
        return self._perform_grid_retry(
            original_click_func, params, target_coordinates, application_window
        )
    
    def _extract_coordinates_from_params(
        self, params: Dict[str, Any], application_window: UIAWrapper
    ) -> Optional[Tuple[int, int]]:
        """从参数中提取目标坐标"""
        try:
            # 处理相对坐标
            if "x" in params and "y" in params:
                x = float(params["x"])
                y = float(params["y"])
                
                # 转换为绝对坐标
                app_rect = application_window.rectangle()
                abs_x = int(app_rect.left + x * app_rect.width())
                abs_y = int(app_rect.top + y * app_rect.height())
                
                return (abs_x, abs_y)
        except Exception as e:
            print_with_color(f"Failed to extract coordinates: {e}", "yellow")
        
        return None
    
    def _perform_grid_retry(
        self,
        original_click_func,
        params: Dict[str, Any],
        center_coords: Tuple[int, int],
        application_window: UIAWrapper
    ) -> str:
        """执行网格化重试点击"""
        center_x, center_y = center_coords
        consecutive_failures = 0
        
        print_with_color(f"Starting grid retry around center ({center_x}, {center_y})", "cyan")
        
        # 按距离层级进行重试
        for distance in self.config.retry_offsets:
            print_with_color(f"Trying offset distance: ±{distance}px", "cyan")
            
            # 在当前距离的8个方向上尝试
            for direction_x, direction_y in self.config.retry_directions:
                if len(self.click_attempts) >= self.config.max_retries:
                    break
                
                # 计算偏移坐标
                offset_x = center_x + (direction_x * distance)
                offset_y = center_y + (direction_y * distance)
                
                # 更新参数中的坐标
                retry_params = self._update_params_with_coordinates(
                    params, offset_x, offset_y, application_window
                )
                
                # 尝试点击
                result = self._attempt_click(
                    original_click_func, retry_params, offset_x, offset_y
                )
                
                if result["success"]:
                    print_with_color(
                        f"Smart retry succeeded at offset ({direction_x * distance}, {direction_y * distance})",
                        "green"
                    )
                    return result["message"]
                
                consecutive_failures += 1
                
                # 检查连续失败和元素稳定性
                if consecutive_failures >= self.config.element_stability_check_count:
                    if not self._should_continue_retry():
                        print_with_color(
                            "Element is no longer stable, stopping retry", "yellow"
                        )
                        break
                    consecutive_failures = 0
                
                # 等待下一次重试
                time.sleep(self.config.wait_between_clicks)
            
            if len(self.click_attempts) >= self.config.max_retries:
                break
        
        # 所有重试都失败
        print_with_color(f"Smart click retry exhausted after {len(self.click_attempts)} attempts", "red")
        return f"Click failed after {len(self.click_attempts)} retry attempts"
    
    def _update_params_with_coordinates(
        self, 
        original_params: Dict[str, Any], 
        abs_x: int, 
        abs_y: int, 
        application_window: UIAWrapper
    ) -> Dict[str, Any]:
        """更新参数中的坐标为新的偏移坐标"""
        params = original_params.copy()
        
        try:
            # 转换绝对坐标为相对坐标
            app_rect = application_window.rectangle()
            rel_x = (abs_x - app_rect.left) / app_rect.width()
            rel_y = (abs_y - app_rect.top) / app_rect.height()
            
            params["x"] = rel_x
            params["y"] = rel_y
            
        except Exception as e:
            print_with_color(f"Failed to update coordinates in params: {e}", "yellow")
        
        return params
    
    def _attempt_click(
        self, click_func, params: Dict[str, Any], x: int, y: int
    ) -> Dict[str, Any]:
        """尝试执行单次点击"""
        attempt = ClickAttempt(x=x, y=y, timestamp=time.time())
        
        try:
            # 执行点击
            result = click_func(params)
            
            # 等待并检查焦点变化
            time.sleep(self.config.wait_between_clicks)
            
            # 检查是否有焦点变化（表示点击成功）
            if self._check_click_success():
                attempt.success = True
                attempt.error_message = "Success"
                self.click_attempts.append(attempt)
                return {"success": True, "message": result}
            else:
                attempt.success = False
                attempt.error_message = "No focus change detected"
        
        except Exception as e:
            attempt.success = False
            attempt.error_message = str(e)
            print_with_color(f"Click attempt failed: {e}", "yellow")
        
        self.click_attempts.append(attempt)
        return {"success": False, "message": attempt.error_message}
    
    def _check_click_success(self) -> bool:
        """检查点击是否成功（通过焦点变化判断）"""
        # 给系统一些时间响应
        start_time = time.time()
        while time.time() - start_time < self.config.focus_change_timeout:
            if self.focus_detector.has_focus_changed(None):  # TODO: 传入正确的window
                return True
            time.sleep(0.05)  # 短暂等待后再次检查
        
        return False
    
    def _should_continue_retry(self) -> bool:
        """判断是否应该继续重试"""
        if self.element_checker:
            return self.element_checker.is_element_stable()
        return True
    
    def get_retry_statistics(self) -> Dict[str, Any]:
        """获取重试统计信息"""
        total_attempts = len(self.click_attempts)
        successful_attempts = sum(1 for attempt in self.click_attempts if attempt.success)
        
        return {
            "total_attempts": total_attempts,
            "successful_attempts": successful_attempts,
            "failure_rate": (total_attempts - successful_attempts) / max(total_attempts, 1),
            "attempts_details": [
                {
                    "x": attempt.x,
                    "y": attempt.y,
                    "success": attempt.success,
                    "error": attempt.error_message,
                    "timestamp": attempt.timestamp
                }
                for attempt in self.click_attempts
            ]
        }


def handle_element_not_enabled_exception(exception: Exception) -> bool:
    """
    检查异常是否为ElementNotEnabled类型
    
    :param exception: 捕获的异常
    :return: 如果是ElementNotEnabled异常则返回True
    """
    error_message = str(exception).lower()
    return any(keyword in error_message for keyword in [
        "elementnotenabled", 
        "element not enabled", 
        "not enabled",
        "element is not enabled"
    ])


def should_trigger_smart_retry(exception: Exception, consecutive_failures: int) -> bool:
    """
    判断是否应该触发智能重试机制
    
    :param exception: 捕获的异常
    :param consecutive_failures: 连续失败次数
    :return: 是否应该启动智能重试
    """
    # 检查是否为ElementNotEnabled异常
    if handle_element_not_enabled_exception(exception):
        return True
    
    # 检查连续失败次数
    if consecutive_failures >= 3:
        return True
    
    return False
