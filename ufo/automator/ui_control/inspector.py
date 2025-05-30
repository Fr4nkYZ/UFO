# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

from __future__ import annotations

import functools
import time
import logging
from abc import ABC, abstractmethod
from typing import Callable, Dict, List, Optional, cast

import comtypes.gen.UIAutomationClient as UIAutomationClient_dll
import psutil
import pywinauto
import pywinauto.uia_defines
import uiautomation as auto
from pywinauto import Desktop
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.uia_element_info import UIAElementInfo


from ufo.config.config import Config

configs = Config.get_instance().config_data

# Enhanced error handling configuration
ENHANCED_ERROR_HANDLING = configs.get("UI_ENHANCED_ERROR_HANDLING", True)
MAX_RETRY_ATTEMPTS = configs.get("UI_MAX_RETRY_ATTEMPTS", 3)
RETRY_DELAY_SECONDS = configs.get("UI_RETRY_DELAY_SECONDS", 0.5)


class BackendFactory:
    """
    A factory class to create backend strategies.
    """

    @staticmethod
    def create_backend(backend: str) -> BackendStrategy:
        """
        Create a backend strategy.
        :param backend: The backend to use.
        :return: The backend strategy.
        """
        if backend == "uia":
            return UIABackendStrategy()
        elif backend == "win32":
            return Win32BackendStrategy()
        else:
            raise ValueError(f"Backend {backend} not supported")


class BackendStrategy(ABC):
    """
    Define an interface for backend strategies.
    """

    @abstractmethod
    def get_desktop_windows(self, remove_empty: bool) -> List[UIAWrapper]:
        """
        Get all the apps on the desktop.
        :param remove_empty: Whether to remove empty titles.
        :return: The apps on the desktop.
        """
        pass

    @abstractmethod
    def find_control_elements_in_descendants(
        self,
        window: UIAWrapper,
        control_type_list: List[str] = [],
        class_name_list: List[str] = [],
        title_list: List[str] = [],
        is_visible: bool = True,
        is_enabled: bool = True,
        depth: int = 0,
    ) -> List[UIAWrapper]:
        """
        Find control elements in descendants of the window.
        :param window: The window to find control elements.
        :param control_type_list: The control types to find.
        :param class_name_list: The class names to find.
        :param title_list: The titles to find.
        :param is_visible: Whether the control elements are visible.
        :param is_enabled: Whether the control elements are enabled.
        :param depth: The depth of the descendants to find.
        :return: The control elements found.
        """

        pass


class UIAElementInfoFix(UIAElementInfo):
    _cached_rect = None
    _time_delay_marker = False

    def __init__(self, element, is_ref=False, source: Optional[str] = None):
        super().__init__(element, is_ref)

        self._source = source

    def sleep(self, ms: float = 0):
        import time

        if UIAElementInfoFix._time_delay_marker:
            ms = max(20, ms)
        else:
            ms = max(1, ms)
        time.sleep(ms / 1000.0)
        UIAElementInfoFix._time_delay_marker = False

    @staticmethod
    def _time_wrap(func):
        def dec(self, *args, **kvargs):
            name = func.__name__
            before = time.time()
            result = func(self, *args, **kvargs)
            if time.time() - before > 0.020:
                print(
                    f"[❌][{name}][{hash(self._element)}] lookup took {(time.time() - before) * 1000:.2f} ms"
                )
                UIAElementInfoFix._time_delay_marker = True
            elif time.time() - before > 0.005:
                print(
                    f"[⚠️][{name}][{hash(self._element)}]Control type lookup took {(time.time() - before) * 1000:.2f} ms"
                )
                UIAElementInfoFix._time_delay_marker = True
            else:
                # print(f"[✅][{name}][{hash(self._element)}]Control type lookup took {(time.time() - before) * 1000:.2f} ms")
                UIAElementInfoFix._time_delay_marker = False
            return result

        return dec

    @_time_wrap
    def _get_current_name(self):
        return super()._get_current_name()

    @_time_wrap
    def _get_current_rich_text(self):
        return super()._get_current_rich_text()

    @_time_wrap
    def _get_current_class_name(self):
        return super()._get_current_class_name()

    @_time_wrap
    def _get_current_control_type(self):
        return super()._get_current_control_type()

    @_time_wrap
    def _get_current_rectangle(self):
        bound_rect = self._element.CurrentBoundingRectangle
        rect = pywinauto.win32structures.RECT()
        rect.left = bound_rect.left
        rect.top = bound_rect.top
        rect.right = bound_rect.right
        rect.bottom = bound_rect.bottom
        return rect

    def _get_cached_rectangle(self) -> tuple[int, int, int, int]:
        if self._cached_rect is None:
            self._cached_rect = self._get_current_rectangle()
        return self._cached_rect

    @property
    def rectangle(self):
        return self._get_cached_rectangle()

    @property
    def source(self):
        return self._source


class UIABackendStrategy(BackendStrategy):
    """
    The backend strategy for UIA with enhanced error handling.
    """
    
    MAX_RETRIES = MAX_RETRY_ATTEMPTS
    RETRY_DELAY = RETRY_DELAY_SECONDS
    
    @staticmethod
    def _is_window_valid(window: UIAWrapper) -> bool:
        """
        Check if window is still valid and accessible.
        :param window: The window to check
        :return: True if window is valid, False otherwise
        """
        try:
            # Try to access basic properties to verify window is still valid
            window.is_enabled()
            window.element_info.rectangle
            return True
        except Exception as e:
            logging.debug(f"Window validation failed: {e}")
            return False
    
    @staticmethod
    def _handle_com_error(error: Exception, window: UIAWrapper, retry_count: int) -> Optional[str]:
        """
        Handle COM errors with appropriate logging and recovery strategy.
        :param error: The COM error that occurred
        :param window: The window being processed
        :param retry_count: Current retry attempt
        :return: Error message or None if should retry
        """
        error_code = getattr(error, 'hresult', None)
        
        # Common COM error codes and their meanings
        error_messages = {
            -2146233083: "UI element access denied or window destroyed",
            -2147024809: "Invalid parameter or window handle",
            -2147467259: "Unspecified error in COM operation",
            -2147023728: "Access denied to UI element",
        }
        
        error_msg = error_messages.get(error_code, f"Unknown COM error: {error_code}")
        
        if retry_count < UIABackendStrategy.MAX_RETRIES:
            logging.warning(f"COM error on retry {retry_count + 1}: {error_msg}. Retrying...")
            return None  # Signal to retry
        else:
            logging.error(f"COM error after {UIABackendStrategy.MAX_RETRIES} retries: {error_msg}")
            return error_msg

    def get_desktop_windows(self, remove_empty: bool) -> List[UIAWrapper]:
        """
        Get all the apps on the desktop.
        :param remove_empty: Whether to remove empty titles.
        :return: The apps on the desktop.
        """

        # UIA Com API would incur severe performance occasionally (such as a new app just started)
        # so we use Win32 to acquire the handle and then convert it to UIA interface

        desktop_windows = Desktop(backend="win32").windows()
        desktop_windows = [app for app in desktop_windows if app.is_visible()]

        if remove_empty:
            desktop_windows = [
                app
                for app in desktop_windows
                if app.window_text() != ""
                and app.element_info.class_name not in ["IME", "MSCTFIME UI"]
            ]

        uia_desktop_windows: List[UIAWrapper] = [
            UIAWrapper(UIAElementInfo(handle_or_elem=window.handle))
            for window in desktop_windows
        ]
        return uia_desktop_windows

    def find_control_elements_in_descendants(
        self,
        window: Optional[UIAWrapper],
        control_type_list: List[str] = [],
        class_name_list: List[str] = [],
        title_list: List[str] = [],
        is_visible: bool = True,
        is_enabled: bool = True,
        depth: int = 0,
    ) -> List[UIAWrapper]:
        """
        Find control elements in descendants of the window for uia backend with enhanced error handling.
        """
        if window is None:
            return []
            
        # Pre-validate window
        if not self._is_window_valid(window):
            logging.warning("Window is no longer valid, skipping control search")
            return []

        assert (
            class_name_list is None or len(class_name_list) == 0
        ), "class_name_list is not supported for UIA backend"

        for retry_count in range(self.MAX_RETRIES):
            try:
                return self._find_controls_with_cache(
                    window, control_type_list, is_visible, is_enabled
                )
                
            except Exception as error:
                # Handle COM errors specifically
                if hasattr(error, 'hresult'):
                    error_msg = self._handle_com_error(error, window, retry_count)
                    if error_msg is None:  # Should retry
                        time.sleep(self.RETRY_DELAY * (retry_count + 1))  # Exponential backoff
                        
                        # Re-validate window before retry
                        if not self._is_window_valid(window):
                            logging.warning("Window became invalid during retry, aborting")
                            return []
                        continue
                    else:
                        logging.error(f"Final COM error: {error_msg}")
                        return []
                else:
                    # Non-COM error, log and retry with shorter delay
                    logging.warning(f"Non-COM error on retry {retry_count + 1}: {str(error)}")
                    if retry_count < self.MAX_RETRIES - 1:
                        time.sleep(0.1)
                        continue
                    else:
                        logging.error(f"Final error after retries: {str(error)}")
                        return []
        
        return []

    def _find_controls_with_cache(
        self,
        window: UIAWrapper,
        control_type_list: List[str],
        is_visible: bool,
        is_enabled: bool
    ) -> List[UIAWrapper]:
        """
        Core method to find controls using cache with proper error handling.
        """
        _, iuia_dll = UIABackendStrategy._get_uia_defs()
        window_elem_info = cast(UIAElementInfo, window.element_info)
        window_elem_com_ref = cast(
            UIAutomationClient_dll.IUIAutomationElement, window_elem_info._element
        )

        # Validate COM reference
        if window_elem_com_ref is None:
            raise ValueError("Window COM reference is None")

        condition = UIABackendStrategy._get_control_filter_condition(
            control_type_list,
            is_visible,
            is_enabled,
        )

        cache_request = UIABackendStrategy._get_cache_request()

        # The critical COM call that may fail
        com_elem_array = window_elem_com_ref.FindAllBuildCache(
            scope=iuia_dll.TreeScope_Descendants,
            condition=condition,
            cacheRequest=cache_request,
        )

        if com_elem_array is None:
            logging.warning("FindAllBuildCache returned None")
            return []

        # Process elements with additional error handling
        return self._process_cached_elements(com_elem_array)

    def _process_cached_elements(self, com_elem_array) -> List[UIAWrapper]:
        """
        Process cached elements with individual error handling.
        """
        control_elements: List[UIAWrapper] = []
        
        try:
            array_length = min(com_elem_array.Length, 500)
        except Exception as e:
            logging.error(f"Failed to get array length: {e}")
            return []

        for n in range(array_length):
            try:
                elem = com_elem_array.GetElement(n)
                if elem is None:
                    continue
                    
                # Extract cached properties with error handling
                try:
                    elem_type = elem.CachedControlType
                    elem_name = elem.CachedName
                    elem_rect = elem.CachedBoundingRectangle
                except Exception as e:
                    logging.debug(f"Failed to get cached properties for element {n}: {e}")
                    continue

                # Skip controls with invalid/zero rectangles
                if (elem_rect.right - elem_rect.left <= 0 or 
                    elem_rect.bottom - elem_rect.top <= 0):
                    continue
                    
                # Create UI element with error handling
                try:
                    uia_wrapper = self._create_uia_wrapper(elem, elem_type, elem_name, elem_rect)
                    if uia_wrapper:
                        control_elements.append(uia_wrapper)
                except Exception as e:
                    logging.debug(f"Failed to create wrapper for element {n}: {e}")
                    continue
                    
            except Exception as e:
                logging.debug(f"Error processing element {n}: {e}")
                continue

        return control_elements

    def _create_uia_wrapper(self, elem, elem_type, elem_name, elem_rect) -> Optional[UIAWrapper]:
        """
        Create UIA wrapper with proper error handling.
        """
        try:
            element_info = UIAElementInfoFix(elem, True, source="uia")
            elem_type_name = UIABackendStrategy._get_uia_control_name_map().get(
                elem_type, ""
            )

            # Set cached properties
            element_info._cached_handle = 0
            element_info._cached_visible = True

            # Fill rectangle
            rect = pywinauto.win32structures.RECT()
            rect.left = elem_rect.left
            rect.top = elem_rect.top
            rect.right = elem_rect.right
            rect.bottom = elem_rect.bottom
            element_info._cached_rect = rect
            element_info._cached_name = elem_name
            element_info._cached_control_type = elem_type_name
            element_info._cached_rich_text = elem_name

            uia_interface = UIAWrapper(element_info)

            def __hash__(self):
                return hash(self.element_info._element)

            uia_interface.__hash__ = __hash__
            return uia_interface
            
        except Exception as e:
            logging.debug(f"Failed to create UIA wrapper: {e}")
            return None

    @staticmethod
    def _get_uia_control_id_map():
        iuia = pywinauto.uia_defines.IUIA()
        return iuia.known_control_types

    @staticmethod
    def _get_uia_control_name_map():
        iuia = pywinauto.uia_defines.IUIA()
        return iuia.known_control_type_ids

    @staticmethod
    @functools.lru_cache()
    def _get_cache_request():
        iuia_com, iuia_dll = UIABackendStrategy._get_uia_defs()
        cache_request = iuia_com.CreateCacheRequest()
        cache_request.AddProperty(iuia_dll.UIA_ControlTypePropertyId)
        cache_request.AddProperty(iuia_dll.UIA_NamePropertyId)
        cache_request.AddProperty(iuia_dll.UIA_BoundingRectanglePropertyId)
        return cache_request

    @staticmethod
    def _get_control_filter_condition(
        control_type_list: List[str] = [],
        is_visible: bool = True,
        is_enabled: bool = True,
    ):
        iuia_com, iuia_dll = UIABackendStrategy._get_uia_defs()
        condition = iuia_com.CreateAndConditionFromArray(
            [
                iuia_com.CreatePropertyCondition(
                    iuia_dll.UIA_IsEnabledPropertyId, is_enabled
                ),
                iuia_com.CreatePropertyCondition(
                    # visibility is determined by IsOffscreen property
                    iuia_dll.UIA_IsOffscreenPropertyId,
                    not is_visible,
                ),
                iuia_com.CreatePropertyCondition(
                    iuia_dll.UIA_IsControlElementPropertyId, True
                ),
                iuia_com.CreateOrConditionFromArray(
                    [
                        iuia_com.CreatePropertyCondition(
                            iuia_dll.UIA_ControlTypePropertyId,
                            (
                                control_type
                                if control_type is int
                                else UIABackendStrategy._get_uia_control_id_map()[
                                    control_type
                                ]
                            ),
                        )
                        for control_type in control_type_list
                    ]
                ),
            ]
        )
        return condition

    @staticmethod
    def _get_uia_defs():
        iuia = pywinauto.uia_defines.IUIA()
        iuia_com: UIAutomationClient_dll.IUIAutomation = iuia.iuia
        iuia_dll: UIAutomationClient_dll = iuia.UIA_dll
        return iuia_com, iuia_dll


class Win32BackendStrategy(BackendStrategy):
    """
    The backend strategy for Win32.
    """

    def get_desktop_windows(self, remove_empty: bool) -> List[UIAWrapper]:
        """
        Get all the apps on the desktop.
        :param remove_empty: Whether to remove empty titles.
        :return: The apps on the desktop.
        """

        desktop_windows = Desktop(backend="win32").windows()
        desktop_windows = [app for app in desktop_windows if app.is_visible()]

        if remove_empty:
            desktop_windows = [
                app
                for app in desktop_windows
                if app.window_text() != ""
                and app.element_info.class_name not in ["IME", "MSCTFIME UI"]
            ]
        return desktop_windows

    def find_control_elements_in_descendants(
        self,
        window: UIAWrapper,
        control_type_list: List[str] = [],
        class_name_list: List[str] = [],
        title_list: List[str] = [],
        is_visible: bool = True,
        is_enabled: bool = True,
        depth: int = 0,
    ) -> List[UIAWrapper]:
        """
        Find control elements in descendants of the window for win32 backend.
        :param window: The window to find control elements.
        :param control_type_list: The control types to find.
        :param class_name_list: The class names to find.
        :param title_list: The titles to find.
        :param is_visible: Whether the control elements are visible.
        :param is_enabled: Whether the control elements are enabled.
        :param depth: The depth of the descendants to find.
        :return: The control elements found.
        """

        if window == None:
            return []

        control_elements = []
        if len(class_name_list) == 0:
            control_elements += window.descendants()
        else:
            for class_name in class_name_list:
                if depth == 0:
                    subcontrols = window.descendants(class_name=class_name)
                else:
                    subcontrols = window.descendants(class_name=class_name, depth=depth)
                control_elements += subcontrols

        if is_visible:
            control_elements = [
                control for control in control_elements if control.is_visible()
            ]
        if is_enabled:
            control_elements = [
                control for control in control_elements if control.is_enabled()
            ]
        if len(title_list) > 0:
            control_elements = [
                control
                for control in control_elements
                if control.window_text() in title_list
            ]
        if len(control_type_list) > 0:
            control_elements = [
                control
                for control in control_elements
                if control.element_info.control_type in control_type_list
            ]

        return [
            control for control in control_elements if control.element_info.name != ""
        ]


class ControlInspectorFacade:
    """
    The singleton facade class for control inspector.
    """

    _instances = {}

    def __new__(cls, backend: str = "uia") -> "ControlInspectorFacade":
        """
        Singleton pattern.
        """
        if backend not in cls._instances:
            instance = super().__new__(cls)
            instance.backend = backend
            instance.backend_strategy = BackendFactory.create_backend(backend)
            cls._instances[backend] = instance
        return cls._instances[backend]

    def __init__(self, backend: str = "uia") -> None:
        """
        Initialize the control inspector.
        :param backend: The backend to use.
        """
        self.backend = backend

    def get_desktop_windows(self, remove_empty: bool = True) -> List[UIAWrapper]:
        """
        Get all the apps on the desktop.
        :param remove_empty: Whether to remove empty titles.
        :return: The apps on the desktop.
        """
        return self.backend_strategy.get_desktop_windows(remove_empty)

    def find_control_elements_in_descendants(
        self,
        window: UIAWrapper,
        control_type_list: List[str] = [],
        class_name_list: List[str] = [],
        title_list: List[str] = [],
        is_visible: bool = True,
        is_enabled: bool = True,
        depth: int = 0,
    ) -> List[UIAWrapper]:
        """
        Find control elements in descendants of the window.
        :param window: The window to find control elements.
        :param control_type_list: The control types to find.
        :param class_name_list: The class names to find.
        :param title_list: The titles to find.
        :param is_visible: Whether the control elements are visible.
        :param is_enabled: Whether the control elements are enabled.
        :param depth: The depth of the descendants to find.
        :return: The control elements found.
        """
        if self.backend == "uia":
            return self.backend_strategy.find_control_elements_in_descendants(
                window, control_type_list, [], title_list, is_visible, is_enabled, depth
            )
        elif self.backend == "win32":
            return self.backend_strategy.find_control_elements_in_descendants(
                window, [], class_name_list, title_list, is_visible, is_enabled, depth
            )
        else:
            return []

    def get_desktop_app_dict(self, remove_empty: bool = True) -> Dict[str, UIAWrapper]:
        """
        Get all the apps on the desktop and return as a dict.
        :param remove_empty: Whether to remove empty titles.
        :return: The apps on the desktop as a dict.
        """
        desktop_windows = self.get_desktop_windows(remove_empty)

        desktop_windows_with_gui = []

        for window in desktop_windows:
            try:
                window.is_normal()
                desktop_windows_with_gui.append(window)
            except:
                pass

        desktop_windows_dict = dict(
            zip(
                [str(i + 1) for i in range(len(desktop_windows_with_gui))],
                desktop_windows_with_gui,
            )
        )
        return desktop_windows_dict

    def get_desktop_app_info(
        self,
        desktop_windows_dict: Dict[str, UIAWrapper],
        field_list: List[str] = ["control_text", "control_type"],
    ) -> List[Dict[str, str]]:
        """
        Get control info of all the apps on the desktop.
        :param desktop_windows_dict: The dict of apps on the desktop.
        :param field_list: The fields of app info to get.
        :return: The control info of all the apps on the desktop.
        """
        desktop_windows_info = self.get_control_info_list_of_dict(
            desktop_windows_dict, field_list
        )
        return desktop_windows_info

    def get_control_info_batch(
        self, window_list: List[UIAWrapper], field_list: List[str] = []
    ) -> List[Dict[str, str]]:
        """
        Get control info of the window.
        :param window_list: The list of windows to get control info.
        :param field_list: The fields to get.
        return: The list of control info of the window.
        """
        control_info_list = []
        for window in window_list:
            control_info_list.append(self.get_control_info(window, field_list))
        return control_info_list

    def get_control_info_list_of_dict(
        self, window_dict: Dict[str, UIAWrapper], field_list: List[str] = []
    ) -> List[Dict[str, str]]:
        """
        Get control info of the window.
        :param window_dict: The dict of windows to get control info.
        :param field_list: The fields to get.
        return: The list of control info of the window.
        """
        control_info_list = []
        for key in window_dict.keys():
            window = window_dict[key]
            control_info = self.get_control_info(window, field_list)
            control_info["label"] = key
            control_info_list.append(control_info)
        return control_info_list

    @staticmethod
    def get_check_state(control_item: auto.Control) -> bool | None:
        """
        get the check state of the control item
        param control_item: the control item to get the check state
        return: the check state of the control item
        """
        is_checked = None
        is_selected = None
        try:
            assert isinstance(
                control_item, auto.Control
            ), f"{control_item =} is not a Control"
            is_checked = (
                control_item.GetLegacyIAccessiblePattern().State
                & auto.AccessibleState.Checked
                == auto.AccessibleState.Checked
            )
            if is_checked:
                return is_checked
            is_selected = (
                control_item.GetLegacyIAccessiblePattern().State
                & auto.AccessibleState.Selected
                == auto.AccessibleState.Selected
            )
            if is_selected:
                return is_selected
            return None
        except Exception as e:
            # print(f'item {control_item} not available for check state.')
            # print(e)
            return None

    @staticmethod
    def get_control_info(
        window: UIAWrapper, field_list: List[str] = []
    ) -> Dict[str, str]:
        """
        Get control info of the window.
        :param window: The window to get control info.
        :param field_list: The fields to get.
        return: The control info of the window.
        """
        control_info: Dict[str, str] = {}

        def assign(prop_name: str, prop_value_func: Callable[[], str]) -> None:
            if len(field_list) > 0 and prop_name not in field_list:
                return
            control_info[prop_name] = prop_value_func()

        try:
            assign("control_type", lambda: window.element_info.control_type)
            assign("control_id", lambda: window.element_info.control_id)
            assign("control_class", lambda: window.element_info.class_name)
            assign("control_name", lambda: window.element_info.name)
            rectangle = window.element_info.rectangle
            assign(
                "control_rect",
                lambda: (
                    rectangle.left,
                    rectangle.top,
                    rectangle.right,
                    rectangle.bottom,
                ),
            )
            assign("control_text", lambda: window.element_info.name)
            assign("control_title", lambda: window.window_text())
            assign("selected", lambda: ControlInspectorFacade.get_check_state(window))

            try:
                source = window.element_info.source
                assign("source", lambda: source)
            except:
                assign("source", lambda: "")

            return control_info
        except:
            return {}

    @staticmethod
    def get_application_root_name(window: UIAWrapper) -> str:
        """
        Get the application name of the window.
        :param window: The window to get the application name.
        :return: The root application name of the window. Empty string ("") if failed to get the name.
        """
        if window == None:
            return ""
        process_id = window.process_id()
        try:
            process = psutil.Process(process_id)
            return process.name()
        except psutil.NoSuchProcess:
            return ""

    @property
    def desktop(self) -> UIAWrapper:
        """
        Get all the desktop windows.
        :return: The uia wrapper of the desktop.
        """
        desktop_element = UIAElementInfo()
        return UIAWrapper(desktop_element)

    def find_control_elements_in_descendants(
        self,
        window: UIAWrapper,
        control_type_list: List[str] = [],
        class_name_list: List[str] = [],
        title_list: List[str] = [],
        is_visible: bool = True,
        is_enabled: bool = True,
        depth: int = 0,
    ) -> List[UIAWrapper]:
        """
        Find control elements in descendants of the window with enhanced error handling.
        """
        if window is None:
            logging.warning("Window is None, returning empty control list")
            return []
            
        try:
            if self.backend == "uia":
                return self.backend_strategy.find_control_elements_in_descendants(
                    window, control_type_list, [], title_list, is_visible, is_enabled, depth
                )
            elif self.backend == "win32":
                return self.backend_strategy.find_control_elements_in_descendants(
                    window, [], class_name_list, title_list, is_visible, is_enabled, depth
                )
            else:
                logging.warning(f"Unsupported backend: {self.backend}")
                return []
                
        except Exception as e:
            logging.error(f"Error finding control elements: {str(e)}")
            # Try fallback strategy
            return self._fallback_control_search(window, control_type_list)
    
    def _fallback_control_search(self, window: UIAWrapper, control_type_list: List[str]) -> List[UIAWrapper]:
        """
        Fallback control search using simpler methods.
        """
        try:
            logging.info("Attempting fallback control search")
            
            # Try to use basic descendants method without cache
            if hasattr(window, 'descendants'):
                descendants = window.descendants()
                # Filter by control type if specified
                if control_type_list:
                    descendants = [
                        d for d in descendants 
                        if hasattr(d.element_info, 'control_type') and 
                        d.element_info.control_type in control_type_list
                    ]
                return descendants[:100]  # Limit to prevent performance issues
                
        except Exception as e:
            logging.error(f"Fallback control search also failed: {str(e)}")
            
        return []
    
    def safe_get_desktop_app_dict(self, remove_empty: bool = True) -> Dict[str, UIAWrapper]:
        """
        Safely get desktop app dict with error handling.
        """
        try:
            return self.get_desktop_app_dict(remove_empty)
        except Exception as e:
            logging.error(f"Error getting desktop apps: {str(e)}")
            # Return minimal safe dict
            return {}
