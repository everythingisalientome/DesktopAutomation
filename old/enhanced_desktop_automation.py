import time
import os
import hashlib
import subprocess
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple, Union
from dataclasses import dataclass
from pathlib import Path

import pyautogui
import easyocr
import cv2
import numpy as np
from PIL import Image
import win32gui
import win32con
import win32api
import win32com.client
from win32com.client import Dispatch

from langgraph.graph import StateGraph, END
from typing_extensions import TypedDict

# Configure pyautogui
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

@dataclass
class AutomationStep:
    window_name: str
    field_name: str
    field_type: str
    event_action_type: str
    special_key_with_data: str
    data: str

class AutomationState(TypedDict):
    exe_path: str
    steps: List[Dict[str, Any]]
    current_step: int
    current_step_data: Dict[str, Any]
    error_message: str
    retry_count: int
    max_retries: int
    
    # Detection components
    ocr_reader: Any
    ui_automation: Any
    element_cache: Dict[str, Any]
    current_window_handle: int
    
    # Template and logging
    templates_dir: str
    analytics_logger: Any
    
    # Element detection results
    element_coords: Optional[Tuple[int, int]]
    element_found: bool
    detection_method: str
    element_handle: Any

class EnhancedDesktopAutomation:
    def __init__(self, templates_dir: str = "templates", logs_dir: str = "logs"):
        self.templates_dir = Path(templates_dir)
        self.logs_dir = Path(logs_dir)
        self.max_retries = 1
        self.detection_timeout = 20
        
        # Create directories
        self.templates_dir.mkdir(exist_ok=True)
        self.logs_dir.mkdir(exist_ok=True)
        
        # Setup logging
        self.setup_logging()
        
        # Initialize detection components
        self.ocr_reader = None
        self.ui_automation = None
        
        # Setup workflow
        self.setup_graph()
    
    def setup_logging(self):
        """Setup analytics logging"""
        log_file = self.logs_dir / "automation_analytics.log"
        
        # Create logger
        self.analytics_logger = logging.getLogger('automation_analytics')
        self.analytics_logger.setLevel(logging.INFO)
        
        # Create file handler
        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(logging.INFO)
        
        # Create formatter
        formatter = logging.Formatter('%(asctime)s - %(message)s')
        file_handler.setFormatter(formatter)
        
        # Add handler to logger
        if not self.analytics_logger.handlers:
            self.analytics_logger.addHandler(file_handler)
    
    def setup_graph(self):
        """Setup the LangGraph workflow"""
        workflow = StateGraph(AutomationState)
        
        # Add nodes
        workflow.add_node("initialize", self.initialize_automation)
        workflow.add_node("launch_application", self.launch_application)
        workflow.add_node("process_step", self.process_step)
        workflow.add_node("activate_window", self.activate_window)
        workflow.add_node("find_element", self.find_element)
        workflow.add_node("execute_action", self.execute_action)
        workflow.add_node("handle_error", self.handle_error)
        workflow.add_node("complete", self.complete_automation)
        
        # Set entry point
        workflow.set_entry_point("initialize")
        
        # Add edges
        workflow.add_edge("initialize", "launch_application")
        workflow.add_conditional_edges(
            "launch_application",
            self.check_launch_success,
            {"success": "process_step", "error": "handle_error"}
        )
        workflow.add_conditional_edges(
            "process_step",
            self.check_steps_remaining,
            {"continue": "activate_window", "complete": "complete"}
        )
        workflow.add_conditional_edges(
            "activate_window",
            self.check_window_activation,
            {"success": "find_element", "error": "handle_error"}
        )
        workflow.add_conditional_edges(
            "find_element",
            self.check_element_found,
            {"found": "execute_action", "not_found": "handle_error"}
        )
        workflow.add_conditional_edges(
            "execute_action",
            self.check_action_success,
            {"success": "process_step", "error": "handle_error"}
        )
        workflow.add_edge("handle_error", END)
        workflow.add_edge("complete", END)
        
        self.graph = workflow.compile()
    
    def initialize_automation(self, state: AutomationState) -> AutomationState:
        """Initialize the automation process"""
        try:
            # Initialize OCR reader
            if not self.ocr_reader:
                self.ocr_reader = easyocr.Reader(['en'], gpu=False)
            
            # Initialize UI Automation
            if not self.ui_automation:
                self.ui_automation = Dispatch("UIAutomationCore.CUIAutomation")
            
            state.update({
                "ocr_reader": self.ocr_reader,
                "ui_automation": self.ui_automation,
                "element_cache": {},
                "current_step": 0,
                "retry_count": 0,
                "max_retries": self.max_retries,
                "error_message": "",
                "current_window_handle": None,
                "templates_dir": str(self.templates_dir),
                "analytics_logger": self.analytics_logger,
                "element_coords": None,
                "element_found": False,
                "detection_method": "",
                "element_handle": None
            })
            
            self.analytics_logger.info("Automation initialized successfully")
            return state
            
        except Exception as e:
            state["error_message"] = f"Failed to initialize automation: {str(e)}"
            return state
    
    def launch_application(self, state: AutomationState) -> AutomationState:
        """Launch the target application"""
        exe_path = state["exe_path"]
        
        try:
            process = subprocess.Popen([exe_path])
            time.sleep(15)  # Wait for application to launch
            
            self.analytics_logger.info(f"Application launched: {exe_path}")
            return state
            
        except Exception as e:
            state["error_message"] = f"Failed to launch application: {str(e)}"
            return state
    
    def process_step(self, state: AutomationState) -> AutomationState:
        """Process the current automation step"""
        current_step = state["current_step"]
        steps = state["steps"]
        
        if current_step < len(steps):
            step_data = steps[current_step]
            state["current_step_data"] = step_data
            state["retry_count"] = 0
            
            # Skip windowActivate steps as we handle window activation automatically
            if step_data.get("event_action_type") == "windowActivate":
                state["current_step"] += 1
                return self.process_step(state)  # Process next step
        
        return state
    
    def activate_window(self, state: AutomationState) -> AutomationState:
        """Activate the target window before element search"""
        step_data = state["current_step_data"]
        window_name = step_data.get("window_name", "")
        
        try:
            window_handle = self._find_window_by_name(window_name)
            if window_handle:
                win32gui.SetForegroundWindow(window_handle)
                win32gui.ShowWindow(window_handle, win32con.SW_RESTORE)
                time.sleep(0.5)  # Wait for window activation
                
                state["current_window_handle"] = window_handle
                self.analytics_logger.info(f"Window activated: {window_name}")
                return state
            else:
                state["error_message"] = f"Window not found: {window_name}"
                return state
                
        except Exception as e:
            state["error_message"] = f"Failed to activate window '{window_name}': {str(e)}"
            return state
    
    def find_element(self, state: AutomationState) -> AutomationState:
        """Find element using hierarchical detection methods"""
        step_data = state["current_step_data"]
        window_handle = state["current_window_handle"]
        field_name = step_data["field_name"]
        field_type = step_data["field_type"]
        
        start_time = time.time()
        
        try:
            # Method 1: UI Automation
            element_info = self._find_element_ui_automation(window_handle, field_name, field_type)
            if element_info:
                state.update({
                    "element_coords": element_info["coords"],
                    "element_handle": element_info.get("handle"),
                    "element_found": True,
                    "detection_method": "UI_Automation"
                })
                self._capture_template(element_info["coords"], field_name, field_type)
                self._log_detection_success("UI_Automation", field_name, field_type, time.time() - start_time)
                return state
            
            # Method 2: Win32 API
            element_info = self._find_element_win32(window_handle, field_name, field_type)
            if element_info:
                state.update({
                    "element_coords": element_info["coords"],
                    "element_handle": element_info.get("handle"),
                    "element_found": True,
                    "detection_method": "Win32_API"
                })
                self._capture_template(element_info["coords"], field_name, field_type)
                self._log_detection_success("Win32_API", field_name, field_type, time.time() - start_time)
                return state
            
            # Method 3: Template Matching
            coords = self._find_element_template_matching(field_name, field_type)
            if coords:
                state.update({
                    "element_coords": coords,
                    "element_handle": None,
                    "element_found": True,
                    "detection_method": "Template_Matching"
                })
                self._log_detection_success("Template_Matching", field_name, field_type, time.time() - start_time)
                return state
            
            # Method 4: OCR (Last Resort)
            coords = self._find_element_ocr(field_name, state["ocr_reader"])
            if coords:
                state.update({
                    "element_coords": coords,
                    "element_handle": None,
                    "element_found": True,
                    "detection_method": "OCR"
                })
                self._log_detection_success("OCR", field_name, field_type, time.time() - start_time)
                return state
            
            # All methods failed
            state["element_found"] = False
            state["error_message"] = f"Element '{field_name}' not found using any detection method"
            self._log_detection_failure(field_name, field_type, time.time() - start_time)
            return state
            
        except Exception as e:
            state["element_found"] = False
            state["error_message"] = f"Error finding element '{field_name}': {str(e)}"
            return state
    
    def execute_action(self, state: AutomationState) -> AutomationState:
        """Execute the specified action"""
        step_data = state["current_step_data"]
        element_coords = state["element_coords"]
        element_handle = state["element_handle"]
        action_type = step_data["event_action_type"]
        data = step_data.get("data", "")
        special_key = step_data.get("SpecialKeyWithData", "")
        
        try:
            if action_type == "click":
                pyautogui.click(element_coords[0], element_coords[1])
                
            elif action_type == "typeInto":
                pyautogui.click(element_coords[0], element_coords[1])
                time.sleep(0.5)
                pyautogui.typewrite(data)
                
            elif action_type == "keyPress":
                if special_key:
                    self._send_special_key(special_key)
                else:
                    pyautogui.press(data.lower())
                    
            else:
                raise ValueError(f"Unsupported action type: {action_type}")
            
            time.sleep(1)  # Wait after action
            state["current_step"] += 1
            
            self.analytics_logger.info(f"Action executed: {action_type} on {step_data['field_name']}")
            return state
            
        except Exception as e:
            state["error_message"] = f"Action execution failed: {str(e)}"
            return state
    
    def complete_automation(self, state: AutomationState) -> AutomationState:
        """Complete the automation process"""
        self.analytics_logger.info("Automation completed successfully!")
        return state
    
    def handle_error(self, state: AutomationState) -> AutomationState:
        """Handle errors"""
        error_message = state.get("error_message", "Unknown error")
        self.analytics_logger.error(f"Automation failed: {error_message}")
        raise Exception(error_message)
    
    # Helper methods for element detection
    
    def _find_window_by_name(self, window_name: str) -> Optional[int]:
        """Find window handle by name (contains matching)"""
        def enum_windows_proc(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if window_name.lower() in window_text.lower():
                    lParam.append(hwnd)
            return True
        
        windows = []
        win32gui.EnumWindows(enum_windows_proc, windows)
        return windows[0] if windows else None
    
    def _find_element_ui_automation(self, window_handle: int, field_name: str, field_type: str) -> Optional[Dict]:
        """Find element using Windows UI Automation"""
        try:
            # Get the automation element for the window
            window_element = self.ui_automation.ElementFromHandle(window_handle)
            if not window_element:
                return None
            
            # Define search conditions based on field_type
            condition = None
            if field_type == "push button":
                condition = self.ui_automation.CreatePropertyCondition(
                    30003,  # UIA_ControlTypePropertyId 
                    50000   # UIA_ButtonControlTypeId
                )
            elif field_type == "editable text":
                condition = self.ui_automation.CreatePropertyCondition(
                    30003,  # UIA_ControlTypePropertyId
                    50004   # UIA_EditControlTypeId
                )
            elif field_type == "link":
                condition = self.ui_automation.CreatePropertyCondition(
                    30003,  # UIA_ControlTypePropertyId
                    50005   # UIA_HyperlinkControlTypeId
                )
            
            # Search for elements
            if condition:
                elements = window_element.FindAll(2, condition)  # TreeScope_Descendants
                for i in range(elements.Length):
                    element = elements.GetElement(i)
                    element_name = element.GetCurrentPropertyValue(30005)  # UIA_NamePropertyId
                    
                    if element_name and field_name.lower() in element_name.lower():
                        # Get element rectangle
                        rect = element.GetCurrentPropertyValue(30001)  # UIA_BoundingRectanglePropertyId
                        if rect:
                            center_x = int(rect.left + rect.width / 2)
                            center_y = int(rect.top + rect.height / 2)
                            return {
                                "coords": (center_x, center_y),
                                "handle": element,
                                "rect": rect
                            }
            
            return None
            
        except Exception as e:
            return None
    
    def _find_element_win32(self, window_handle: int, field_name: str, field_type: str) -> Optional[Dict]:
        """Find element using Win32 API"""
        try:
            def enum_child_proc(hwnd, lParam):
                window_text = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                
                # Check if text matches
                if window_text and field_name.lower() in window_text.lower():
                    # Get window rectangle
                    rect = win32gui.GetWindowRect(hwnd)
                    center_x = (rect[0] + rect[2]) // 2
                    center_y = (rect[1] + rect[3]) // 2
                    
                    lParam.append({
                        "coords": (center_x, center_y),
                        "handle": hwnd,
                        "text": window_text,
                        "class": class_name
                    })
                
                return True
            
            elements = []
            win32gui.EnumChildWindows(window_handle, enum_child_proc, elements)
            
            # Return first matching element
            return elements[0] if elements else None
            
        except Exception as e:
            return None
    
    def _find_element_template_matching(self, field_name: str, field_type: str) -> Optional[Tuple[int, int]]:
        """Find element using OpenCV template matching"""
        try:
            template_path = self._get_template_path(field_name, field_type)
            if not template_path.exists():
                return None
            
            # Take screenshot
            screenshot = pyautogui.screenshot()
            screenshot_np = np.array(screenshot)
            screenshot_gray = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2GRAY)
            
            # Load template
            template = cv2.imread(str(template_path), 0)
            if template is None:
                return None
            
            # Template matching
            result = cv2.matchTemplate(screenshot_gray, template, cv2.TM_CCOEFF_NORMED)
            _, max_val, _, max_loc = cv2.minMaxLoc(result)
            
            # Check confidence threshold
            if max_val > 0.8:
                template_h, template_w = template.shape
                center_x = max_loc[0] + template_w // 2
                center_y = max_loc[1] + template_h // 2
                return (center_x, center_y)
            
            return None
            
        except Exception as e:
            return None
    
    def _find_element_ocr(self, field_name: str, ocr_reader) -> Optional[Tuple[int, int]]:
        """Find element using OCR"""
        try:
            screenshot = pyautogui.screenshot()
            screenshot_np = np.array(screenshot)
            
            results = ocr_reader.readtext(screenshot_np)
            
            for (bbox, detected_text, confidence) in results:
                if confidence > 0.5 and field_name.lower() in detected_text.lower():
                    x_coords = [point[0] for point in bbox]
                    y_coords = [point[1] for point in bbox]
                    center_x = int(sum(x_coords) / len(x_coords))
                    center_y = int(sum(y_coords) / len(y_coords))
                    return (center_x, center_y)
            
            return None
            
        except Exception as e:
            return None
    
    def _capture_template(self, coords: Tuple[int, int], field_name: str, field_type: str):
        """Capture template image for future use"""
        try:
            template_path = self._get_template_path(field_name, field_type)
            
            if template_path.exists():
                return  # Template already exists
            
            # Create template directory if needed
            template_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Capture screenshot around the element
            x, y = coords
            margin = 20
            
            screenshot = pyautogui.screenshot(region=(
                max(0, x - margin),
                max(0, y - margin),
                margin * 2,
                margin * 2
            ))
            
            screenshot.save(template_path)
            self.analytics_logger.info(f"Template captured: {template_path}")
            
        except Exception as e:
            self.analytics_logger.warning(f"Failed to capture template: {str(e)}")
    
    def _get_template_path(self, field_name: str, field_type: str) -> Path:
        """Generate template file path"""
        field_hash = hashlib.md5(field_name.encode()).hexdigest()[:8]
        filename = f"{field_type}_{field_hash}.png"
        return self.templates_dir / filename
    
    def _send_special_key(self, special_key: str):
        """Send special key combinations"""
        key_map = {
            "Enter": "enter",
            "Tab": "tab",
            "Escape": "esc",
            "Space": "space",
            "Backspace": "backspace",
            "Delete": "del",
            "Home": "home",
            "End": "end",
            "PageUp": "pageup",
            "PageDown": "pagedown",
            "Up": "up",
            "Down": "down",
            "Left": "left",
            "Right": "right"
        }
        
        # Handle function keys
        for i in range(1, 13):
            key_map[f"F{i}"] = f"f{i}"
        
        # Handle Ctrl combinations
        if special_key.startswith("Ctrl+"):
            key = special_key.split("+")[1].lower()
            pyautogui.hotkey('ctrl', key)
        elif special_key.startswith("Alt+"):
            key = special_key.split("+")[1].lower()
            pyautogui.hotkey('alt', key)
        elif special_key.startswith("Shift+"):
            key = special_key.split("+")[1].lower()
            pyautogui.hotkey('shift', key)
        else:
            key = key_map.get(special_key, special_key.lower())
            pyautogui.press(key)
    
    def _log_detection_success(self, method: str, field_name: str, field_type: str, duration: float):
        """Log successful element detection"""
        self.analytics_logger.info(
            f"DETECTION_SUCCESS - Method: {method}, Element: {field_name}, "
            f"Type: {field_type}, Duration: {duration:.2f}s"
        )
    
    def _log_detection_failure(self, field_name: str, field_type: str, duration: float):
        """Log failed element detection"""
        self.analytics_logger.warning(
            f"DETECTION_FAILURE - Element: {field_name}, Type: {field_type}, "
            f"Duration: {duration:.2f}s"
        )
    
    # Conditional edge functions
    def check_launch_success(self, state: AutomationState) -> str:
        return "error" if state.get("error_message") else "success"
    
    def check_steps_remaining(self, state: AutomationState) -> str:
        return "complete" if state["current_step"] >= len(state["steps"]) else "continue"
    
    def check_window_activation(self, state: AutomationState) -> str:
        return "error" if state.get("error_message") else "success"
    
    def check_element_found(self, state: AutomationState) -> str:
        return "found" if state.get("element_found", False) else "not_found"
    
    def check_action_success(self, state: AutomationState) -> str:
        return "error" if state.get("error_message") else "success"
    
    def run_automation(self, exe_path: str, steps_json: List[Dict[str, Any]]) -> None:
        """Run the automation process"""
        initial_state = AutomationState(
            exe_path=exe_path,
            steps=steps_json,
            current_step=0,
            current_step_data={},
            error_message="",
            retry_count=0,
            max_retries=self.max_retries,
            ocr_reader=None,
            ui_automation=None,
            element_cache={},
            current_window_handle=None,
            templates_dir="",
            analytics_logger=None,
            element_coords=None,
            element_found=False,
            detection_method="",
            element_handle=None
        )
        
        try:
            self.graph.invoke(initial_state)
            self.analytics_logger.info("Automation completed successfully!")
            
        except Exception as e:
            self.analytics_logger.error(f"Automation failed: {str(e)}")
            raise