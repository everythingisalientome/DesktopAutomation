import time
import subprocess
import pyautogui
import easyocr
import numpy as np
from PIL import Image
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass
from langgraph.graph import StateGraph, END
from langgraph.graph.message import add_messages
from typing_extensions import Annotated, TypedDict
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configure pyautogui
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

@dataclass
class AutomationStep:
    field_name: str
    field_type: str
    event_action_type: str
    special_key_with_data: str
    data: str

class AutomationState(TypedDict):
    exe_path: str
    steps: List[Dict[str, Any]]
    current_step: int
    error_message: str
    retry_count: int
    max_retries: int
    ocr_reader: Any
    process: Any

class DesktopAutomationTool:
    def __init__(self):
        self.ocr_reader = None
        self.max_retries = 1
        self.setup_graph()
    
    def setup_graph(self):
        """Setup the LangGraph workflow"""
        workflow = StateGraph(AutomationState)
        
        # Add nodes
        workflow.add_node("initialize", self.initialize_automation)
        workflow.add_node("launch_application", self.launch_application)
        workflow.add_node("process_step", self.process_step)
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
            {
                "success": "process_step",
                "error": "handle_error"
            }
        )
        workflow.add_conditional_edges(
            "process_step",
            self.check_steps_remaining,
            {
                "continue": "find_element",
                "complete": "complete"
            }
        )
        workflow.add_conditional_edges(
            "find_element",
            self.check_element_found,
            {
                "found": "execute_action",
                "not_found": "handle_error"
            }
        )
        workflow.add_conditional_edges(
            "execute_action",
            self.check_action_success,
            {
                "success": "process_step",
                "retry": "find_element",
                "error": "handle_error"
            }
        )
        workflow.add_edge("handle_error", END)
        workflow.add_edge("complete", END)
        
        self.graph = workflow.compile()
    
    def initialize_automation(self, state: AutomationState) -> AutomationState:
        """Initialize the automation process"""
        logger.info("Initializing automation...")
        
        try:
            # Initialize OCR reader
            if not self.ocr_reader:
                logger.info("Initializing EasyOCR reader...")
                self.ocr_reader = easyocr.Reader(['en'], gpu=False)
            
            state["ocr_reader"] = self.ocr_reader
            state["current_step"] = 0
            state["retry_count"] = 0
            state["max_retries"] = self.max_retries
            state["error_message"] = ""
            state["process"] = None
            
            logger.info("Automation initialized successfully")
            return state
            
        except Exception as e:
            state["error_message"] = f"Failed to initialize automation: {str(e)}"
            logger.error(state["error_message"])
            return state
    
    def launch_application(self, state: AutomationState) -> AutomationState:
        """Launch the target application"""
        exe_path = state["exe_path"]
        logger.info(f"Launching application: {exe_path}")
        
        try:
            # Launch the application
            process = subprocess.Popen([exe_path])
            state["process"] = process
            
            # Wait for application to launch
            time.sleep(3)
            
            logger.info("Application launched successfully")
            return state
            
        except Exception as e:
            state["error_message"] = f"Failed to launch application: {str(e)}"
            logger.error(state["error_message"])
            return state
    
    def process_step(self, state: AutomationState) -> AutomationState:
        """Process the current automation step"""
        current_step = state["current_step"]
        steps = state["steps"]
        
        if current_step < len(steps):
            step_data = steps[current_step]
            logger.info(f"Processing step {current_step + 1}/{len(steps)}: {step_data['field_name']}")
            state["current_step_data"] = step_data
            state["retry_count"] = 0
        
        return state
    
    def find_element(self, state: AutomationState) -> AutomationState:
        """Find element using text search or OCR"""
        step_data = state["current_step_data"]
        field_name = step_data["field_name"]
        
        logger.info(f"Searching for element: {field_name}")
        
        try:
            # First try: Direct text search using pyautogui
            element_coords = self._find_element_by_text(field_name)
            
            if element_coords:
                state["element_coords"] = element_coords
                state["element_found"] = True
                logger.info(f"Element found using text search at coordinates: {element_coords}")
                return state
            
            # Second try: OCR-based search
            logger.info("Text search failed, trying OCR...")
            element_coords = self._find_element_by_ocr(field_name, state["ocr_reader"])
            
            if element_coords:
                state["element_coords"] = element_coords
                state["element_found"] = True
                logger.info(f"Element found using OCR at coordinates: {element_coords}")
                return state
            
            # Element not found
            state["element_found"] = False
            state["error_message"] = f"Element '{field_name}' not found using text search or OCR"
            logger.error(state["error_message"])
            return state
            
        except Exception as e:
            state["element_found"] = False
            state["error_message"] = f"Error finding element '{field_name}': {str(e)}"
            logger.error(state["error_message"])
            return state
    
    def _find_element_by_text(self, text: str) -> Optional[Tuple[int, int]]:
        """Try to find element using pyautogui's text search capabilities"""
        try:
            # Take screenshot and search for text
            screenshot = pyautogui.screenshot()
            
            # Simple approach: try to locate text on screen
            # Note: This is a basic implementation - pyautogui doesn't have built-in text search
            # In a real implementation, you might want to use image templates or other methods
            return None
            
        except Exception as e:
            logger.warning(f"Text search failed: {str(e)}")
            return None
    
    def _find_element_by_ocr(self, text: str, ocr_reader) -> Optional[Tuple[int, int]]:
        """Find element using OCR"""
        try:
            # Take screenshot
            screenshot = pyautogui.screenshot()
            screenshot_np = np.array(screenshot)
            
            # Use OCR to find text
            results = ocr_reader.readtext(screenshot_np)
            
            # Search for matching text
            for (bbox, detected_text, confidence) in results:
                if confidence > 0.5 and text.lower() in detected_text.lower():
                    # Calculate center of bounding box
                    x_coords = [point[0] for point in bbox]
                    y_coords = [point[1] for point in bbox]
                    center_x = int(sum(x_coords) / len(x_coords))
                    center_y = int(sum(y_coords) / len(y_coords))
                    
                    return (center_x, center_y)
            
            return None
            
        except Exception as e:
            logger.warning(f"OCR search failed: {str(e)}")
            return None
    
    def execute_action(self, state: AutomationState) -> AutomationState:
        """Execute the specified action"""
        step_data = state["current_step_data"]
        element_coords = state["element_coords"]
        action_type = step_data["event_action_type"]
        data = step_data.get("data", "")
        
        logger.info(f"Executing action: {action_type} at coordinates: {element_coords}")
        
        try:
            if action_type == "windowActivate":
                self._activate_window(step_data["field_name"])
            elif action_type == "click":
                pyautogui.click(element_coords[0], element_coords[1])
            elif action_type == "typeInto":
                pyautogui.click(element_coords[0], element_coords[1])
                time.sleep(0.5)
                pyautogui.typewrite(data)
            else:
                raise ValueError(f"Unsupported action type: {action_type}")
            
            # Wait after action
            time.sleep(1)
            
            state["action_success"] = True
            state["current_step"] += 1
            logger.info(f"Action {action_type} executed successfully")
            return state
            
        except Exception as e:
            state["action_success"] = False
            state["error_message"] = f"Action execution failed: {str(e)}"
            logger.error(state["error_message"])
            return state
    
    def _activate_window(self, window_title: str):
        """Activate window by title"""
        try:
            import win32gui
            import win32con
            
            def enum_windows_proc(hwnd, lParam):
                if win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd)
                    if window_title.lower() in window_text.lower():
                        win32gui.SetForegroundWindow(hwnd)
                        return False
                return True
            
            win32gui.EnumWindows(enum_windows_proc, None)
            
        except ImportError:
            # Fallback: try to click on the window if we can find it
            logger.warning("win32gui not available, using fallback method")
            time.sleep(0.5)
    
    def handle_error(self, state: AutomationState) -> AutomationState:
        """Handle errors and retries"""
        error_message = state.get("error_message", "Unknown error")
        logger.error(f"Handling error: {error_message}")
        
        # Log final error and stop
        raise Exception(error_message)
    
    def complete_automation(self, state: AutomationState) -> AutomationState:
        """Complete the automation process"""
        logger.info("Automation completed successfully!")
        return state
    
    # Conditional edge functions
    def check_launch_success(self, state: AutomationState) -> str:
        return "error" if state.get("error_message") else "success"
    
    def check_steps_remaining(self, state: AutomationState) -> str:
        return "complete" if state["current_step"] >= len(state["steps"]) else "continue"
    
    def check_element_found(self, state: AutomationState) -> str:
        return "found" if state.get("element_found", False) else "not_found"
    
    def check_action_success(self, state: AutomationState) -> str:
        if state.get("action_success", False):
            return "success"
        elif state["retry_count"] < state["max_retries"]:
            state["retry_count"] += 1
            logger.info(f"Retrying action (attempt {state['retry_count']}/{state['max_retries']})")
            time.sleep(2)  # Pause before retry
            return "retry"
        else:
            return "error"
    
    def run_automation(self, exe_path: str, steps_json: List[Dict[str, Any]]) -> None:
        """Run the automation process"""
        logger.info("Starting desktop automation...")
        
        # Prepare initial state
        initial_state = AutomationState(
            exe_path=exe_path,
            steps=steps_json,
            current_step=0,
            error_message="",
            retry_count=0,
            max_retries=self.max_retries,
            ocr_reader=None,
            process=None
        )
        
        try:
            # Run the graph
            final_state = self.graph.invoke(initial_state)
            logger.info("Automation completed successfully!")
            
        except Exception as e:
            logger.error(f"Automation failed: {str(e)}")
            raise