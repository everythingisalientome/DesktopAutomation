# desktop_automation_node.py
"""
Desktop Automation LangGraph Node - Complete Single File Implementation

A LangGraph-compatible node for automating Windows desktop applications using:
- UI Automation (primary)
- Win32 API (fallback)
- OpenAI LLM (intelligent fallback for hard-to-find elements)

Author: AI Assistant
Version: 1.0
"""

import json
import subprocess
import time
import win32gui
import win32con
import win32process
import win32api
import os
from typing import TypedDict, List, Dict, Any, Optional

# Optional imports with graceful fallback
try:
    import uiautomation as auto
    UI_AUTOMATION_AVAILABLE = True
    print("✅ UI Automation available")
except ImportError:
    UI_AUTOMATION_AVAILABLE = False
    print("⚠️  Warning: uiautomation not available, using Win32 API only")

try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
    print("✅ OpenAI available for LLM fallback")
except ImportError:
    OPENAI_AVAILABLE = False
    print("⚠️  Warning: OpenAI not available, LLM fallback disabled")


# =============================================================================
# STATE DEFINITION
# =============================================================================

class AutomationState(TypedDict):
    """LangGraph state for desktop automation"""
    exe_path: str
    automation_steps: List[Dict[str, Any]]
    automation_result: str
    automation_success: bool


# =============================================================================
# WINDOW MANAGER CLASS
# =============================================================================

class WindowManager:
    """Handles application launching and window management"""
    
    def __init__(self):
        self.launched_processes = []
    
    def launch_application(self, exe_path: str) -> Optional[subprocess.Popen]:
        """Launch the application if not already running"""
        try:
            if not os.path.exists(exe_path):
                print(f"❌ Executable not found: {exe_path}")
                return None
            
            print(f"🚀 Launching application: {exe_path}")
            process = subprocess.Popen(exe_path)
            self.launched_processes.append(process)
            
            # Wait for the process to start
            time.sleep(3)
            print(f"✅ Application launched successfully (PID: {process.pid})")
            return process
            
        except Exception as e:
            print(f"❌ Failed to launch application: {str(e)}")
            return None
    
    def activate_window(self, window_name: str) -> bool:
        """Activate window by name"""
        try:
            print(f"🪟 Searching for window: {window_name}")
            
            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if window_name.lower() in title.lower():
                        windows.append(hwnd)
                return True
            
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            if windows:
                hwnd = windows[0]
                win32gui.SetForegroundWindow(hwnd)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.5)
                print(f"✅ Activated window: {window_name}")
                return True
            else:
                print(f"❌ Window not found: {window_name}")
                return False
                
        except Exception as e:
            print(f"❌ Error activating window: {str(e)}")
            return False


# =============================================================================
# ELEMENT DETECTOR CLASS
# =============================================================================

class ElementDetector:
    """Three-tier element detection: UI Automation → Win32 API → LLM fallback"""
    
    def __init__(self, openai_api_key: Optional[str] = None, openai_model: str = "gpt-4"):
        self.timeout = 30
        self.openai_client = None
        self.openai_model = openai_model
        
        if OPENAI_AVAILABLE and openai_api_key:
            self.openai_client = OpenAI(api_key=openai_api_key)
            print("🤖 LLM fallback enabled with OpenAI")
        elif openai_api_key and not OPENAI_AVAILABLE:
            print("⚠️  Warning: OpenAI API key provided but openai package not installed")
    
    def find_element(self, window_name: str, field_name: str, field_type: str, timeout: int = 30) -> Optional[Any]:
        """Find element using three-tier detection strategy"""
        print(f"🔍 Looking for element: '{field_name}' (type: {field_type}) in window: '{window_name}'")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            # Tier 1: UI Automation (Primary)
            if UI_AUTOMATION_AVAILABLE:
                element = self._find_with_ui_automation(window_name, field_name, field_type)
                if element:
                    print(f"✅ Found element with UI Automation: {field_name}")
                    return element
            
            # Tier 2: Win32 API (Fallback)
            element = self._find_with_win32_api(window_name, field_name, field_type)
            if element:
                print(f"✅ Found element with Win32 API: {field_name}")
                return element
            
            time.sleep(0.5)  # Wait before retrying
        
        # Tier 3: LLM Fallback (Last resort)
        if self.openai_client and UI_AUTOMATION_AVAILABLE:
            print(f"🤖 Trying LLM fallback for element: {field_name}")
            element = self._find_with_llm_fallback(window_name, field_name, field_type)
            if element:
                print(f"✅ Found element with LLM fallback: {field_name}")
                return element
        
        print(f"❌ Element not found after {timeout} seconds: {field_name}")
        return None
    
    def _find_with_ui_automation(self, window_name: str, field_name: str, field_type: str) -> Optional[Any]:
        """Find element using UI Automation"""
        try:
            if not UI_AUTOMATION_AVAILABLE:
                return None
                
            # Find the main window - try different approaches
            window = None
            
            # Method 1: Direct name match
            try:
                window = auto.WindowControl(searchDepth=1, Name=window_name)
                if not window.Exists(0, False):
                    window = None
            except:
                pass
            
            # Method 2: Search through all windows if direct match failed
            if not window:
                try:
                    # Get desktop and search for window
                    desktop = auto.GetRootControl()
                    windows = desktop.GetChildren()
                    for w in windows:
                        try:
                            if (hasattr(w, 'Name') and w.Name and 
                                window_name.lower() in w.Name.lower() and
                                hasattr(w, 'ControlType') and
                                w.ControlType == auto.ControlType.WindowControl):
                                window = w
                                break
                        except:
                            continue
                except:
                    pass
            
            if not window:
                return None
            
            # Search strategies in order of preference
            # Strategy 1: Exact AutomationId match
            try:
                elements = window.GetChildren()
                for element in elements:
                    try:
                        if hasattr(element, 'AutomationId') and element.AutomationId == field_name:
                            return {'type': 'ui_automation', 'element': element}
                    except:
                        continue
                        
                # Search recursively in children
                def search_recursive(ctrl, target_name, max_depth=5, current_depth=0):
                    if current_depth >= max_depth:
                        return None
                    try:
                        if hasattr(ctrl, 'AutomationId') and ctrl.AutomationId == target_name:
                            return ctrl
                        if hasattr(ctrl, 'Name') and ctrl.Name == target_name:
                            return ctrl
                        
                        # Search children
                        children = ctrl.GetChildren()
                        for child in children:
                            result = search_recursive(child, target_name, max_depth, current_depth + 1)
                            if result:
                                return result
                    except:
                        pass
                    return None
                
                element = search_recursive(window, field_name)
                if element:
                    return {'type': 'ui_automation', 'element': element}
                    
            except Exception as e:
                print(f"   AutomationId search failed: {e}")
            
            # Strategy 2: Exact Name match
            try:
                def search_by_name(ctrl, target_name, max_depth=5, current_depth=0):
                    if current_depth >= max_depth:
                        return None
                    try:
                        if hasattr(ctrl, 'Name') and ctrl.Name == target_name:
                            return ctrl
                        
                        # Search children
                        children = ctrl.GetChildren()
                        for child in children:
                            result = search_by_name(child, target_name, max_depth, current_depth + 1)
                            if result:
                                return result
                    except:
                        pass
                    return None
                
                element = search_by_name(window, field_name)
                if element:
                    return {'type': 'ui_automation', 'element': element}
                    
            except Exception as e:
                print(f"   Name search failed: {e}")
            
            # Strategy 3: ClassName partial match
            if field_type:
                try:
                    def search_by_class(ctrl, target_name, max_depth=5, current_depth=0):
                        if current_depth >= max_depth:
                            return None
                        try:
                            if (hasattr(ctrl, 'ClassName') and ctrl.ClassName and 
                                target_name.lower() in ctrl.ClassName.lower()):
                                return ctrl
                            
                            # Search children
                            children = ctrl.GetChildren()
                            for child in children:
                                result = search_by_class(child, target_name, max_depth, current_depth + 1)
                                if result:
                                    return result
                        except:
                            pass
                        return None
                    
                    element = search_by_class(window, field_name)
                    if element:
                        return {'type': 'ui_automation', 'element': element}
                        
                except Exception as e:
                    print(f"   ClassName search failed: {e}")
            
            # Strategy 4: For Notepad specifically, try finding Edit controls
            if 'notepad' in window_name.lower():
                try:
                    def find_edit_control(ctrl, max_depth=5, current_depth=0):
                        if current_depth >= max_depth:
                            return None
                        try:
                            # Look for Edit or Document controls
                            if (hasattr(ctrl, 'ControlType') and 
                                (ctrl.ControlType == auto.ControlType.EditControl or
                                 ctrl.ControlType == auto.ControlType.DocumentControl)):
                                return ctrl
                            
                            # Search children
                            children = ctrl.GetChildren()
                            for child in children:
                                result = find_edit_control(child, max_depth, current_depth + 1)
                                if result:
                                    return result
                        except:
                            pass
                        return None
                    
                    element = find_edit_control(window)
                    if element:
                        return {'type': 'ui_automation', 'element': element}
                        
                except Exception as e:
                    print(f"   Edit control search failed: {e}")
            
            return None
            
        except Exception as e:
            print(f"⚠️  UI Automation search failed: {str(e)}")
            return None
    
    def _find_with_win32_api(self, window_name: str, field_name: str, field_type: str) -> Optional[Any]:
        """Find element using Win32 API"""
        try:
            # Find the main window
            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if window_name.lower() in title.lower():
                        windows.append(hwnd)
                return True
            
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            if not windows:
                return None
            
            main_hwnd = windows[0]
            
            # Find child window/control
            element_hwnd = self._find_child_window(main_hwnd, field_name)
            
            if element_hwnd:
                return {'type': 'win32', 'hwnd': element_hwnd}
            
            return None
            
        except Exception as e:
            print(f"⚠️  Win32 API search failed: {str(e)}")
            return None
    
    def _find_child_window(self, parent_hwnd: int, field_name: str) -> Optional[int]:
        """Find child window by various properties"""
        try:
            found_hwnd = None
            
            def enum_child_callback(hwnd, param):
                nonlocal found_hwnd
                
                # Try by window text
                text = win32gui.GetWindowText(hwnd)
                if text and field_name.lower() in text.lower():
                    found_hwnd = hwnd
                    return False
                
                # Try by class name
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name and field_name.lower() in class_name.lower():
                        found_hwnd = hwnd
                        return False
                except:
                    pass
                
                return True
            
            win32gui.EnumChildWindows(parent_hwnd, enum_child_callback, None)
            return found_hwnd
            
        except Exception as e:
            print(f"⚠️  Child window search failed: {str(e)}")
            return None
    
    def _find_with_llm_fallback(self, window_name: str, field_name: str, field_type: str) -> Optional[Any]:
        """Use LLM to analyze all UI elements and find the best match"""
        try:
            if not self.openai_client or not UI_AUTOMATION_AVAILABLE:
                return None
            
            print("📊 Collecting all UI elements for LLM analysis...")
            
            # Get all UI elements from the window
            all_elements = self._collect_all_ui_elements(window_name)
            if not all_elements:
                print("❌ No UI elements found for LLM analysis")
                return None
            
            # Prepare the LLM prompt
            prompt = self._create_llm_prompt(field_name, field_type, all_elements)
            
            # Get LLM response
            llm_response = self._query_llm(prompt)
            if not llm_response:
                return None
            
            # Parse LLM response and find the element
            return self._find_element_from_llm_response(llm_response, window_name)
            
        except Exception as e:
            print(f"❌ LLM fallback failed: {str(e)}")
            return None
    
    def _collect_all_ui_elements(self, window_name: str) -> List[Dict[str, Any]]:
        """Collect all UI elements from the window with their properties"""
        try:
            # Find the main window
            window = None
            
            # Try direct name match first
            try:
                window = auto.WindowControl(searchDepth=1, Name=window_name)
                if not window.Exists(0, False):
                    window = None
            except:
                pass
            
            # If direct match failed, search through all windows
            if not window:
                try:
                    desktop = auto.GetRootControl()
                    windows = desktop.GetChildren()
                    for w in windows:
                        try:
                            if (hasattr(w, 'Name') and w.Name and 
                                window_name.lower() in w.Name.lower() and
                                hasattr(w, 'ControlType') and
                                w.ControlType == auto.ControlType.WindowControl):
                                window = w
                                break
                        except:
                            continue
                except:
                    pass
            
            if not window:
                return []
            
            elements = []
            
            def collect_element_info(ctrl, depth):
                try:
                    # Limit depth to avoid infinite recursion
                    if depth > 10:
                        return
                    
                    # Skip invisible or problematic elements
                    try:
                        if not hasattr(ctrl, 'IsEnabled') or not ctrl.IsEnabled:
                            return
                        
                        if hasattr(ctrl, 'BoundingRectangle') and ctrl.BoundingRectangle:
                            rect = ctrl.BoundingRectangle
                            if rect.width() < 5 or rect.height() < 5:
                                return
                        else:
                            # If no bounding rectangle, skip size check
                            pass
                    except:
                        # If we can't check properties, skip this element
                        return
                    
                    element_info = {
                        'name': getattr(ctrl, 'Name', ''),
                        'automation_id': getattr(ctrl, 'AutomationId', ''),
                        'class_name': getattr(ctrl, 'ClassName', ''),
                        'control_type': str(getattr(ctrl, 'ControlType', '')),
                        'localized_control_type': getattr(ctrl, 'LocalizedControlType', ''),
                        'is_enabled': getattr(ctrl, 'IsEnabled', False),
                        'is_visible': getattr(ctrl, 'IsVisible', False),
                        'help_text': getattr(ctrl, 'HelpText', ''),
                        'access_key': getattr(ctrl, 'AccessKey', ''),
                    }
                    
                    # Try to get bounds, but don't fail if we can't
                    try:
                        if hasattr(ctrl, 'BoundingRectangle') and ctrl.BoundingRectangle:
                            rect = ctrl.BoundingRectangle
                            element_info['bounds'] = {
                                'x': rect.left,
                                'y': rect.top,
                                'width': rect.width(),
                                'height': rect.height()
                            }
                        else:
                            element_info['bounds'] = {'x': 0, 'y': 0, 'width': 0, 'height': 0}
                    except:
                        element_info['bounds'] = {'x': 0, 'y': 0, 'width': 0, 'height': 0}
                    
                    # Only add elements with some identifying information
                    if (element_info['name'] or 
                        element_info['automation_id'] or 
                        element_info['class_name']):
                        elements.append(element_info)
                    
                    # Recursively collect child elements
                    try:
                        children = ctrl.GetChildren()
                        for child in children:
                            collect_element_info(child, depth + 1)
                    except:
                        pass
                        
                except Exception:
                    # Skip problematic elements
                    pass
            
            # Start collection from the main window
            collect_element_info(window, 0)
            
            print(f"📊 Collected {len(elements)} UI elements for LLM analysis")
            return elements
            
        except Exception as e:
            print(f"❌ Failed to collect UI elements: {str(e)}")
            return []
    
    def _create_llm_prompt(self, field_name: str, field_type: str, elements: List[Dict[str, Any]]) -> str:
        """Create a prompt for the LLM to analyze elements"""
        elements_json = json.dumps(elements, indent=2)
        
        prompt = f"""I need to find a UI element in a Windows desktop application. Here are the details:

TARGET ELEMENT:
- Field Name: "{field_name}"
- Field Type: "{field_type}"

AVAILABLE UI ELEMENTS:
{elements_json}

TASK:
Analyze all the UI elements and identify which one best matches the target element "{field_name}" of type "{field_type}".

Consider these matching strategies:
1. Exact match on 'name', 'automation_id', or 'class_name'
2. Partial match (case-insensitive) on any text property
3. Control type compatibility (e.g., "editable text" → EditControl, TextBox)
4. Contextual similarity (e.g., "btn" prefix for buttons, "txt" for text fields)

RESPONSE FORMAT:
Return ONLY a JSON object with the exact element properties that match, like this:
{{
    "matched_element": {{
        "name": "exact_name_from_list",
        "automation_id": "exact_automation_id_from_list", 
        "class_name": "exact_class_name_from_list",
        "control_type": "exact_control_type_from_list"
    }},
    "confidence": 0.95,
    "reasoning": "Brief explanation of why this element matches"
}}

If no good match is found, return:
{{
    "matched_element": null,
    "confidence": 0.0,
    "reasoning": "No suitable match found"
}}

IMPORTANT: Only return valid JSON. Do not include any other text or explanation outside the JSON."""
        return prompt
    
    def _query_llm(self, prompt: str) -> Optional[Dict[str, Any]]:
        """Query the LLM and parse the response"""
        try:
            response = self.openai_client.chat.completions.create(
                model=self.openai_model,
                messages=[
                    {
                        "role": "system", 
                        "content": "You are an expert at analyzing Windows UI elements. Return only valid JSON responses."
                    },
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,  # Low temperature for consistent results
                max_tokens=1000
            )
            
            response_text = response.choices[0].message.content.strip()
            
            # Try to parse the JSON response
            try:
                llm_result = json.loads(response_text)
                return llm_result
            except json.JSONDecodeError:
                print(f"❌ LLM returned invalid JSON: {response_text}")
                return None
                
        except Exception as e:
            print(f"❌ LLM query failed: {str(e)}")
            return None
    
    def _find_element_from_llm_response(self, llm_response: Dict[str, Any], window_name: str) -> Optional[Any]:
        """Find the actual UI element based on LLM response"""
        try:
            matched_element = llm_response.get('matched_element')
            confidence = llm_response.get('confidence', 0.0)
            reasoning = llm_response.get('reasoning', '')
            
            print(f"🤖 LLM Response - Confidence: {confidence:.2f}, Reasoning: {reasoning}")
            
            if not matched_element or confidence < 0.5:
                print("❌ LLM confidence too low or no match found")
                return None
            
            # Find the window again
            window = None
            try:
                window = auto.WindowControl(searchDepth=1, Name=window_name)
                if not window.Exists(0, False):
                    window = None
            except:
                pass
            
            if not window:
                try:
                    desktop = auto.GetRootControl()
                    windows = desktop.GetChildren()
                    for w in windows:
                        try:
                            if (hasattr(w, 'Name') and w.Name and 
                                window_name.lower() in w.Name.lower() and
                                hasattr(w, 'ControlType') and
                                w.ControlType == auto.ControlType.WindowControl):
                                window = w
                                break
                        except:
                            continue
                except:
                    pass
            
            if not window:
                return None
            
            # Search for the element using LLM-provided properties
            name = matched_element.get('name', '')
            automation_id = matched_element.get('automation_id', '')
            class_name = matched_element.get('class_name', '')
            
            # Search function that works without FindFirst
            def search_element(ctrl, target_name, target_automation_id, target_class, max_depth=5, current_depth=0):
                if current_depth >= max_depth:
                    return None
                try:
                    # Check if this control matches
                    if target_automation_id and hasattr(ctrl, 'AutomationId') and ctrl.AutomationId == target_automation_id:
                        return ctrl
                    if target_name and hasattr(ctrl, 'Name') and ctrl.Name == target_name:
                        return ctrl
                    if target_class and hasattr(ctrl, 'ClassName') and ctrl.ClassName == target_class:
                        return ctrl
                    
                    # Search children
                    children = ctrl.GetChildren()
                    for child in children:
                        result = search_element(child, target_name, target_automation_id, target_class, max_depth, current_depth + 1)
                        if result:
                            return result
                except:
                    pass
                return None
            
            element = search_element(window, name, automation_id, class_name)
            if element:
                print(f"✅ Found element using LLM guidance: {name or automation_id or class_name}")
                return {'type': 'ui_automation', 'element': element}
            
            print("❌ Could not find element despite LLM guidance")
            return None
            
        except Exception as e:
            print(f"❌ Failed to find element from LLM response: {str(e)}")
            return None


# =============================================================================
# ACTION EXECUTOR CLASS
# =============================================================================

class ActionExecutor:
    """Executes automation actions on UI elements"""
    
    def __init__(self):
        # Special key mappings for Win32 API
        self.special_keys = {
            'enter': win32con.VK_RETURN,
            'tab': win32con.VK_TAB,
            'escape': win32con.VK_ESCAPE,
            'space': win32con.VK_SPACE,
            'backspace': win32con.VK_BACK,
            'delete': win32con.VK_DELETE,
            'home': win32con.VK_HOME,
            'end': win32con.VK_END,
            'pageup': win32con.VK_PRIOR,
            'pagedown': win32con.VK_NEXT,
            'up': win32con.VK_UP,
            'down': win32con.VK_DOWN,
            'left': win32con.VK_LEFT,
            'right': win32con.VK_RIGHT,
        }
        
        # Function keys F1-F12
        for i in range(1, 13):
            self.special_keys[f'f{i}'] = win32con.VK_F1 + i - 1
    
    def execute_action(self, element: Dict[str, Any], action_type: str, data: str = '', special_key: str = '') -> bool:
        """Execute the specified action on the element"""
        try:
            action_type = action_type.lower()
            print(f"⚡ Executing action: {action_type}")
            
            if action_type == 'click':
                return self._click_element(element)
            elif action_type == 'typeinto':
                return self._type_into_element(element, data)
            elif action_type == 'keypress':
                return self._key_press(element, special_key)
            elif action_type == 'select':
                return self._select_element(element, data)
            else:
                print(f"❌ Unknown action type: {action_type}")
                return False
                
        except Exception as e:
            print(f"❌ Action execution failed: {str(e)}")
            return False
    
    def _click_element(self, element: Dict[str, Any]) -> bool:
        """Click on the element"""
        try:
            if element['type'] == 'ui_automation':
                ui_element = element['element']
                ui_element.Click()
                print(f"🖱️  Clicked element via UI Automation")
                return True
                
            elif element['type'] == 'win32':
                hwnd = element['hwnd']
                # Get element position and click
                rect = win32gui.GetWindowRect(hwnd)
                x = (rect[0] + rect[2]) // 2
                y = (rect[1] + rect[3]) // 2
                
                # Send click message
                win32gui.SetForegroundWindow(hwnd)
                win32api.SetCursorPos((x, y))
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
                print(f"🖱️  Clicked element via Win32 API at ({x}, {y})")
                return True
            
            return False
            
        except Exception as e:
            print(f"❌ Click failed: {str(e)}")
            return False
    
    def _type_into_element(self, element: Dict[str, Any], text: str) -> bool:
        """Type text into the element"""
        try:
            # First click on the element to focus it
            self._click_element(element)
            time.sleep(0.2)
            
            if element['type'] == 'ui_automation':
                ui_element = element['element']
                # Clear existing text and type new text
                ui_element.SendKeys('{Ctrl}a')
                time.sleep(0.1)
                ui_element.SendKeys(text)
                print(f"⌨️  Typed '{text}' via UI Automation")
                return True
                
            elif element['type'] == 'win32':
                hwnd = element['hwnd']
                win32gui.SetForegroundWindow(hwnd)
                
                # Clear existing text
                win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
                win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord('a'), 0)
                win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
                time.sleep(0.1)
                
                # Type new text
                for char in text:
                    win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                    time.sleep(0.01)
                
                print(f"⌨️  Typed '{text}' via Win32 API")
                return True
            
            return False
            
        except Exception as e:
            print(f"❌ Type into failed: {str(e)}")
            return False
    
    def _key_press(self, element: Dict[str, Any], key_combination: str) -> bool:
        """Press special keys or key combinations"""
        try:
            # First focus the element
            self._click_element(element)
            time.sleep(0.1)
            
            # Parse key combination
            keys = key_combination.lower().split('+')
            modifier_keys = []
            main_key = None
            
            for key in keys:
                key = key.strip()
                if key == 'ctrl':
                    modifier_keys.append(win32con.VK_CONTROL)
                elif key == 'alt':
                    modifier_keys.append(win32con.VK_MENU)
                elif key == 'shift':
                    modifier_keys.append(win32con.VK_SHIFT)
                else:
                    main_key = key
            
            if not main_key:
                print(f"❌ No main key found in combination: {key_combination}")
                return False
            
            # Get the virtual key code
            if main_key in self.special_keys:
                vk_code = self.special_keys[main_key]
            elif len(main_key) == 1:
                vk_code = ord(main_key.upper())
            else:
                print(f"❌ Unknown key: {main_key}")
                return False
            
            # Send key combination
            if element['type'] == 'ui_automation':
                ui_element = element['element']
                key_string = self._build_ui_automation_key_string(keys)
                ui_element.SendKeys(key_string)
                print(f"🔑 Pressed key combination '{key_combination}' via UI Automation")
                return True
                
            elif element['type'] == 'win32':
                hwnd = element['hwnd']
                win32gui.SetForegroundWindow(hwnd)
                
                # Press modifier keys
                for mod_key in modifier_keys:
                    win32api.keybd_event(mod_key, 0, 0, 0)
                
                # Press main key
                win32api.keybd_event(vk_code, 0, 0, 0)
                win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                
                # Release modifier keys
                for mod_key in reversed(modifier_keys):
                    win32api.keybd_event(mod_key, 0, win32con.KEYEVENTF_KEYUP, 0)
                
                print(f"🔑 Pressed key combination '{key_combination}' via Win32 API")
                return True
            
            return False
            
        except Exception as e:
            print(f"❌ Key press failed: {str(e)}")
            return False
    
    def _select_element(self, element: Dict[str, Any], value: str) -> bool:
        """Select an item from a dropdown or list"""
        try:
            if element['type'] == 'ui_automation':
                ui_element = element['element']
                # Try to expand if it's a combobox
                try:
                    ui_element.Expand()
                    time.sleep(0.2)
                except:
                    pass
                
                # Find and select the item
                item = ui_element.FindFirst(lambda ctrl, depth: 
                    hasattr(ctrl, 'Name') and ctrl.Name.lower() == value.lower()
                )
                
                if item:
                    item.Select()
                    print(f"📋 Selected '{value}' from dropdown via UI Automation")
                    return True
                else:
                    # Try clicking and typing
                    ui_element.Click()
                    time.sleep(0.2)
                    ui_element.SendKeys(value)
                    time.sleep(0.2)
                    ui_element.SendKeys('{Enter}')
                    print(f"📋 Selected '{value}' by typing via UI Automation")
                    return True
                    
            elif element['type'] == 'win32':
                hwnd = element['hwnd']
                # Click on the element first
                self._click_element(element)
                time.sleep(0.2)
                
                # Try sending the text directly
                for char in value:
                    win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                    time.sleep(0.01)
                
                # Press Enter to confirm
                win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                print(f"📋 Selected '{value}' via Win32 API")
                return True
            
            return False
            
        except Exception as e:
            print(f"❌ Select failed: {str(e)}")
            return False
    
    def _build_ui_automation_key_string(self, keys):
        """Build UI Automation key string from key combination"""
        key_map = {
            'ctrl': '{Ctrl}',
            'alt': '{Alt}',
            'shift': '{Shift}',
            'enter': '{Enter}',
            'tab': '{Tab}',
            'escape': '{Esc}',
            'space': ' ',
            'backspace': '{Backspace}',
            'delete': '{Delete}',
        }
        
        # Handle function keys
        for i in range(1, 13):
            key_map[f'f{i}'] = f'{{F{i}}}'
        
        if len(keys) == 1:
            key = keys[0].strip().lower()
            return key_map.get(key, key)
        else:
            # Handle combinations like Ctrl+A
            result = ''
            for key in keys[:-1]:  # All but last are modifiers
                key = key.strip().lower()
                if key in ['ctrl', 'alt', 'shift']:
                    result += key_map[key]
            
            # Add the main key
            main_key = keys[-1].strip().lower()
            result += key_map.get(main_key, main_key)
            
            return result


# =============================================================================
# MAIN DESKTOP AUTOMATION NODE
# =============================================================================

class DesktopAutomationNode:
    """
    Main LangGraph node for desktop automation
    
    Integrates WindowManager, ElementDetector, and ActionExecutor
    to provide complete desktop automation capabilities.
    """
    
    def __init__(self, openai_api_key: Optional[str] = None, openai_model: str = "gpt-4"):
        """
        Initialize the desktop automation node
        
        Args:
            openai_api_key: Optional OpenAI API key for LLM fallback
            openai_model: OpenAI model to use (default: gpt-4)
        """
        print("🚀 Initializing Desktop Automation Node...")
        self.window_manager = WindowManager()
        self.element_detector = ElementDetector(openai_api_key, openai_model)
        self.action_executor = ActionExecutor()
        print("✅ Desktop Automation Node ready!")
        
    def execute_automation(self, state: AutomationState) -> AutomationState:
        """
        Main LangGraph node function that executes desktop automation
        
        Args:
            state: AutomationState containing exe_path and automation_steps
            
        Returns:
            Updated AutomationState with results
        """
        print("=" * 80)
        print(f"🖥️  STARTING DESKTOP AUTOMATION")
        print("=" * 80)
        print(f"📁 Executable: {state['exe_path']}")
        print(f"📋 Total steps: {len(state['automation_steps'])}")
        print("=" * 80)
        
        try:
            # Step 1: Launch the executable
            process = self.window_manager.launch_application(state['exe_path'])
            if not process:
                error_msg = f"Failed to launch application: {state['exe_path']}"
                print(f"❌ {error_msg}")
                return {
                    **state,
                    'automation_result': error_msg,
                    'automation_success': False
                }
            
            print(f"✅ Application launched successfully")
            time.sleep(2)  # Give app time to fully load
            
            # Step 2: Process each automation step
            processed_steps = 0
            skipped_steps = 0
            
            for i, step in enumerate(state['automation_steps']):
                step_num = i + 1
                action_type = step.get('event_action_type', 'unknown')
                field_name = step.get('field_name', '')
                
                print(f"\n📍 Step {step_num}/{len(state['automation_steps'])}: {action_type} on '{field_name}'")
                
                # Skip steps with field_name = "nan"
                if field_name.lower() == 'nan':
                    print(f"⏭️  Skipping step {step_num} - field_name is 'nan'")
                    skipped_steps += 1
                    continue
                
                # Execute the step with retry logic
                success = self._execute_step_with_retry(step, step_num)
                if not success:
                    error_msg = f"Failed to execute step {step_num}: {step}"
                    print(f"❌ {error_msg}")
                    print(f"\n📊 AUTOMATION SUMMARY:")
                    print(f"   ✅ Processed: {processed_steps}")
                    print(f"   ⏭️  Skipped: {skipped_steps}")
                    print(f"   ❌ Failed at step: {step_num}")
                    return {
                        **state,
                        'automation_result': error_msg,
                        'automation_success': False
                    }
                
                processed_steps += 1
                time.sleep(0.5)  # Small delay between steps
            
            # Success!
            success_msg = f"Desktop automation completed successfully! Processed: {processed_steps}, Skipped: {skipped_steps}"
            print(f"\n🎉 {success_msg}")
            print("=" * 80)
            
            return {
                **state,
                'automation_result': success_msg,
                'automation_success': True
            }
            
        except Exception as e:
            error_msg = f"Automation failed with exception: {str(e)}"
            print(f"\n💥 {error_msg}")
            print("=" * 80)
            return {
                **state,
                'automation_result': error_msg,
                'automation_success': False
            }
    
    def _execute_step_with_retry(self, step: Dict[str, Any], step_num: int) -> bool:
        """Execute a single step with one retry on failure"""
        for attempt in range(2):  # Original attempt + 1 retry
            try:
                if attempt > 0:
                    print(f"🔄 Retry attempt {attempt} for step {step_num}")
                
                success = self._execute_single_step(step)
                if success:
                    if attempt > 0:
                        print(f"✅ Step {step_num} succeeded on retry")
                    else:
                        print(f"✅ Step {step_num} completed successfully")
                    return True
                    
            except Exception as e:
                print(f"⚠️  Step execution error (attempt {attempt + 1}): {str(e)}")
                
            if attempt == 0:
                time.sleep(1)  # Wait before retry
        
        print(f"❌ Step {step_num} failed after retry")
        return False
    
    def _execute_single_step(self, step: Dict[str, Any]) -> bool:
        """Execute a single automation step"""
        action_type = step.get('event_action_type', '').lower()
        window_name = step.get('window_name', '')
        field_name = step.get('field_name', '')
        
        # Special case: window activation
        if action_type == 'windowactivate':
            return self.window_manager.activate_window(window_name or field_name)
        
        # For other actions, find the element first
        element = self.element_detector.find_element(
            window_name=window_name,
            field_name=field_name,
            field_type=step.get('field_type', ''),
            timeout=30
        )
        
        if not element:
            print(f"❌ Element not found: {field_name}")
            return False
        
        # Execute the action
        return self.action_executor.execute_action(
            element=element,
            action_type=action_type,
            data=step.get('Data', ''),
            special_key=step.get('SpecialKeyWithData', '')
        )


# =============================================================================
# CONVENIENCE FUNCTIONS AND EXPORTS
# =============================================================================

def create_automation_node(openai_api_key: Optional[str] = None, openai_model: str = "gpt-4") -> DesktopAutomationNode:
    """
    Convenience function to create a desktop automation node
    
    Args:
        openai_api_key: Optional OpenAI API key for LLM fallback
        openai_model: OpenAI model to use (default: gpt-4)
        
    Returns:
        Configured DesktopAutomationNode instance
    """
    return DesktopAutomationNode(openai_api_key, openai_model)


def create_sample_state(exe_path: str, automation_steps: List[Dict[str, Any]]) -> AutomationState:
    """
    Convenience function to create automation state
    
    Args:
        exe_path: Path to the executable
        automation_steps: List of automation step dictionaries
        
    Returns:
        AutomationState ready for execution
    """
    return {
        'exe_path': exe_path,
        'automation_steps': automation_steps,
        'automation_result': '',
        'automation_success': False
    }


# Export the main classes and functions
__all__ = [
    'DesktopAutomationNode',
    'AutomationState', 
    'WindowManager',
    'ElementDetector',
    'ActionExecutor',
    'create_automation_node',
    'create_sample_state'
]


# =============================================================================
# DEMO AND TESTING CODE
# =============================================================================

if __name__ == "__main__":
    """
    Demo script to test the desktop automation node
    """
    print("🖥️  Desktop Automation Node - Single File Implementation")
    print("=" * 80)
    
    # Sample automation steps (your original data)
    sample_steps = [
        {
            "window_name": "Settlement Request Test Application",
            "field_name": "Settlement Request Test Application",
            "field_type": "client",
            "event_action_type": "windowActivate",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "nan",
            "field_name": "nan",
            "field_type": "title bar",
            "event_action_type": "Click",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Settlement Request Test Application",
            "field_name": "txtECNNumber",
            "field_type": "editable text",
            "event_action_type": "typeInto",
            "SpecialKeyWithData": "",
            "Data": "2154785568"
        },
        {
            "window_name": "Settlement Request Test Application",
            "field_name": "txtECNNumber",
            "field_type": "editable text",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "Tab",
            "Data": ""
        },
        {
            "window_name": "Settlement Request Test Application",
            "field_name": "btnLoadInfo",
            "field_type": "push button",
            "event_action_type": "click",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Settlement Request Test Application",
            "field_name": "cmbSettlementOffer",
            "field_type": "list item",
            "event_action_type": "select",
            "SpecialKeyWithData": "",
            "Data": "review"
        }
    ]
    
    # Simple demo with Notepad (more accessible for testing)
    notepad_steps = [
        {
            "window_name": "Untitled - Notepad",
            "field_name": "Untitled - Notepad",
            "field_type": "window",
            "event_action_type": "windowActivate",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Untitled - Notepad",
            "field_name": "Text Editor",
            "field_type": "editable text",
            "event_action_type": "typeInto",
            "SpecialKeyWithData": "",
            "Data": "Hello from Desktop Automation Node!"
        }
    ]
    
    print("Choose a test option:")
    print("1. Test with your Settlement Request application")
    print("2. Test with Notepad (recommended for initial testing)")
    print("3. Exit")
    
    choice = input("\nEnter choice (1-3): ").strip()
    
    if choice == "1":
        exe_path = input("Enter full path to Settlement Request executable: ").strip()
        if not exe_path:
            exe_path = r"C:\Path\To\SettlementRequest.exe"  # Default
            
        # Check for OpenAI API key
        openai_key = os.getenv('OPENAI_API_KEY') or input("Enter OpenAI API key (optional): ").strip()
        
        # Create state and node
        state = create_sample_state(exe_path, sample_steps)
        node = create_automation_node(openai_key if openai_key else None)
        
        # Execute automation
        result = node.execute_automation(state)
        
        print(f"\n🏁 FINAL RESULT:")
        print(f"   Success: {result['automation_success']}")
        print(f"   Message: {result['automation_result']}")
        
    elif choice == "2":
        # Test with Notepad
        state = create_sample_state(r'C:\Windows\System32\notepad.exe', notepad_steps)
        node = create_automation_node()  # No LLM needed for simple test
        
        result = node.execute_automation(state)
        
        print(f"\n🏁 FINAL RESULT:")
        print(f"   Success: {result['automation_success']}")
        print(f"   Message: {result['automation_result']}")
        
    elif choice == "3":
        print("👋 Goodbye!")
        
    else:
        print("❌ Invalid choice!")
        
    print("=" * 80)