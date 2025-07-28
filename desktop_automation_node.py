# desktop_automation_node.py
"""
Desktop Automation LangGraph Node - Complete Clean Implementation

A LangGraph-compatible node for automating Windows desktop applications using:
- UI Automation (primary)
- Win32 API (fallback) 
- OpenAI LLM (intelligent fallback for hard-to-find elements)
- Full AccessibleName support for better element detection

Author: AI Assistant
Version: 2.0 - Clean Rewrite with AccessibleName Support
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
    """Three-tier element detection with full AccessibleName support"""
    
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
        """Find element using three-tier detection strategy with accessibility support"""
        print(f"🔍 Looking for element: '{field_name}' (type: {field_type}) in window: '{window_name}'")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            # Tier 1: UI Automation (Primary) - with accessibility support
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
        
        # Tier 3: LLM Fallback (Last resort) - with accessibility support
        if self.openai_client and UI_AUTOMATION_AVAILABLE:
            print(f"🤖 Trying LLM fallback for element: {field_name}")
            element = self._find_with_llm_fallback(window_name, field_name, field_type)
            if element:
                print(f"✅ Found element with LLM fallback: {field_name}")
                return element
        
        print(f"❌ Element not found after {timeout} seconds: {field_name}")
        return None
    
    def _find_with_ui_automation(self, window_name: str, field_name: str, field_type: str) -> Optional[Any]:
        """Find element using UI Automation with full accessibility support"""
        try:
            if not UI_AUTOMATION_AVAILABLE:
                return None
            
            # Find the main window
            window = self._find_window(window_name)
            if not window:
                return None
            
            # Comprehensive search with ALL accessibility properties
            element = self._search_element_comprehensive(window, field_name)
            if element:
                return {'type': 'ui_automation', 'element': element}
            
            return None
            
        except Exception as e:
            print(f"⚠️  UI Automation search failed: {str(e)}")
            return None
    
    def _find_window(self, window_name: str):
        """Find window using multiple strategies"""
        # Method 1: Direct name match
        try:
            window = auto.WindowControl(searchDepth=1, Name=window_name)
            if window.Exists(0, False):
                return window
        except:
            pass
        
        # Method 2: Search through all windows
        try:
            desktop = auto.GetRootControl()
            windows = desktop.GetChildren()
            for w in windows:
                try:
                    if (hasattr(w, 'Name') and w.Name and 
                        window_name.lower() in w.Name.lower() and
                        hasattr(w, 'ControlType') and
                        w.ControlType == auto.ControlType.WindowControl):
                        return w
                except:
                    continue
        except:
            pass
        
        return None
    
    def _search_element_comprehensive(self, window, field_name: str, max_depth: int = 5):
        """Comprehensive element search with ALL accessibility properties"""
        def search_recursive(ctrl, target_name, current_depth=0):
            if current_depth >= max_depth:
                return None
            
            try:
                # Standard property checks
                if hasattr(ctrl, 'AutomationId') and ctrl.AutomationId == target_name:
                    return ctrl
                if hasattr(ctrl, 'Name') and ctrl.Name == target_name:
                    return ctrl
                
                # Accessibility property checks - THE KEY ENHANCEMENT!
                if hasattr(ctrl, 'AccessibleName') and ctrl.AccessibleName == target_name:
                    return ctrl
                
                # Advanced accessibility checks via GetPropertyValue
                try:
                    name_prop = ctrl.GetPropertyValue(auto.PropertyId.NameProperty)
                    if name_prop and str(name_prop) == target_name:
                        return ctrl
                except:
                    pass
                
                try:
                    legacy_name = ctrl.GetPropertyValue(auto.PropertyId.LegacyIAccessibleNameProperty)
                    if legacy_name and str(legacy_name) == target_name:
                        return ctrl
                except:
                    pass
                
                try:
                    legacy_value = ctrl.GetPropertyValue(auto.PropertyId.LegacyIAccessibleValueProperty)
                    if legacy_value and str(legacy_value) == target_name:
                        return ctrl
                except:
                    pass
                
                # Partial matching for accessibility properties
                accessibility_values = []
                
                if hasattr(ctrl, 'Name') and ctrl.Name:
                    accessibility_values.append(ctrl.Name)
                if hasattr(ctrl, 'AccessibleName') and ctrl.AccessibleName:
                    accessibility_values.append(ctrl.AccessibleName)
                
                try:
                    name_prop = ctrl.GetPropertyValue(auto.PropertyId.NameProperty)
                    if name_prop:
                        accessibility_values.append(str(name_prop))
                except:
                    pass
                
                try:
                    legacy_name = ctrl.GetPropertyValue(auto.PropertyId.LegacyIAccessibleNameProperty)
                    if legacy_name:
                        accessibility_values.append(str(legacy_name))
                except:
                    pass
                
                # Check partial matches
                for value in accessibility_values:
                    if target_name.lower() in value.lower() or value.lower() in target_name.lower():
                        return ctrl
                
                # Search children recursively
                children = ctrl.GetChildren()
                for child in children:
                    result = search_recursive(child, target_name, current_depth + 1)
                    if result:
                        return result
                        
            except:
                pass
            
            return None
        
        return search_recursive(window, field_name)
    
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
        """Use LLM to analyze all UI elements with accessibility support"""
        try:
            if not self.openai_client or not UI_AUTOMATION_AVAILABLE:
                return None
            
            print("📊 Collecting all UI elements (including accessibility) for LLM analysis...")
            
            # Get all UI elements with accessibility properties
            all_elements = self._collect_all_ui_elements(window_name)
            if not all_elements:
                print("❌ No UI elements found for LLM analysis")
                return None
            
            # Create enhanced prompt for LLM
            prompt = self._create_llm_prompt(field_name, field_type, all_elements)
            
            # Get LLM response
            llm_response = self._query_llm(prompt)
            if not llm_response:
                return None
            
            # Find element based on LLM response
            return self._find_element_from_llm_response(llm_response, window_name)
            
        except Exception as e:
            print(f"❌ LLM fallback failed: {str(e)}")
            return None
    
    def _collect_all_ui_elements(self, window_name: str) -> List[Dict[str, Any]]:
        """Collect all UI elements with comprehensive accessibility properties"""
        try:
            window = self._find_window(window_name)
            if not window:
                return []
            
            elements = []
            
            def collect_element_info(ctrl, depth):
                try:
                    if depth > 10:  # Limit recursion depth
                        return
                    
                    # Basic element info
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
                    
                    # Comprehensive accessibility properties - THE KEY ENHANCEMENT!
                    element_info['accessible_name'] = getattr(ctrl, 'AccessibleName', '')
                    
                    # Advanced accessibility properties via GetPropertyValue
                    try:
                        name_prop = ctrl.GetPropertyValue(auto.PropertyId.NameProperty)
                        element_info['name_property'] = str(name_prop) if name_prop else ''
                    except:
                        element_info['name_property'] = ''
                    
                    try:
                        legacy_name = ctrl.GetPropertyValue(auto.PropertyId.LegacyIAccessibleNameProperty)
                        element_info['legacy_accessible_name'] = str(legacy_name) if legacy_name else ''
                    except:
                        element_info['legacy_accessible_name'] = ''
                    
                    try:
                        legacy_value = ctrl.GetPropertyValue(auto.PropertyId.LegacyIAccessibleValueProperty)
                        element_info['legacy_accessible_value'] = str(legacy_value) if legacy_value else ''
                    except:
                        element_info['legacy_accessible_value'] = ''
                    
                    try:
                        legacy_desc = ctrl.GetPropertyValue(auto.PropertyId.LegacyIAccessibleDescriptionProperty)
                        element_info['legacy_accessible_description'] = str(legacy_desc) if legacy_desc else ''
                    except:
                        element_info['legacy_accessible_description'] = ''
                    
                    # Bounds information
                    try:
                        if hasattr(ctrl, 'BoundingRectangle') and ctrl.BoundingRectangle:
                            rect = ctrl.BoundingRectangle
                            element_info['bounds'] = {
                                'x': rect.left, 'y': rect.top,
                                'width': rect.width(), 'height': rect.height()
                            }
                        else:
                            element_info['bounds'] = {'x': 0, 'y': 0, 'width': 0, 'height': 0}
                    except:
                        element_info['bounds'] = {'x': 0, 'y': 0, 'width': 0, 'height': 0}
                    
                    # Include element if it has ANY identifying information (including accessibility)
                    if (element_info['name'] or element_info['automation_id'] or 
                        element_info['class_name'] or element_info['accessible_name'] or
                        element_info['name_property'] or element_info['legacy_accessible_name'] or
                        element_info['legacy_accessible_value']):
                        elements.append(element_info)
                    
                    # Recursively collect children
                    try:
                        children = ctrl.GetChildren()
                        for child in children:
                            collect_element_info(child, depth + 1)
                    except:
                        pass
                        
                except:
                    pass  # Skip problematic elements
            
            collect_element_info(window, 0)
            print(f"📊 Collected {len(elements)} UI elements with accessibility properties")
            return elements
            
        except Exception as e:
            print(f"❌ Failed to collect UI elements: {str(e)}")
            return []
    
    def _create_llm_prompt(self, field_name: str, field_type: str, elements: List[Dict[str, Any]]) -> str:
        """Create enhanced LLM prompt emphasizing accessibility properties"""
        elements_json = json.dumps(elements, indent=2)
        
        prompt = f"""I need to find a UI element in a Windows desktop application. Here are the details:

TARGET ELEMENT:
- Field Name: "{field_name}"
- Field Type: "{field_type}"

AVAILABLE UI ELEMENTS (with full accessibility properties):
{elements_json}

TASK:
Find the element that best matches "{field_name}". The field name might be stored in ANY of these properties:

🔍 SEARCH PRIORITY (check ALL of these):
1. automation_id (AutomationId property)
2. name (Name property)
3. accessible_name (AccessibleName property) ⭐ VERY IMPORTANT
4. name_property (NameProperty via GetPropertyValue) 
5. legacy_accessible_name (Legacy IAccessible Name) ⭐ IMPORTANT
6. legacy_accessible_value (Legacy IAccessible Value)
7. class_name (ClassName property)

MATCHING STRATEGIES:
- Exact match on ANY accessibility property
- Partial match (case-insensitive) on accessibility properties
- Control type compatibility
- Contextual similarity

RESPONSE FORMAT (JSON only):
{{
    "matched_element": {{
        "name": "exact_name_from_list",
        "automation_id": "exact_automation_id_from_list",
        "accessible_name": "exact_accessible_name_from_list",
        "legacy_accessible_name": "exact_legacy_accessible_name_from_list",
        "class_name": "exact_class_name_from_list",
        "control_type": "exact_control_type_from_list"
    }},
    "confidence": 0.95,
    "reasoning": "Brief explanation of match",
    "matched_property": "name_of_property_that_matched"
}}

If no match: {{"matched_element": null, "confidence": 0.0, "reasoning": "No match found", "matched_property": null}}"""
        
        return prompt
    
    def _query_llm(self, prompt: str) -> Optional[Dict[str, Any]]:
        """Query LLM and parse response"""
        try:
            response = self.openai_client.chat.completions.create(
                model=self.openai_model,
                messages=[
                    {"role": "system", "content": "You are an expert at Windows UI element analysis. Return only valid JSON."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.1,
                max_tokens=1000
            )
            
            response_text = response.choices[0].message.content.strip()
            
            try:
                return json.loads(response_text)
            except json.JSONDecodeError:
                print(f"❌ LLM returned invalid JSON: {response_text}")
                return None
                
        except Exception as e:
            print(f"❌ LLM query failed: {str(e)}")
            return None
    
    def _find_element_from_llm_response(self, llm_response: Dict[str, Any], window_name: str) -> Optional[Any]:
        """Find element based on LLM response with accessibility support"""
        try:
            matched_element = llm_response.get('matched_element')
            confidence = llm_response.get('confidence', 0.0)
            reasoning = llm_response.get('reasoning', '')
            matched_property = llm_response.get('matched_property', '')
            
            print(f"🤖 LLM Match - Confidence: {confidence:.2f}, Property: {matched_property}")
            print(f"   Reasoning: {reasoning}")
            
            if not matched_element or confidence < 0.5:
                print("❌ LLM confidence too low or no match found")
                return None
            
            window = self._find_window(window_name)
            if not window:
                return None
            
            # Extract all possible identifiers from LLM response
            search_criteria = {
                'name': matched_element.get('name', ''),
                'automation_id': matched_element.get('automation_id', ''),
                'accessible_name': matched_element.get('accessible_name', ''),
                'legacy_accessible_name': matched_element.get('legacy_accessible_name', ''),
                'class_name': matched_element.get('class_name', '')
            }
            
            # Search using the comprehensive method
            def search_with_llm_criteria(ctrl, criteria, max_depth=5, current_depth=0):
                if current_depth >= max_depth:
                    return None
                
                try:
                    # Check all criteria from LLM
                    if criteria['automation_id'] and hasattr(ctrl, 'AutomationId') and ctrl.AutomationId == criteria['automation_id']:
                        return ctrl
                    if criteria['name'] and hasattr(ctrl, 'Name') and ctrl.Name == criteria['name']:
                        return ctrl
                    if criteria['accessible_name'] and hasattr(ctrl, 'AccessibleName') and ctrl.AccessibleName == criteria['accessible_name']:
                        return ctrl
                    if criteria['class_name'] and hasattr(ctrl, 'ClassName') and ctrl.ClassName == criteria['class_name']:
                        return ctrl
                    
                    # Check legacy accessible name
                    if criteria['legacy_accessible_name']:
                        try:
                            legacy_name = ctrl.GetPropertyValue(auto.PropertyId.LegacyIAccessibleNameProperty)
                            if legacy_name and str(legacy_name) == criteria['legacy_accessible_name']:
                                return ctrl
                        except:
                            pass
                    
                    # Search children
                    children = ctrl.GetChildren()
                    for child in children:
                        result = search_with_llm_criteria(child, criteria, max_depth, current_depth + 1)
                        if result:
                            return result
                            
                except:
                    pass
                
                return None
            
            element = search_with_llm_criteria(window, search_criteria)
            if element:
                identifier = (search_criteria['accessible_name'] or search_criteria['name'] or 
                            search_criteria['automation_id'] or search_criteria['class_name'])
                print(f"✅ Found element via LLM guidance: {identifier}")
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
            'enter': win32con.VK_RETURN, 'tab': win32con.VK_TAB, 'escape': win32con.VK_ESCAPE,
            'space': win32con.VK_SPACE, 'backspace': win32con.VK_BACK, 'delete': win32con.VK_DELETE,
            'home': win32con.VK_HOME, 'end': win32con.VK_END, 'pageup': win32con.VK_PRIOR,
            'pagedown': win32con.VK_NEXT, 'up': win32con.VK_UP, 'down': win32con.VK_DOWN,
            'left': win32con.VK_LEFT, 'right': win32con.VK_RIGHT,
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
                rect = win32gui.GetWindowRect(hwnd)
                x = (rect[0] + rect[2]) // 2
                y = (rect[1] + rect[3]) // 2
                
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
            self._click_element(element)
            time.sleep(0.2)
            
            if element['type'] == 'ui_automation':
                ui_element = element['element']
                ui_element.SendKeys('{Ctrl}a')
                time.sleep(0.5)
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
                time.sleep(0.5)
                
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
            self._click_element(element)
            time.sleep(0.1)
            
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
                
                # Try to find and select the item
                try:
                    children = ui_element.GetChildren()
                    for child in children:
                        if hasattr(child, 'Name') and child.Name.lower() == value.lower():
                            child.Select()
                            print(f"📋 Selected '{value}' from dropdown via UI Automation")
                            return True
                except:
                    pass
                
                # Fallback: click and type
                ui_element.Click()
                time.sleep(0.2)
                ui_element.SendKeys(value)
                time.sleep(0.2)
                ui_element.SendKeys('{Enter}')
                print(f"📋 Selected '{value}' by typing via UI Automation")
                return True
                    
            elif element['type'] == 'win32':
                hwnd = element['hwnd']
                self._click_element(element)
                time.sleep(0.2)
                
                # Type the value
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
            'ctrl': '{Ctrl}', 'alt': '{Alt}', 'shift': '{Shift}',
            'enter': '{Enter}', 'tab': '{Tab}', 'escape': '{Esc}',
            'space': ' ', 'backspace': '{Backspace}', 'delete': '{Delete}',
        }
        
        # Function keys
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
    Main LangGraph node for desktop automation with full AccessibleName support
    
    Integrates WindowManager, ElementDetector, and ActionExecutor to provide
    complete desktop automation capabilities with enhanced accessibility support.
    """
    
    def __init__(self, openai_api_key: Optional[str] = None, openai_model: str = "gpt-4"):
        """
        Initialize the desktop automation node
        
        Args:
            openai_api_key: Optional OpenAI API key for LLM fallback
            openai_model: OpenAI model to use (default: gpt-4)
        """
        print("🚀 Initializing Desktop Automation Node with AccessibleName support...")
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
        
        # For other actions, find the element first (with AccessibleName support)
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
    Demo script to test the desktop automation node with AccessibleName support
    """
    print("🖥️  Desktop Automation Node - Complete Clean Implementation")
    print("🎯 Now with full AccessibleName support!")
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
    
    # Notepad test (for quick validation)
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
            "Data": "Hello from Enhanced Desktop Automation with AccessibleName support!"
        }
    ]
    
    print("Choose a test option:")
    print("1. Test with your Settlement Request application (AccessibleName support)")
    print("2. Test with Notepad (quick validation)")
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
        
        print("\n🎯 Testing with AccessibleName support - your field names should now be found!")
        
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