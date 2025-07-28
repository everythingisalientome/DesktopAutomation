# desktop_automation_node/examples/sample_data.json
[
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

# desktop_automation_node/examples/test_node.py
"""
Test script to run the desktop automation node standalone
"""
import json
import sys
import os

# Add the parent directory to Python path to import core modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from desktop_automation import DesktopAutomationNode, AutomationState


def load_sample_data():
    """Load sample automation steps from JSON file"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_file = os.path.join(script_dir, 'sample_data.json')
    
    with open(json_file, 'r') as f:
        return json.load(f)


def test_automation_node():
    """Test the automation node with sample data"""
    
    # Create initial state
    state: AutomationState = {
        'exe_path': r'C:\Path\To\Your\Application.exe',  # Replace with actual path
        'automation_steps': load_sample_data(),
        'automation_result': '',
        'automation_success': False
    }
    
    print("=" * 60)
    print("DESKTOP AUTOMATION NODE TEST")
    print("=" * 60)
    print(f"Executable Path: {state['exe_path']}")
    print(f"Number of Steps: {len(state['automation_steps'])}")
    print("=" * 60)
    
    # Create and run the automation node
    automation_node = DesktopAutomationNode()
    result_state = automation_node.execute_automation(state)
    
    # Display results
    print("\n" + "=" * 60)
    print("AUTOMATION RESULTS")
    print("=" * 60)
    print(f"Success: {result_state['automation_success']}")
    print(f"Result: {result_state['automation_result']}")
    print("=" * 60)
    
    return result_state


def test_with_custom_data():
    """Test with custom executable and steps"""
    
    # Example: Test with Notepad
    custom_state: AutomationState = {
        'exe_path': r'C:\Windows\System32\notepad.exe',
        'automation_steps': [
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
            },
            {
                "window_name": "Untitled - Notepad",
                "field_name": "Text Editor", 
                "field_type": "editable text",
                "event_action_type": "keyPress",
                "SpecialKeyWithData": "Ctrl+A",
                "Data": ""
            }
        ],
        'automation_result': '',
        'automation_success': False
    }
    
    print("\n" + "=" * 60)
    print("TESTING WITH NOTEPAD")
    print("=" * 60)
    
    automation_node = DesktopAutomationNode()
    result_state = automation_node.execute_automation(custom_state)
    
    print(f"Success: {result_state['automation_success']}")
    print(f"Result: {result_state['automation_result']}")
    
    return result_state


if __name__ == '__main__':
    print("Desktop Automation Node - Test Script")
    print("Choose test option:")
    print("1. Test with sample data (requires Settlement Request Test Application)")
    print("2. Test with Notepad (Windows built-in)")
    print("3. Exit")
    
    choice = input("\nEnter choice (1-3): ").strip()
    
    if choice == '1':
        # Update the exe_path in the function above before running
        print("\nIMPORTANT: Update the exe_path in test_automation_node() function")
        print("with the actual path to your Settlement Request Test Application")
        
        confirm = input("Have you updated the exe_path? (y/n): ").strip().lower()
        if confirm == 'y':
            test_automation_node()
        else:
            print("Please update the exe_path and run again.")
            
    elif choice == '2':
        test_with_custom_data()
        
    elif choice == '3':
        print("Exiting...")
    else:
        print("Invalid choice!")
