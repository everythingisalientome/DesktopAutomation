"""
Enhanced Desktop Automation Tool Examples
Demonstrates the new window-scoped automation with improved detection methods
"""

from enhanced_desktop_automation import EnhancedDesktopAutomation


def example_wells_fargo_edge_automation():
    """Example: Wells Fargo automation in Microsoft Edge"""
    
    steps = [
        {
            "window_name": "Wells Fargo Technology  [ORGANIZATION]? Edge",
            "field_name": "Wells Fargo Technology - Work - Microsoft? Edge",
            "field_type": "client",
            "event_action_type": "windowActivate",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Wells Fargo Technology  [ORGANIZATION]? Edge",
            "field_name": "Address and search bar",
            "field_type": "editable text",
            "event_action_type": "click",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Wells Fargo Technology  [ORGANIZATION]? Edge",
            "field_name": "Address and search bar",
            "field_type": "editable text",
            "event_action_type": "typeInto",
            "SpecialKeyWithData": "",
            "Data": "www.wellsfargo.com"
        },
        {
            "window_name": "Wells Fargo Technology  [ORGANIZATION]? Edge",
            "field_name": "Wells Fargo Technology - Work - Microsoft? Edge",
            "field_type": "client",
            "event_action_type": "windowActivate",
            "SpecialKeyWithData": "",
            "Data": "Enter"
        },
        {
            "window_name": "Wells Fargo Technology  [ORGANIZATION]? Edge",
            "field_name": "Address and search bar",
            "field_type": "editable text",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "Enter",
            "Data": "Enter"
        }
    ]
    
    automation_tool = EnhancedDesktopAutomation()
    exe_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
    
    try:
        print("Starting Wells Fargo Edge automation...")
        automation_tool.run_automation(exe_path, steps)
        print("Automation completed successfully!")
        
    except Exception as e:
        print(f"Automation failed: {str(e)}")


def example_calculator_automation():
    """Example: Windows Calculator automation with special keys"""
    
    steps = [
        {
            "window_name": "Calculator",
            "field_name": "Calculator",
            "field_type": "client",
            "event_action_type": "windowActivate",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Calculator",
            "field_name": "Five",
            "field_type": "push button",
            "event_action_type": "click",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Calculator",
            "field_name": "Plus",
            "field_type": "push button",
            "event_action_type": "click",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Calculator",
            "field_name": "Three",
            "field_type": "push button",
            "event_action_type": "click",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Calculator",
            "field_name": "Calculator",
            "field_type": "client",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "Enter",
            "Data": ""
        }
    ]
    
    automation_tool = EnhancedDesktopAutomation()
    exe_path = "calc.exe"
    
    try:
        print("Starting Calculator automation...")
        automation_tool.run_automation(exe_path, steps)
        print("Automation completed successfully!")
        
    except Exception as e:
        print(f"Automation failed: {str(e)}")


def example_notepad_with_special_keys():
    """Example: Notepad automation with various key combinations"""
    
    steps = [
        {
            "window_name": "Untitled - Notepad",
            "field_name": "Text Editor",
            "field_type": "editable text",
            "event_action_type": "click",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "Untitled - Notepad",
            "field_name": "Text Editor",
            "field_type": "editable text",
            "event_action_type": "typeInto",
            "SpecialKeyWithData": "",
            "Data": "Hello World! This is automated text."
        },
        {
            "window_name": "Untitled - Notepad",
            "field_name": "Text Editor",
            "field_type": "editable text",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "Ctrl+a",
            "Data": ""
        },
        {
            "window_name": "Untitled - Notepad",
            "field_name": "Text Editor",
            "field_type": "editable text",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "Ctrl+c",
            "Data": ""
        },
        {
            "window_name": "Untitled - Notepad",
            "field_name": "Text Editor",
            "field_type": "editable text",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "End",
            "Data": ""
        },
        {
            "window_name": "Untitled - Notepad",
            "field_name": "Text Editor",
            "field_type": "editable text",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "Enter",
            "Data": ""
        },
        {
            "window_name": "Untitled - Notepad",
            "field_name": "Text Editor",
            "field_type": "editable text",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "Ctrl+v",
            "Data": ""
        }
    ]
    
    automation_tool = EnhancedDesktopAutomation()
    exe_path = "C:\\Windows\\System32\\notepad.exe"
    
    try:
        print("Starting Notepad automation with special keys...")
        automation_tool.run_automation(exe_path, steps)
        print("Automation completed successfully!")
        
    except Exception as e:
        print(f"Automation failed: {str(e)}")


def example_file_explorer_automation():
    """Example: File Explorer navigation"""
    
    steps = [
        {
            "window_name": "File Explorer",
            "field_name": "File Explorer",
            "field_type": "client",
            "event_action_type": "windowActivate",
            "SpecialKeyWithData": "",
            "Data": ""
        },
        {
            "window_name": "File Explorer",
            "field_name": "Address bar",
            "field_type": "editable text",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "Ctrl+l",
            "Data": ""
        },
        {
            "window_name": "File Explorer",
            "field_name": "Address bar",
            "field_type": "editable text",
            "event_action_type": "typeInto",
            "SpecialKeyWithData": "",
            "Data": "C:\\Users\\Public\\Documents"
        },
        {
            "window_name": "File Explorer",
            "field_name": "Address bar",
            "field_type": "editable text",
            "event_action_type": "keyPress",
            "SpecialKeyWithData": "Enter",
            "Data": ""
        }
    ]
    
    automation_tool = EnhancedDesktopAutomation()
    exe_path = "explorer.exe"
    
    try:
        print("Starting File Explorer automation...")
        automation_tool.run_automation(exe_path, steps)
        print("Automation completed successfully!")
        
    except Exception as e:
        print(f"Automation failed: {str(e)}")


def run_automation_from_dict(automation_config):
    """Run automation from a configuration dictionary"""
    
    exe_path = automation_config.get('exe_path')
    steps = automation_config.get('steps')
    
    if not exe_path or not steps:
        print("Invalid configuration. Must contain 'exe_path' and 'steps'.")
        return
    
    try:
        automation_tool = EnhancedDesktopAutomation()
        print(f"Starting automation for: {exe_path}")
        automation_tool.run_automation(exe_path, steps)
        print("Automation completed successfully!")
        
    except Exception as e:
        print(f"Automation failed: {str(e)}")


def example_custom_configuration():
    """Example: Running automation from custom configuration"""
    
    # Custom configuration dictionary
    config = {
        "exe_path": "calc.exe",
        "steps": [
            {
                "window_name": "Calculator",
                "field_name": "Calculator",
                "field_type": "client",
                "event_action_type": "windowActivate",
                "SpecialKeyWithData": "",
                "Data": ""
            },
            {
                "window_name": "Calculator",
                "field_name": "Two",
                "field_type": "push button",
                "event_action_type": "click",
                "SpecialKeyWithData": "",
                "Data": ""
            },
            {
                "window_name": "Calculator",
                "field_name": "Multiply by",
                "field_type": "push button",
                "event_action_type": "click",
                "SpecialKeyWithData": "",
                "Data": ""
            },
            {
                "window_name": "Calculator",
                "field_name": "Six",
                "field_type": "push button",
                "event_action_type": "click",
                "SpecialKeyWithData": "",
                "Data": ""
            },
            {
                "window_name": "Calculator",
                "field_name": "Calculator",
                "field_type": "client",
                "event_action_type": "keyPress",
                "SpecialKeyWithData": "F9",
                "Data": ""
            }
        ]
    }
    
    print("Starting custom configuration automation...")
    run_automation_from_dict(config)


def example_analytics_demonstration():
    """Demonstrate analytics logging by running multiple examples"""
    
    print("Running multiple automations to demonstrate analytics...")
    print("Check the 'logs/automation_analytics.log' file for detailed metrics.")
    
    # Run Calculator example
    try:
        example_calculator_automation()
    except:
        pass
    
    # Small delay between examples
    import time
    time.sleep(2)
    
    # Run Notepad example
    try:
        example_notepad_with_special_keys()
    except:
        pass
    
    print("\nAnalytics logging demonstration completed!")
    print("Review logs/automation_analytics.log to see:")
    print("- Detection methods used")
    print("- Success/failure rates")
    print("- Timing information")
    print("- Element types and names")


if __name__ == "__main__":
    print("Enhanced Desktop Automation Tool - Examples")
    print("=" * 50)
    
    print("Available examples:")
    print("1. Wells Fargo Edge automation")
    print("2. Calculator automation")
    print("3. Notepad with special keys")
    print("4. File Explorer navigation")
    print("5. Custom configuration example")
    print("6. Analytics demonstration")
    
    choice = input("Enter your choice (1-6): ").strip()
    
    if choice == "1":
        example_wells_fargo_edge_automation()
    elif choice == "2":
        example_calculator_automation()
    elif choice == "3":
        example_notepad_with_special_keys()
    elif choice == "4":
        example_file_explorer_automation()
    elif choice == "5":
        example_custom_configuration()
    elif choice == "6":
        example_analytics_demonstration()
    else:
        print("Invalid choice. Please run again and select 1-6.")