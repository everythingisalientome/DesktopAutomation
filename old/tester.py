from enhanced_desktop_automation import EnhancedDesktopAutomation

def TestAutomation():
    steps = [
        {
            "window_name": "Settlement Request Form",
            "field_name": "ECN Number",
            "field_type": "editable text",
            "event_action_type": "typeInto",
            "SpecialKeyWithData": "",
            "Data": "12345678"
        },
        {
            "window_name": "Settlement Request Form",
            "field_name": "Account Number",
            "field_type": "editable text",
            "event_action_type": "typeInto",
            "SpecialKeyWithData": "",
            "Data": "12345678"
        },        
        {
            "window_name": "Settlement Request Form",
            "field_name": "Submit",
            "field_type": "push button",
            "event_action_type": "click",
            "SpecialKeyWithData": "",
            "Data": ""
        }
    ]
    
    automation_tool = EnhancedDesktopAutomation()
    exe_path = r"D:\VisualStudioCodeWrkSpce\Demo_App\dist\SettlementRequest.exe"
    
    try:
        print("Starting Wells Fargo Edge automation...")
        automation_tool.run_automation(exe_path, steps)
        print("Automation completed successfully!")
        
    except Exception as e:
        print(f"Automation failed: {str(e)}")


if __name__ == "__main__":
    TestAutomation()
    # Add any additional test cases or functionality here