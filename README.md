# Desktop Automation LangGraph Node - Single File Implementation

A powerful, self-contained LangGraph-compatible node for automating Windows desktop applications using intelligent element detection and AI-powered fallbacks.

## 🚀 Features

### **Three-Tier Element Detection**
1. **UI Automation** (Primary) - Fast, reliable element detection
2. **Win32 API** (Secondary) - Legacy application fallback  
3. **🤖 LLM Analysis** (Tertiary) - AI-powered element matching when others fail

### **Comprehensive Automation**
- ✅ **All Action Types**: click, typeInto, keyPress, select, windowActivate
- ✅ **Special Key Support**: Full support for Ctrl+combinations, F1-F12, Tab, etc.
- ✅ **Smart Waiting**: Waits up to 30 seconds for elements to appear
- ✅ **Error Handling**: One retry per step, stops entire flow on failure
- ✅ **Case Insensitive**: Actions work regardless of case
- ✅ **Skip Logic**: Automatically skips steps where field_name is "nan"

### **LangGraph Ready**
- ✅ **TypedDict State**: Proper state management with `AutomationState`
- ✅ **Single Node**: No internal orchestration - perfect for larger workflows
- ✅ **Thread Safe**: Supports multiple concurrent automations
- ✅ **No Dependencies**: Self-contained single file eliminates import issues

## 📦 Installation

### **Required Dependencies**
```bash
pip install pywin32 langgraph
```

### **Optional Dependencies**
```bash
pip install uiautomation  # Recommended for better element detection
pip install openai        # Required for LLM fallback
```

## 🏗️ Project Structure

```
your_project/
├── desktop_automation_node.py    # Complete single-file implementation
└── your_langgraph_project.py     # Your main LangGraph workflow
```

## 🎯 Quick Start

### **1. Basic Usage**

```python
from desktop_automation_node import DesktopAutomationNode, create_sample_state

# Your automation steps (JSON format)
automation_steps = [
    {
        "window_name": "My Application",
        "field_name": "txtUsername",
        "field_type": "editable text",
        "event_action_type": "typeInto",
        "SpecialKeyWithData": "",
        "Data": "john.doe"
    },
    {
        "window_name": "My Application",
        "field_name": "btnLogin",
        "field_type": "push button",
        "event_action_type": "click",
        "SpecialKeyWithData": "",
        "Data": ""
    }
]

# Create automation node and state
node = DesktopAutomationNode()
state = create_sample_state(
    exe_path=r'C:\Path\To\Your\App.exe',
    automation_steps=automation_steps
)

# Execute automation
result = node.execute_automation(state)
print(f"Success: {result['automation_success']}")
print(f"Result: {result['automation_result']}")
```

### **2. With LLM Fallback**

```python
import os

# Enable AI-powered element detection
node = DesktopAutomationNode(
    openai_api_key=os.getenv('OPENAI_API_KEY'),
    openai_model="gpt-4"  # or "gpt-3.5-turbo"
)
```

### **3. LangGraph Integration**

```python
from langgraph.graph import StateGraph
from desktop_automation_node import DesktopAutomationNode, AutomationState

# Create your LangGraph workflow
def create_automation_workflow():
    # Initialize the automation node
    automation_node = DesktopAutomationNode(
        openai_api_key=os.getenv('OPENAI_API_KEY')
    )
    
    # Create the graph
    graph = StateGraph(AutomationState)
    graph.add_node("desktop_automation", automation_node.execute_automation)
    
    # Add your other nodes...
    # graph.add_node("data_processing", your_other_node)
    # graph.add_edge("desktop_automation", "data_processing")
    
    return graph.compile()

# Use in your workflow
workflow = create_automation_workflow()
initial_state = {
    'exe_path': r'C:\Your\App.exe',
    'automation_steps': your_steps,
    'automation_result': '',
    'automation_success': False
}

final_state = workflow.invoke(initial_state)
```

## 📋 State Schema

```python
class AutomationState(TypedDict):
    exe_path: str                    # Path to executable
    automation_steps: List[Dict]     # JSON array of automation steps
    automation_result: str           # Result message
    automation_success: bool         # Success/failure flag
```

## 🎮 JSON Step Format

```json
{
  "window_name": "Application Window Name",
  "field_name": "elementId or elementName", 
  "field_type": "editable text|push button|list item|etc",
  "event_action_type": "click|typeInto|keyPress|select|windowActivate",
  "SpecialKeyWithData": "Tab|Enter|Ctrl+A|F5|etc", 
  "Data": "text to type or value to select"
}
```

## 🎯 Supported Actions

| Action Type | Description | Required Fields | Example |
|-------------|-------------|-----------------|---------|
| `windowActivate` | Activate/focus window | window_name or field_name | Bring app to foreground |
| `click` | Click on element | field_name | Click button or link |
| `typeInto` | Type text into element | field_name, Data | Fill text field |
| `keyPress` | Press key combination | field_name, SpecialKeyWithData | Press Tab, Ctrl+S, etc. |
| `select` | Select from dropdown/list | field_name, Data | Choose dropdown option |

## ⌨️ Special Key Combinations

The system supports all standard Windows key combinations:

### **Modifiers**
- `Ctrl`, `Alt`, `Shift` (can be combined)
- Examples: `"Ctrl+A"`, `"Ctrl+Shift+S"`, `"Alt+F4"`

### **Special Keys**
- `Tab`, `Enter`, `Escape`, `Space`, `Backspace`, `Delete`
- `Home`, `End`, `PageUp`, `PageDown`
- `Up`, `Down`, `Left`, `Right`

### **Function Keys**
- `F1`, `F2`, `F3`, ... `F12`

### **Examples**
```json
{
  "event_action_type": "keyPress",
  "SpecialKeyWithData": "Ctrl+A",
  "Data": ""
}
```

## 🤖 LLM Fallback (Advanced)

When both UI Automation and Win32 API fail to find an element, the system can use OpenAI to intelligently analyze all UI elements and find the best match.

### **How It Works**
1. **Collects** all UI elements with their properties
2. **Analyzes** element data using GPT-4/3.5-turbo
3. **Matches** target elements using AI reasoning
4. **Returns** confident matches with explanations

### **Setup**
```bash
# Set your OpenAI API key
export OPENAI_API_KEY="your-api-key-here"
# or
set OPENAI_API_KEY=your-api-key-here  # Windows
```

### **Cost Considerations**
- Only triggers when other methods fail
- Typically costs $0.01-0.03 per failed element
- Uses focused prompts to minimize token usage

## 🧪 Testing

### **Built-in Test**
Run the file directly to test:
```bash
python desktop_automation_node.py
```

**Test Options:**
1. **Your application** - Test with your Settlement Request app
2. **Notepad test** - Quick validation with Windows Notepad
3. **Exit**

### **Sample Test Data**
The file includes your original sample data:
```python
sample_steps = [
    {
        "window_name": "Settlement Request Test Application",
        "field_name": "Settlement Request Test Application",
        "field_type": "client",
        "event_action_type": "windowActivate",
        "SpecialKeyWithData": "",
        "Data": ""
    },
    # ... your other steps
]
```

## 🔧 Configuration

### **Environment Variables**
```bash
# Optional: OpenAI API key for LLM fallback
export OPENAI_API_KEY="your-openai-api-key"

# Optional: Specify different OpenAI model
export OPENAI_MODEL="gpt-4"  # or "gpt-3.5-turbo"
```

### **Initialization Options**
```python
# Basic usage (no LLM fallback)
node = DesktopAutomationNode()

# With LLM fallback
node = DesktopAutomationNode(
    openai_api_key="your-key",
    openai_model="gpt-4"
)

# Using environment variable
node = DesktopAutomationNode(
    openai_api_key=os.getenv('OPENAI_API_KEY')
)
```

## ⚡ Performance

### **Typical Timing**
- **UI Automation**: 50-200ms per element search
- **Win32 API**: 100-500ms per element search  
- **LLM Fallback**: 2-5 seconds (includes API call)
- **Total Timeout**: 30 seconds per element before giving up

### **Optimization Tips**
1. Use specific element identifiers when possible
2. Enable LLM fallback only when needed
3. Consider pre-launching applications to reduce startup time
4. Use shorter timeouts for faster failure detection

## 🚨 Error Handling

### **Automatic Retry Logic**
- **Element Search**: Waits up to 30 seconds for elements to appear
- **Action Execution**: One automatic retry per step
- **Application Launch**: Immediate failure with clear error message
- **Flow Control**: Stops entire automation on step failure

### **Skip Logic**
Steps with `field_name: "nan"` are automatically skipped:
```json
{
  "window_name": "nan",
  "field_name": "nan",
  "field_type": "title bar",
  "event_action_type": "Click",
  "SpecialKeyWithData": "",
  "Data": ""
}
```

## 🎯 Use Cases

### **Form Automation**
- Fill out complex desktop application forms
- Handle multiple windows and dialogs
- Support for all input types (text, dropdowns, buttons)

### **Data Entry**
- Bulk data entry into legacy applications
- CSV/Excel data processing workflows
- Automated report generation

### **Testing & QA**
- Automated UI testing of desktop applications
- Regression testing workflows
- User acceptance testing automation

### **Workflow Integration**
- Connect desktop apps with modern workflows
- Bridge legacy systems with new processes
- Automate repetitive business tasks

### **Legacy System Support**
- Handle applications with poor automation support
- Work with older applications using Win32 fallback
- AI-powered element detection for difficult UIs

## 🔍 Troubleshooting

### **Element Not Found**
1. **Check element names** - Use UI inspection tools (Inspect.exe, UISpy)
2. **Verify window titles** - Ensure exact or partial matches
3. **Enable LLM fallback** - For complex element identification
4. **Increase timeout** - For slow-loading applications

### **LLM Fallback Issues**
1. **API Key** - Verify OpenAI API key is valid and has credits
2. **Internet Connection** - Check connectivity for API calls
3. **Element Collection** - Review logs for UI element collection
4. **Confidence Threshold** - LLM needs >50% confidence to proceed

### **Performance Issues**
1. **Reduce timeout** - For faster failure detection
2. **Specific identifiers** - Use AutomationId when available
3. **Pre-launch apps** - Reduce application startup time
4. **Monitor resources** - Check system performance during automation

### **Import Errors**
✅ **Solved!** - Single file eliminates all import issues

## 📚 Advanced Examples

### **Complex Workflow**
```python
# Multi-step automation with error handling
automation_steps = [
    # Step 1: Activate application
    {
        "window_name": "My App",
        "field_name": "My App",
        "field_type": "window",
        "event_action_type": "windowActivate",
        "SpecialKeyWithData": "",
        "Data": ""
    },
    # Step 2: Navigate to form
    {
        "window_name": "My App",
        "field_name": "menuFile",
        "field_type": "menu",
        "event_action_type": "click",
        "SpecialKeyWithData": "",
        "Data": ""
    },
    # Step 3: Fill form data
    {
        "window_name": "My App",
        "field_name": "txtCustomerName",
        "field_type": "editable text",
        "event_action_type": "typeInto",
        "SpecialKeyWithData": "",
        "Data": "John Smith"
    },
    # Step 4: Submit form
    {
        "window_name": "My App",
        "field_name": "btnSubmit",
        "field_type": "push button",
        "event_action_type": "click",
        "SpecialKeyWithData": "",
        "Data": ""
    }
]
```

### **Dynamic Data Processing**
```python
# Process multiple records
def process_customer_data(customers):
    node = DesktopAutomationNode()
    
    for customer in customers:
        # Create dynamic steps for each customer
        steps = create_customer_steps(customer)
        state = create_sample_state(app_path, steps)
        
        result = node.execute_automation(state)
        if not result['automation_success']:
            print(f"Failed to process customer: {customer['name']}")
            break
        else:
            print(f"Successfully processed: {customer['name']}")
```

## 🛡️ Security & Limitations

### **Security Considerations**
- **Local Execution**: All automation runs locally on your machine
- **No Data Upload**: Element data only sent to OpenAI for LLM fallback
- **API Key Security**: Store OpenAI keys securely (environment variables)
- **Process Isolation**: Each automation runs in its own process space

### **Limitations**
- **Windows Only**: Uses Windows-specific APIs (pywin32, UI Automation)
- **Desktop Apps Only**: Not for web browsers or mainframes
- **No OCR**: Relies purely on UI element detection
- **Sequential Execution**: No parallel step execution
- **Internet Required**: LLM fallback needs internet connectivity

## 📞 Support & Contributing

### **Common Issues**
- Check Windows version compatibility (Windows 7+)
- Ensure pywin32 is properly installed
- Verify application accessibility settings
- Test with simple applications first (Notepad, Calculator)

### **Feature Requests**
The single-file architecture makes it easy to extend:
- Add new action types in `ActionExecutor`
- Enhance element detection in `ElementDetector`
- Improve error handling in `DesktopAutomationNode`

### **Best Practices**
1. **Start Simple** - Test with basic applications first
2. **Use Descriptive Names** - Clear field names improve success rates
3. **Handle Timing** - Allow adequate time for slow applications
4. **Monitor Logs** - Rich logging helps debug issues
5. **Gradual Complexity** - Build complex workflows incrementally

---

## 🎉 Ready to Use!

The desktop automation node is completely self-contained and ready to integrate into your LangGraph workflows. No more import issues, no complex file structures - just powerful desktop automation in a single file!

**Happy Automating! 🚀**