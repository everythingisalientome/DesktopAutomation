# Enhanced Desktop Automation Tool

A powerful LangGraph-based desktop automation tool with hierarchical element detection using Windows UI Automation, Win32 API, OpenCV template matching, and OCR fallback.

## Key Features

✅ **Hierarchical Element Detection**: UI Automation → Win32 API → Template Matching → OCR  
✅ **Window-Scoped Automation**: Target specific application windows  
✅ **Auto Template Capture**: Automatically captures and stores element templates  
✅ **Comprehensive Analytics**: Detailed logging of detection methods and performance  
✅ **Special Key Support**: Full keyboard support including Ctrl+, Alt+, F-keys  
✅ **Element Caching**: Per-session caching for improved performance  
✅ **Corporate-Friendly**: No external dependencies, works in restricted environments  

## Installation

```bash
pip install -r requirements.txt
```

## Project Structure

```
enhanced_desktop_automation/
├── enhanced_desktop_automation.py    # Main automation tool
├── enhanced_examples.py              # Example implementations  
├── requirements.txt                  # Dependencies
├── templates/                        # Auto-generated templates
│   ├── push_button_a1b2c3d4.png
│   └── editable_text_e5f6g7h8.png
├── logs/                            # Analytics and logging
│   └── automation_analytics.log
└── README.md                        # This file
```

## Detection Methods (Priority Order)

### 1. **Windows UI Automation** (Primary - 20s timeout)
- Uses Microsoft UI Automation framework
- Detects elements by name, type, and properties
- Works with Win32 and legacy VB applications
- Most reliable for standard Windows controls

### 2. **Win32 API Detection** (Secondary - 5s timeout)  
- Uses Windows API calls (EnumChildWindows, FindWindow)
- Text-based element matching
- Good fallback for older applications

### 3. **OpenCV Template Matching** (Tertiary)
- Matches visual templates captured during first runs
- Auto-captures templates when elements found via methods 1-2
- Reliable for repeated automation tasks

### 4. **OCR Detection** (Last Resort)
- EasyOCR text detection on screenshots
- CPU-only, no GPU requirements
- Works when all other methods fail

## Usage

### Basic Usage

```python
from enhanced_desktop_automation import EnhancedDesktopAutomation

# Your automation steps with new window_name field
steps = [
    {
        "window_name": "Calculator",
        "field_name": "Five",
        "field_type": "push button",
        "event_action_type": "click",
        "SpecialKeyWithData": "",
        "Data": ""
    }
    # ... more steps
]

# Run automation
automation = EnhancedDesktopAutomation()
automation.run_automation("calc.exe", steps)
```

### Run Examples

```bash
python enhanced_examples.py
```

Choose from:
1. Wells Fargo Edge automation
2. Calculator automation  
3. Notepad with special keys
4. File Explorer navigation
5. Custom configuration example
6. Analytics demonstration

## JSON Step Format

Each automation step uses this enhanced structure:

```python
{
    "window_name": "Target application window title",      # NEW: Window scoping
    "field_name": "Element text or identifier",
    "field_type": "client|push button|editable text|link",
    "event_action_type": "click|typeInto|keyPress|windowActivate",
    "SpecialKeyWithData": "Enter|Ctrl+a|F1|etc",         # NEW: Special keys
    "Data": "Text to type or key to press"
}
```

## Supported Actions

- **click**: Click on element
- **typeInto**: Type text into element  
- **keyPress**: Send keyboard input (including special keys)
- **windowActivate**: Bring window to foreground (auto-handled)

## Special Key Support

### Basic Keys
- Enter, Tab, Escape, Space, Backspace, Delete
- Home, End, PageUp, PageDown
- Arrow keys: Up, Down, Left, Right
- Function keys: F1-F12

### Key Combinations  
- **Ctrl+** combinations: `Ctrl+a`, `Ctrl+c`, `Ctrl+v`, etc.
- **Alt+** combinations: `Alt+Tab`, `Alt+F4`, etc.  
- **Shift+** combinations: `Shift+Tab`, `Shift+F10`, etc.

## Element Detection Strategy

### Field Type Mapping
- **push button** → UI Automation Button controls
- **editable text** → UI Automation Edit controls  
- **link** → UI Automation Hyperlink controls
- **client** → Window-level operations

### Text Matching
- Uses "contains" matching (case-insensitive)
- `field_name: "Save"` matches "Save File", "Save As", etc.

## Analytics & Logging

### Automatic Logging
All automation runs generate detailed analytics in `logs/automation_analytics.log`:

```
2025-07-27 10:30:15 - DETECTION_SUCCESS - Method: UI_Automation, Element: Five, Type: push_button, Duration: 0.85s
2025-07-27 10:30:16 - Action executed: click on Five  
2025-07-27 10:30:17 - DETECTION_FAILURE - Element: Unknown_Button, Type: push_button, Duration: 20.12s
```

### Metrics Tracked
- Detection method used (UI_Automation, Win32_API, Template_Matching, OCR)
- Success/failure rates
- Detection timing
- Element types and names
- Window information

## Template Management

### Auto-Capture
- Templates automatically captured when elements found via UI Automation or Win32 API
- Stored in `templates/` directory with hash-based naming
- Format: `{field_type}_{field_name_hash}.png`

### Template Storage
```
templates/
├── push_button_a1b2c3d4.png     # "Five" button template
├── push_button_e5f6g7h8.png     # "Plus" button template  
└── editable_text_i9j0k1l2.png   # "Address bar" template
```

## Configuration Options

### Custom Directories
```python
automation = EnhancedDesktopAutomation(
    templates_dir="my_templates",
    logs_dir="my_logs"
)
```

### Detection Timeouts
- UI Automation: 20 seconds (configurable via `detection_timeout`)
- Win32 API: 5 seconds  
- Template Matching: Instant
- OCR: ~2-3 seconds

## Error Handling

### Window Not Found
- Stops automation immediately
- No desktop-wide fallback search
- Clear error message with window name

### Element Not Found  
- Tries all 4 detection methods sequentially
- Stops if all methods fail
- Logs which methods were attempted

### Action Failures
- Single retry with error on second failure
- No cross-method retry logic
- Clear error reporting

## Best Practices

### 1. Window Names
- Use specific, unique parts of window titles
- Avoid changing parts like document names
- Example: Use "Microsoft Edge" not "Google - Microsoft Edge"

### 2. Field Names
- Use text that appears consistently  
- Avoid UI text that changes (like timestamps)
- Use accessibility names when possible

### 3. Performance
- First run will be slower (template capture)
- Subsequent runs benefit from template matching
- UI Automation cache improves repeat element access

### 4. Debugging
- Check `logs/automation_analytics.log` for detection details
- Enable debug logging for detailed troubleshooting:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

## Troubleshooting

### Common Issues

**UI Automation Slow**: 
- Normal on first access to application
- Caching improves subsequent performance

**Templates Not Working**:
- Check `templates/` directory exists
- Verify template images were captured
- UI changes may invalidate templates

**Window Not Found**:
- Verify exact window title text
- Check if application launched successfully  
- Use Task Manager to see actual window titles

**Special Keys Not Working**:
- Verify key name spelling (case-sensitive)
- Check if target application has focus
- Some applications may block certain key combinations

### Debug Mode

Enable detailed logging:

```python
import logging
logging.basicConfig(level=logging.DEBUG, 
                   format='%(asctime)s - %(levelname)s - %(message)s')
```

## Dependencies

- **langgraph**: Workflow state management
- **pyautogui**: GUI automation and screenshots  
- **easyocr**: OCR text detection (CPU-only)
- **opencv-python**: Template matching and image processing
- **pywin32**: Windows UI Automation and API access
- **PIL/Pillow**: Image handling
- **numpy**: Array operations for image processing

## Limitations

- **Windows Only**: Uses Windows-specific APIs
- **Desktop Applications**: Not for web browsers (unless automated as desktop apps)
- **Standard Controls**: Works best with standard Windows UI elements
- **Visual Changes**: Template matching sensitive to UI theme/scaling changes

## Corporate Environment Notes

- **No GPU Required**: All image processing is CPU-only
- **No External APIs**: All processing done locally
- **Pip Installable**: All dependencies available via pip
- **File-Based**: Templates and logs stored locally
- **No Network**: No external network calls required

This enhanced tool provides enterprise-grade desktop automation with intelligent element detection and comprehensive analytics, perfect for scaling from pilot to production use.