# VBA Source Code - Import Guide

## Overview

This folder contains all VBA source code for the Rakuten MS2RSS Stock Data Collector.

**✅ COMPILE ERROR FIXED**: All files have been converted to English and Attribute lines removed to resolve compilation issues.

**IMPORTANT**: 
- All `Attribute VB_Name` lines have been removed (VBA Editor generates them automatically)
- UserForms and Class modules removed to avoid complex syntax errors
- Simple InputBox interface used instead of complex GUI

## File Structure

### 📁 modules/ - VBA Modules
| File | Purpose | Key Functions |
|------|---------|---------------|
| **MainModule.bas** | Main entry point | `ShowMainForm()`, `QuickTest()` |
| **WorksheetMacros.bas** | Worksheet button macros | `StartDataCollection()` etc. |
| **DataCollector.bas** | Data collection engine | `CollectStockData()` |
| **CSVExporter.bas** | CSV export functionality | `ExportStockDataToCSV()` |
| **Utils.bas** | Utilities & logging | `LogMessage()`, `EnsureDirectoryExists()` |
| **SimpleTest.bas** | Test functions | `TestBasic()`, `TestStockCollection()` |

### 📁 forms/ - User Forms
| File | Purpose |
|------|---------|
| ~~**MainForm.frm**~~ | Removed due to compatibility issues - using InputBox interface instead |

### 📁 classes/ - Class Modules
**Note**: Class modules have been removed to avoid Attribute compilation errors. All functionality is implemented in standard modules.

## Import Instructions for Excel

### 1. Create New Excel File
1. Open Microsoft Excel
2. Create new workbook
3. Save as `StockDataCollector.xlsm` (macro-enabled workbook)

### 2. Open VBA Editor
1. Press `Alt + F11` to open VBA Editor
2. Confirm VBAProject in Project Explorer

### 3. Add References (Optional)
1. In VBA Editor, select "Tools" → "References"
2. Check the following items if needed:
   - ✅ Microsoft Office Object Library
   - ✅ Microsoft Forms 2.0 Object Library

### 4. Import Modules

#### Standard Modules (.bas)
1. Right-click in Project Explorer
2. Select "Import File"
3. Import files in this order:
   ```
   1. modules/Utils.bas          (utilities first)
   2. modules/CSVExporter.bas    (export functions)
   3. modules/DataCollector.bas  (data collection)
   4. modules/MainModule.bas     (main functions)
   5. modules/WorksheetMacros.bas (button macros)
   ```

#### User Forms (.frm)
**Note**: User forms have been removed to avoid compatibility issues. The system now uses simple InputBox interface instead.

#### Class Modules (.cls)
**Note**: Class modules have been removed to avoid Attribute VB_Name compilation errors.

### 5. Worksheet Setup

#### Sheet1 Configuration
Set up Sheet1 with the following layout:

```
A1: Rakuten MS2RSS Stock Data Collector v1.0
A3: [Start Data Collection] (Button)
A5: [Quick Test] (Button)
A7: [Connection Test] (Button)
A9: [Help] (Button)

C3: [Open Output Folder] (Button)
C5: [Open Log Folder] (Button)
C7: [About] (Button)
```

#### Button Macro Assignments
Assign the following macros to each button:

| Button | Macro |
|--------|-------|
| Start Data Collection | `StartDataCollection` |
| Quick Test | `RunQuickTest` |
| Connection Test | `TestConnection` |
| Help | `ShowHelp` |
| Open Output Folder | `OpenOutputFolder` |
| Open Log Folder | `OpenLogFolder` |
| About | `AboutApp` |

## Basic Usage

### 1. Basic Test
```vba
' Test basic VBA functionality
Sub Test_Basic()
    Call TestBasic
End Sub
```

### 2. Launch Application
```vba
' Show main interface (InputBox based)
Sub Test_ShowMainForm()
    Call ShowMainForm
End Sub
```

### 3. Run Quick Test
```vba
' Test connection and data collection
Sub Test_QuickTest()
    Call QuickTest
End Sub
```

### 4. Direct Program Execution
```vba
Sub Test_DirectCall()
    Dim result As Boolean
    
    ' Collect Toyota 5-minute data for 1 week
    result = CollectStockData("7203", "5M", Date-7, Date)
    
    If result Then
        MsgBox "Data collection successful"
    Else
        MsgBox "Data collection failed"
    End If
End Sub
```

## Key Function Reference

### ShowMainForm()
Display main GUI form and start data collection

### CollectStockData(stockCode, timeFrame, startDate, endDate)
- **stockCode**: Stock code ("7203", "7203.T" etc.)
- **timeFrame**: Time frame ("1M", "5M", "15M", "30M", "60M", "D")
- **startDate**: Start date
- **endDate**: End date
- **Return**: Boolean (True if successful)

### CollectMultipleStocks(stockCodes, timeFrame, startDate, endDate)
Batch data collection for multiple stocks
- **stockCodes**: Comma-separated stock codes ("7203,6758,9984")

## Troubleshooting

### Common Errors

1. **"Procedure declaration does not match"**
   - ✅ FIXED: All files converted to English
   - Ensure all modules are imported correctly
   - Check reference settings

2. **"RSS function returns error"**
   - Ensure MarketSpeed2 is running
   - Verify RSS function is enabled

3. **"Cannot save file"**
   - Check if output folder exists
   - Verify folder write permissions

### Debug Methods

1. **Step execution**: F8 key for line-by-line execution
2. **Breakpoints**: F9 key to set breakpoints
3. **Immediate Window**: Ctrl+G to display
4. **Log check**: Check log files in `output/logs/` folder

## Important Notes

- Enable macro execution in Excel security settings
- MarketSpeed2 RSS function must be enabled
- Large data collection may take considerable time
- Perform thorough testing before production use

## Customization

### Adding New Time Frames
Modify `ValidateTimeFrame` function in `Utils.bas`

### Adding New Markets
Modify `ValidateStockCode` function in `DataCollector.bas`

### Changing UI Elements
Modify design in `MainForm.frm`

For detailed customization methods, refer to `docs/vba-guide.md`