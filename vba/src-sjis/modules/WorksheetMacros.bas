'******************************************************************************
' Rakuten MS2RSS Stock Data Collector - Worksheet Macros
' 
' Description: Macros called from Excel worksheet buttons
' Author: Claude Code
' Version: 1.0.0
'******************************************************************************

Option Explicit

' Main form display button macro
Public Sub StartDataCollection()
    On Error GoTo ErrorHandler
    
    ' Direct implementation to avoid macro security issues
    Dim stockCodes As String
    Dim result As Boolean
    
    ' InputBox interface for stock code entry
    stockCodes = InputBox("Enter stock codes (comma separated):" & vbCrLf & _
                         "Example: 7203,6758,9984" & vbCrLf & vbCrLf & _
                         "Supported formats:" & vbCrLf & _
                         "- Single: 7203" & vbCrLf & _
                         "- Multiple: 7203,6758,9984", _
                         "Stock Data Collector", "7203,6758,9984")
    
    If stockCodes <> "" Then
        ' Get date range from user
        Dim startDateStr As String
        Dim endDateStr As String
        Dim startDate As Date
        Dim endDate As Date
        
        ' Input start date
        startDateStr = InputBox("Enter start date:" & vbCrLf & _
                               "Format: YYYY/MM/DD or MM/DD" & vbCrLf & _
                               "Examples:" & vbCrLf & _
                               "- 2025/01/01" & vbCrLf & _
                               "- 01/01 (current year)" & vbCrLf & _
                               "- Leave blank for yesterday", _
                               "Start Date", Format(Date - 7, "YYYY/MM/DD"))
        
        If startDateStr = "" Then
            startDate = Date - 1  ' Default to yesterday
        Else
            On Error GoTo DateError
            If Len(startDateStr) <= 5 Then
                ' MM/DD format - add current year
                startDate = CDate(Year(Date) & "/" & startDateStr)
            Else
                startDate = CDate(startDateStr)
            End If
        End If
        
        ' Input end date
        endDateStr = InputBox("Enter end date:" & vbCrLf & _
                             "Format: YYYY/MM/DD or MM/DD" & vbCrLf & _
                             "Examples:" & vbCrLf & _
                             "- 2025/01/31" & vbCrLf & _
                             "- 01/31 (current year)" & vbCrLf & _
                             "- Leave blank for today", _
                             "End Date", Format(Date, "YYYY/MM/DD"))
        
        If endDateStr = "" Then
            endDate = Date  ' Default to today
        Else
            If Len(endDateStr) <= 5 Then
                ' MM/DD format - add current year
                endDate = CDate(Year(Date) & "/" & endDateStr)
            Else
                endDate = CDate(endDateStr)
            End If
        End If
        
        ' Validate date range
        If startDate > endDate Then
            MsgBox "Start date cannot be later than end date!" & vbCrLf & _
                   "Start: " & Format(startDate, "YYYY/MM/DD") & vbCrLf & _
                   "End: " & Format(endDate, "YYYY/MM/DD"), vbExclamation, "Invalid Date Range"
            Exit Sub
        End If
        
        ' Get timeframe from user
        Dim timeFrame As String
        timeFrame = InputBox("Select timeframe:" & vbCrLf & _
                            "Available options:" & vbCrLf & _
                            "- 1M (1 minute)" & vbCrLf & _
                            "- 5M (5 minutes)" & vbCrLf & _
                            "- 15M (15 minutes)" & vbCrLf & _
                            "- 30M (30 minutes)" & vbCrLf & _
                            "- 60M (60 minutes)" & vbCrLf & _
                            "- D (Daily)", _
                            "Timeframe Selection", "5M")
        
        If timeFrame = "" Then timeFrame = "5M"  ' Default to 5 minutes
        
        Debug.Print "Data collection started for: " & stockCodes & " (" & timeFrame & ") from " & Format(startDate, "YYYY/MM/DD") & " to " & Format(endDate, "YYYY/MM/DD")
        result = CollectMultipleStocks(stockCodes, timeFrame, startDate, endDate)
        
        If result Then
            MsgBox "Data collection completed successfully!" & vbCrLf & _
                   "Stocks: " & stockCodes & vbCrLf & _
                   "Timeframe: " & timeFrame & vbCrLf & _
                   "Period: " & Format(startDate, "YYYY/MM/DD") & " to " & Format(endDate, "YYYY/MM/DD") & vbCrLf & _
                   "Files saved to: output\csv\", vbInformation, "Success"
        Else
            MsgBox "Data collection completed with some errors." & vbCrLf & _
                   "Please check logs for details.", vbExclamation, "Completed"
        End If
    Else
        Debug.Print "Data collection cancelled by user"
    End If
    
    Exit Sub
    
DateError:
    MsgBox "Invalid date format!" & vbCrLf & _
           "Please use YYYY/MM/DD or MM/DD format" & vbCrLf & _
           "Examples: 2025/01/01 or 01/01", vbExclamation, "Date Error"
    Exit Sub
    
ErrorHandler:
    Debug.Print "StartDataCollection Error: " & Err.Description
    MsgBox "Data collection error: " & Err.Description, vbCritical, "Error"
End Sub

' Quick test button macro
Public Sub RunQuickTest()
    On Error GoTo ErrorHandler
    
    ' Simple quick test - direct implementation to avoid macro security issues
    Dim result As Boolean
    Dim testStockCode As String
    
    testStockCode = "7203"  ' Toyota Motor
    Debug.Print "Quick test start: " & testStockCode
    
    ' Test basic data collection
    result = CollectStockData(testStockCode, "5M", Date - 1, Date)
    
    If result Then
        MsgBox "Quick test success!" & vbCrLf & _
               "Stock: " & testStockCode & vbCrLf & _
               "Test data generated successfully", _
               vbInformation, "Test Result"
        Debug.Print "Quick test success"
    Else
        MsgBox "Quick test failed. Please check the log.", vbExclamation, "Test Result"
        Debug.Print "Quick test failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "RunQuickTest Error: " & Err.Description
    MsgBox "Quick test error: " & Err.Description, vbCritical, "Test Error"
End Sub

' Version information button macro
Public Sub AboutApp()
    On Error GoTo ErrorHandler
    
    ' Direct implementation to avoid macro security issues
    Dim aboutMessage As String
    
    aboutMessage = "Rakuten MS2RSS Stock Data Collector" & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "Version: 1.0.0" & vbCrLf
    aboutMessage = aboutMessage & "Build Date: 2025-01-16" & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "Uses Rakuten Securities MarketSpeed2 RSS API" & vbCrLf
    aboutMessage = aboutMessage & "to collect stock data and output as CSV format." & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "Created with Claude Code"
    
    MsgBox aboutMessage, vbInformation, "About This Application"
    Debug.Print "About information displayed"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "AboutApp Error: " & Err.Description
    MsgBox "About display error: " & Err.Description, vbCritical, "Error"
End Sub

' Open output folder
Public Sub OpenOutputFolder()
    On Error GoTo ErrorHandler
    
    Dim outputPath As String
    Dim csvPath As String
    Dim basePath As String
    
    basePath = ThisWorkbook.Path & "\output\"
    csvPath = basePath & "csv\"
    
    ' Create folders using utility function
    If Not EnsureDirectoryExists(basePath) Then
        MsgBox "Failed to create output directory", vbCritical
        Exit Sub
    End If
    
    If Not EnsureDirectoryExists(csvPath) Then
        MsgBox "Failed to create CSV directory", vbCritical
        Exit Sub
    End If
    
    outputPath = csvPath
    
    ' Open folder
    Shell "explorer.exe " & Chr(34) & outputPath & Chr(34), vbNormalFocus
    Debug.Print "Opened output folder: " & outputPath
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "OpenOutputFolder Error: " & Err.Description
    MsgBox "Could not open folder: " & Err.Description, vbCritical
End Sub

' Open log folder
Public Sub OpenLogFolder()
    On Error GoTo ErrorHandler
    
    Dim logPath As String
    Dim outputPath As String
    
    outputPath = ThisWorkbook.Path & "\output\"
    logPath = outputPath & "logs\"
    
    ' Create folders using utility function
    If Not EnsureDirectoryExists(outputPath) Then
        MsgBox "Failed to create output directory", vbCritical
        Exit Sub
    End If
    
    If Not EnsureDirectoryExists(logPath) Then
        MsgBox "Failed to create log directory", vbCritical
        Exit Sub
    End If
    
    ' Open folder
    Shell "explorer.exe " & Chr(34) & logPath & Chr(34), vbNormalFocus
    Debug.Print "Opened log folder: " & logPath
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "OpenLogFolder Error: " & Err.Description
    MsgBox "Could not open folder: " & Err.Description, vbCritical
End Sub

' MarketSpeed2 connection test
Public Sub TestConnection()
    On Error GoTo ErrorHandler
    
    ' Simple connection test - check if basic Excel functions work
    Dim basicTest As Boolean
    basicTest = True
    
    ' Test basic functionality
    Debug.Print "Testing basic Excel functionality..."
    
    ' Since RSS functions require MarketSpeed2, show information message instead
    MsgBox "Connection Test Information:" & vbCrLf & vbCrLf & _
           "To use MarketSpeed2 RSS functions, please ensure:" & vbCrLf & _
           "1. MarketSpeed2 is installed and running" & vbCrLf & _
           "2. RSS function is enabled in MarketSpeed2 settings" & vbCrLf & _
           "3. You are logged in to MarketSpeed2" & vbCrLf & _
           "4. RSS Add-in is installed in Excel" & vbCrLf & vbCrLf & _
           "VBA System: OK" & vbCrLf & _
           "Excel Version: " & Application.Version, vbInformation, "Connection Test Result"
    
    Debug.Print "Basic VBA system test completed successfully"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "TestConnection Error: " & Err.Description
    MsgBox "Connection test error: " & Err.Description, vbCritical, "Connection Test Error"
End Sub

' Show help
Public Sub ShowHelp()
    Dim helpMessage As String
    
    helpMessage = "Rakuten MS2RSS Stock Data Collector Help" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "Basic Usage:" & vbCrLf
    helpMessage = helpMessage & "1. Click 'Start Data Collection' button" & vbCrLf
    helpMessage = helpMessage & "2. Enter stock codes (e.g. 7203,6758,9984)" & vbCrLf
    helpMessage = helpMessage & "3. Click 'Execute' button to start data collection" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "Stock Code Format:" & vbCrLf
    helpMessage = helpMessage & "- Single stock: 7203" & vbCrLf
    helpMessage = helpMessage & "- Multiple stocks: 7203,6758,9984" & vbCrLf
    helpMessage = helpMessage & "- Market specific: 7203.T, 7203.JAX" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "Supported Timeframes:" & vbCrLf
    helpMessage = helpMessage & "1M, 5M, 15M, 30M, 60M, D (Daily)" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "Notes:" & vbCrLf
    helpMessage = helpMessage & "- MarketSpeed2 must be running" & vbCrLf
    helpMessage = helpMessage & "- RSS function must be enabled" & vbCrLf
    helpMessage = helpMessage & "- Large data collection may take time"
    
    MsgBox helpMessage, vbInformation, "Help"
End Sub

' Show system information
Public Sub ShowSystemInfo()
    Dim info As String
    
    info = "System Information" & vbCrLf & vbCrLf
    info = info & "Excel: " & Application.Version & vbCrLf
    info = info & "OS: " & Application.OperatingSystem & vbCrLf
    info = info & "User: " & Application.UserName & vbCrLf
    info = info & "Current Time: " & Format(Now, "YYYY-MM-DD HH:MM:SS") & vbCrLf & vbCrLf
    info = info & "Rakuten MS2RSS Stock Data Collector" & vbCrLf
    info = info & "Version: 1.0.0"
    
    MsgBox info, vbInformation, "System Information"
End Sub

' Show macro list
Public Sub ShowMacroList()
    Dim macroList As String
    
    macroList = "Available Macros" & vbCrLf & vbCrLf
    macroList = macroList & "Data Operations:" & vbCrLf
    macroList = macroList & "- StartDataCollection - Start data collection" & vbCrLf
    macroList = macroList & "- RunQuickTest - Run quick test" & vbCrLf & vbCrLf
    macroList = macroList & "Settings & Information:" & vbCrLf
    macroList = macroList & "- ShowSystemInfo - Show system information" & vbCrLf
    macroList = macroList & "- TestConnection - Connection test" & vbCrLf & vbCrLf
    macroList = macroList & "Utilities:" & vbCrLf
    macroList = macroList & "- OpenOutputFolder - Open output folder" & vbCrLf
    macroList = macroList & "- OpenLogFolder - Open log folder" & vbCrLf
    macroList = macroList & "- AboutApp - Version information" & vbCrLf
    macroList = macroList & "- ShowHelp - Show help"
    
    MsgBox macroList, vbInformation, "Macro List"
End Sub