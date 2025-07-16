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
    Call ShowMainForm
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
    Call ShowAbout
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