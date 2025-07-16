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
    Call QuickTest
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
    
    Dim testResult As Variant
    
    ' Connection test with Nikkei 225 current value
    testResult = Application.WorksheetFunction.RssIndexMarket("0000", "Current Value")
    
    If IsError(testResult) Then
        MsgBox "Failed to connect to MarketSpeed2." & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "1. MarketSpeed2 is running" & vbCrLf & _
               "2. RSS function is enabled" & vbCrLf & _
               "3. Login status is normal", vbExclamation, "Connection Test Result"
        Debug.Print "MS2 connection test failed"
    Else
        MsgBox "Successfully connected to MarketSpeed2!" & vbCrLf & vbCrLf & _
               "Nikkei 225 Current Value: " & testResult, vbInformation, "Connection Test Result"
        Debug.Print "MS2 connection test success: " & testResult
    End If
    
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