Attribute VB_Name = "MainModule"
'******************************************************************************
' Rakuten MS2RSS Stock Data Collector - Main Module
' 
' Description: Application entry point and main control
' Author: Claude Code
' Version: 1.0.0
'******************************************************************************

Option Explicit

' Application information
Public Const APP_NAME As String = "Rakuten MS2RSS Stock Data Collector"
Public Const APP_VERSION As String = "1.0.0"
Public Const BUILD_DATE As String = "2025-01-16"

' Show main form
Public Sub ShowMainForm()
    On Error GoTo ErrorHandler
    
    ' Log initialization
    Debug.Print "Application start: " & APP_NAME & " v" & APP_VERSION
    
    ' Initial setup check
    If Not CheckInitialSetup() Then
        MsgBox "Initial setup problem. Please check the log.", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' Show main form
    Load MainForm
    MainForm.Show vbModal
    
    ' Cleanup after form is closed
    Unload MainForm
    Set MainForm = Nothing
    
    Debug.Print "Application end"
    Exit Sub
    
ErrorHandler:
    Debug.Print "ShowMainForm Error: " & Err.Description
    MsgBox "Application startup error: " & Err.Description, vbCritical, APP_NAME
End Sub

' Initial setup check
Private Function CheckInitialSetup() As Boolean
    On Error GoTo ErrorHandler
    
    Dim setupOK As Boolean
    setupOK = True
    
    ' Create output directories if they don't exist
    If Not EnsureDirectoryExists(ThisWorkbook.Path & "\output\csv\") Then
        Debug.Print "Failed to create CSV output directory"
        setupOK = False
    End If
    
    If Not EnsureDirectoryExists(ThisWorkbook.Path & "\output\logs\") Then
        Debug.Print "Failed to create log directory"
        setupOK = False
    End If
    
    CheckInitialSetup = setupOK
    Exit Function
    
ErrorHandler:
    Debug.Print "CheckInitialSetup Error: " & Err.Description
    CheckInitialSetup = False
End Function

' Quick test execution
Public Sub QuickTest()
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim testStockCode As String
    Dim testTimeFrame As String
    Dim testStartDate As Date
    Dim testEndDate As Date
    
    ' Test parameters
    testStockCode = "7203"  ' Toyota Motor
    testTimeFrame = "5M"    ' 5-minute bars
    testStartDate = Date - 1  ' Yesterday
    testEndDate = Date        ' Today
    
    Debug.Print "Quick test start: " & testStockCode
    
    ' Data collection test
    result = CollectStockData(testStockCode, testTimeFrame, testStartDate, testEndDate)
    
    If result Then
        MsgBox "Quick test success!" & vbCrLf & _
               "Stock: " & testStockCode & vbCrLf & _
               "Timeframe: " & testTimeFrame & vbCrLf & _
               "Period: " & Format(testStartDate, "MM/DD") & " - " & Format(testEndDate, "MM/DD"), _
               vbInformation, "Test Result"
        Debug.Print "Quick test success"
    Else
        MsgBox "Quick test failed. Please check the log.", vbExclamation, "Test Result"
        Debug.Print "Quick test failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "QuickTest Error: " & Err.Description & " (Stock: " & testStockCode & ")"
    MsgBox "Test execution error: " & Err.Description, vbCritical, "Error"
End Sub

' Show version information
Public Sub ShowAbout()
    Dim aboutMessage As String
    
    aboutMessage = APP_NAME & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "Version: " & APP_VERSION & vbCrLf
    aboutMessage = aboutMessage & "Build Date: " & BUILD_DATE & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "Uses Rakuten Securities MarketSpeed2 RSS API" & vbCrLf
    aboutMessage = aboutMessage & "to collect stock data and output as CSV format." & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "Created with Claude Code"
    
    MsgBox aboutMessage, vbInformation, "About This Application"
End Sub

' Application cleanup
Public Sub CleanupApplication()
    On Error Resume Next
    
    ' Clear progress display
    Application.StatusBar = False
    
    ' Restore application settings
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Log message
    Debug.Print "Application cleanup complete"
End Sub

' Directory existence check and creation
Private Function EnsureDirectoryExists(dirPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Dir(dirPath, vbDirectory) = "" Then
        MkDir dirPath
        Debug.Print "Directory created: " & dirPath
    End If
    
    EnsureDirectoryExists = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Directory creation error: " & dirPath & " - " & Err.Description
    EnsureDirectoryExists = False
End Function