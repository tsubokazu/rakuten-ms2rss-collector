'******************************************************************************
' Rakuten MS2RSS Stock Data Collector - Utility Module
' 
' Description: Common utility functions and logging
' Author: Claude Code
' Version: 1.0.0
'******************************************************************************

Option Explicit

' Log level constants
Public Const LOG_DEBUG As String = "DEBUG"
Public Const LOG_INFO As String = "INFO"
Public Const LOG_WARN As String = "WARN"
Public Const LOG_ERROR As String = "ERROR"

' Log message output (simple version)
Public Sub LogMessage(level As String, message As String)
    On Error Resume Next
    
    Dim logLine As String
    Dim timestamp As String
    
    ' Generate timestamp
    timestamp = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    ' Create log line
    logLine = timestamp & " [" & level & "] " & message
    
    ' Console output (Immediate Window)
    Debug.Print logLine
End Sub

' Directory existence check and creation (simple version)
Public Function EnsureDirectoryExists(dirPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Dir(dirPath, vbDirectory) = "" Then
        MkDir dirPath
        Call LogMessage(LOG_INFO, "Directory created: " & dirPath)
    End If
    
    EnsureDirectoryExists = True
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "Directory creation error: " & dirPath & " - " & Err.Description)
    EnsureDirectoryExists = False
End Function

' Detailed error information logging (simple version)
Public Sub LogDetailedError(functionName As String, errorDescription As String, _
                          Optional additionalInfo As String = "")
    
    Dim errorMessage As String
    
    errorMessage = "Function: " & functionName & " / Error: " & errorDescription
    
    If additionalInfo <> "" Then
        errorMessage = errorMessage & " / Additional info: " & additionalInfo
    End If
    
    Call LogMessage(LOG_ERROR, errorMessage)
End Sub