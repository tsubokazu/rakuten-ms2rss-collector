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

' Stock code validation
Public Function ValidateStockCode(stockCode As String) As Boolean
    On Error GoTo ErrorHandler
    
    ValidateStockCode = False
    
    If stockCode = "" Then Exit Function
    
    ' Remove market suffix if present
    Dim codeOnly As String
    If InStr(stockCode, ".") > 0 Then
        codeOnly = Split(stockCode, ".")(0)
    Else
        codeOnly = stockCode
    End If
    
    ' Check if it's numeric and proper length
    If IsNumeric(codeOnly) And (Len(codeOnly) = 4 Or Len(codeOnly) = 5) Then
        ValidateStockCode = True
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("ValidateStockCode", Err.Description, "StockCode: " & stockCode)
    ValidateStockCode = False
End Function

' Calculate required data points for period-based data collection
Public Function CalculateRequiredDataPoints(startDate As Date, endDate As Date, timeFrame As String) As Long
    On Error GoTo ErrorHandler
    
    Dim daysDiff As Long
    Dim businessDays As Long
    Dim dataPoints As Long
    
    daysDiff = DateDiff("d", startDate, endDate) + 1
    businessDays = CalculateBusinessDays(startDate, endDate)
    
    Select Case UCase(timeFrame)
        Case "1M"
            dataPoints = businessDays * 390  ' 6.5 hours * 60 minutes
        Case "5M"
            dataPoints = businessDays * 78   ' 6.5 hours * 60 / 5
        Case "15M"
            dataPoints = businessDays * 26   ' 6.5 hours * 60 / 15
        Case "30M"
            dataPoints = businessDays * 13   ' 6.5 hours * 60 / 30
        Case "60M"
            dataPoints = businessDays * 7    ' 6.5 hours (rounded up)
        Case "D"
            dataPoints = businessDays
        Case "W"
            dataPoints = businessDays / 5
        Case "M"
            dataPoints = businessDays / 20  ' Approximately 20 business days per month
        Case Else
            dataPoints = businessDays * 78  ' Default to 5M
    End Select
    
    ' Add safety margin for minute data
    If InStr(timeFrame, "M") > 0 And timeFrame <> "M" Then
        dataPoints = dataPoints * 1.2  ' 20% safety margin
    End If
    
    CalculateRequiredDataPoints = dataPoints
    
    Call LogMessage(LOG_INFO, "Required data points calculated: " & dataPoints & " for " & timeFrame & " from " & Format(startDate, "YYYY-MM-DD") & " to " & Format(endDate, "YYYY-MM-DD"))
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CalculateRequiredDataPoints", Err.Description, "TimeFrame: " & timeFrame)
    CalculateRequiredDataPoints = 0
End Function

' Calculate business days between two dates
Public Function CalculateBusinessDays(startDate As Date, endDate As Date) As Long
    On Error GoTo ErrorHandler
    
    Dim currentDate As Date
    Dim businessDays As Long
    
    currentDate = startDate
    businessDays = 0
    
    Do While currentDate <= endDate
        ' Check if it's a weekday (Monday = 2, Friday = 6)
        If Weekday(currentDate) >= 2 And Weekday(currentDate) <= 6 Then
            ' Simple business day check (excludes weekends only)
            ' TODO: Add holiday calendar support
            businessDays = businessDays + 1
        End If
        currentDate = currentDate + 1
    Loop
    
    CalculateBusinessDays = businessDays
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CalculateBusinessDays", Err.Description)
    CalculateBusinessDays = 0
End Function

' Determine which RSS function to use based on timeframe
Public Function GetRSSFunctionType(timeFrame As String) As String
    On Error GoTo ErrorHandler
    
    Select Case UCase(timeFrame)
        Case "1M", "5M", "15M", "30M", "60M"
            GetRSSFunctionType = "RssChart"  ' For minute data
        Case "D", "W", "M"
            GetRSSFunctionType = "RssChartPast"  ' For daily and above
        Case Else
            GetRSSFunctionType = "RssChart"  ' Default
    End Select
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("GetRSSFunctionType", Err.Description, "TimeFrame: " & timeFrame)
    GetRSSFunctionType = "RssChart"
End Function

' Calculate number of batches needed for data collection
Public Function CalculateBatchCount(totalDataPoints As Long, Optional maxBatchSize As Long = 3000) As Long
    On Error GoTo ErrorHandler
    
    If totalDataPoints <= maxBatchSize Then
        CalculateBatchCount = 1
    Else
        CalculateBatchCount = ((totalDataPoints - 1) \ maxBatchSize) + 1
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CalculateBatchCount", Err.Description)
    CalculateBatchCount = 1
End Function