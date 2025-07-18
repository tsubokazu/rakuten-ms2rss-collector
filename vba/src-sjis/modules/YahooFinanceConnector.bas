'******************************************************************************
' Yahoo Finance API Connector - Data Collection Module
' 
' Description: Stock data collection via Yahoo Finance API
' Author: Claude Code
' Version: 1.0.0
'******************************************************************************

Option Explicit

' Yahoo Finance API data collection function
Public Function CollectStockDataYahoo(stockCode As String, timeFrame As String, _
                                     startDate As Date, endDate As Date, _
                                     Optional outputPath As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim pythonCommand As String
    Dim outputFile As String
    Dim resultText As String
    
    ' Input validation
    If Not ValidateStockCode(stockCode) Then
        Call LogMessage(LOG_ERROR, "Invalid stock code: " & stockCode)
        CollectStockDataYahoo = False
        Exit Function
    End If
    
    If startDate > endDate Then
        Call LogMessage(LOG_ERROR, "Start date is later than end date")
        CollectStockDataYahoo = False
        Exit Function
    End If
    
    ' Normalize timeframe format
    timeFrame = NormalizeTimeFrame(timeFrame)
    
    Call LogMessage(LOG_INFO, "Yahoo Finance data collection start: " & stockCode & " (" & timeFrame & ") from " & Format(startDate, "YYYY-MM-DD") & " to " & Format(endDate, "YYYY-MM-DD"))
    
    ' Add .T suffix if not present (for Japanese stocks)
    If Not InStr(stockCode, ".") > 0 Then
        stockCode = stockCode & ".T"
    End If
    
    ' Generate output path if not specified
    If outputPath = "" Then
        Dim outputDir As String
        outputDir = ThisWorkbook.Path & "\output\csv\"
        If Not EnsureDirectoryExists(outputDir) Then
            Call LogMessage(LOG_ERROR, "Failed to create output directory: " & outputDir)
            CollectStockDataYahoo = False
            Exit Function
        End If
        
        Dim stockCodeClean As String
        stockCodeClean = Replace(stockCode, ".T", "")
        Dim startDateStr As String
        startDateStr = Format(startDate, "YYYYMMDD")
        Dim endDateStr As String
        endDateStr = Format(endDate, "YYYYMMDD")
        
        outputFile = outputDir & stockCodeClean & "_" & timeFrame & "_" & startDateStr & "-" & endDateStr & ".csv"
    Else
        outputFile = outputPath
    End If
    
    ' Call Yahoo Finance API via Python
    result = CallYahooFinanceAPI(stockCode, timeFrame, startDate, endDate, outputFile)
    
    If result Then
        Call LogMessage(LOG_INFO, "Yahoo Finance data collection completed successfully")
        Call LogMessage(LOG_INFO, "Output file: " & outputFile)
        CollectStockDataYahoo = True
    Else
        Call LogMessage(LOG_ERROR, "Yahoo Finance data collection failed")
        CollectStockDataYahoo = False
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CollectStockDataYahoo", Err.Description, "StockCode: " & stockCode & ", TimeFrame: " & timeFrame)
    CollectStockDataYahoo = False
End Function

' Call Yahoo Finance API via Python subprocess
Private Function CallYahooFinanceAPI(stockCode As String, timeFrame As String, _
                                    startDate As Date, endDate As Date, _
                                    outputFile As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim pythonCommand As String
    Dim startDateStr As String
    Dim endDateStr As String
    Dim commandResult As String
    
    ' Format dates for API call
    startDateStr = Format(startDate, "YYYY-MM-DD")
    endDateStr = Format(endDate, "YYYY-MM-DD")
    
    ' Build Python command
    ' Use absolute path to Python in virtual environment
    Dim pythonPath As String
    pythonPath = ThisWorkbook.Path & "\.venv\Scripts\python.exe"
    
    ' If virtual environment doesn't exist, try system python
    If Dir(pythonPath) = "" Then
        pythonPath = "python"
    End If
    
    ' Build command for Yahoo Finance client
    pythonCommand = pythonPath & " -m yahoo_finance_client.vba_bridge """ & _
                   stockCode & """ """ & timeFrame & """ """ & _
                   startDateStr & """ """ & endDateStr & """ """ & _
                   outputFile & """"
    
    Call LogMessage(LOG_INFO, "Executing Python command: " & pythonCommand)
    
    ' Execute command and capture result
    commandResult = ExecuteCommand(pythonCommand)
    
    ' Parse JSON result
    Dim jsonResult As Object
    Set jsonResult = ParseJSON(commandResult)
    
    If Not jsonResult Is Nothing Then
        If jsonResult("success") = True Then
            Call LogMessage(LOG_INFO, "Yahoo Finance API call successful")
            Call LogMessage(LOG_INFO, "Records retrieved: " & jsonResult("record_count"))
            Call LogMessage(LOG_INFO, "Output file: " & jsonResult("output_file"))
            
            ' Log date range if available
            If Not IsEmpty(jsonResult("date_range")) Then
                Call LogMessage(LOG_INFO, "Date range: " & jsonResult("date_range")("start") & " to " & jsonResult("date_range")("end"))
            End If
            
            CallYahooFinanceAPI = True
        Else
            Call LogMessage(LOG_ERROR, "Yahoo Finance API call failed: " & jsonResult("error"))
            CallYahooFinanceAPI = False
        End If
    Else
        Call LogMessage(LOG_ERROR, "Failed to parse Yahoo Finance API response")
        Call LogMessage(LOG_ERROR, "Raw response: " & commandResult)
        CallYahooFinanceAPI = False
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CallYahooFinanceAPI", Err.Description, "Command: " & pythonCommand)
    CallYahooFinanceAPI = False
End Function

' Execute shell command and return result
Private Function ExecuteCommand(command As String) As String
    On Error GoTo ErrorHandler
    
    Dim shell As Object
    Dim exec As Object
    Dim result As String
    
    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec(command)
    
    ' Wait for command to complete and capture output
    Do While exec.Status = 0
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    
    result = exec.StdOut.ReadAll
    
    If exec.ExitCode <> 0 Then
        Dim errorOutput As String
        errorOutput = exec.StdErr.ReadAll
        Call LogMessage(LOG_ERROR, "Command execution failed with exit code: " & exec.ExitCode)
        Call LogMessage(LOG_ERROR, "Error output: " & errorOutput)
    End If
    
    ExecuteCommand = result
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("ExecuteCommand", Err.Description, "Command: " & command)
    ExecuteCommand = ""
End Function

' Simple JSON parser for result parsing
Private Function ParseJSON(jsonString As String) As Object
    On Error GoTo ErrorHandler
    
    ' Try to use built-in JSON parser (Excel 2016+)
    Dim jsonObject As Object
    
    ' Remove any extra whitespace
    jsonString = Trim(jsonString)
    
    ' Simple JSON parsing - look for success field
    If InStr(jsonString, """success"": true") > 0 Then
        ' Create a simple object to simulate JSON parsing
        Set jsonObject = CreateObject("Scripting.Dictionary")
        jsonObject("success") = True
        
        ' Extract record count
        Dim recordCountPos As Integer
        recordCountPos = InStr(jsonString, """record_count"": ")
        If recordCountPos > 0 Then
            Dim recordCountStr As String
            recordCountStr = Mid(jsonString, recordCountPos + 16)
            recordCountStr = Left(recordCountStr, InStr(recordCountStr, ",") - 1)
            jsonObject("record_count") = CLng(recordCountStr)
        End If
        
        ' Extract output file
        Dim outputFilePos As Integer
        outputFilePos = InStr(jsonString, """output_file"": """)
        If outputFilePos > 0 Then
            Dim outputFileStr As String
            outputFileStr = Mid(jsonString, outputFilePos + 15)
            outputFileStr = Left(outputFileStr, InStr(outputFileStr, """") - 1)
            jsonObject("output_file") = outputFileStr
        End If
        
        ' Create date range object
        Dim dateRangeDict As Object
        Set dateRangeDict = CreateObject("Scripting.Dictionary")
        
        ' Extract start date
        Dim startDatePos As Integer
        startDatePos = InStr(jsonString, """start"": """)
        If startDatePos > 0 Then
            Dim startDateStr As String
            startDateStr = Mid(jsonString, startDatePos + 10)
            startDateStr = Left(startDateStr, InStr(startDateStr, """") - 1)
            dateRangeDict("start") = startDateStr
        End If
        
        ' Extract end date
        Dim endDatePos As Integer
        endDatePos = InStr(jsonString, """end"": """)
        If endDatePos > 0 Then
            Dim endDateStr As String
            endDateStr = Mid(jsonString, endDatePos + 8)
            endDateStr = Left(endDateStr, InStr(endDateStr, """") - 1)
            dateRangeDict("end") = endDateStr
        End If
        
        jsonObject("date_range") = dateRangeDict
        
    Else
        ' Parse error case
        Set jsonObject = CreateObject("Scripting.Dictionary")
        jsonObject("success") = False
        
        ' Extract error message
        Dim errorPos As Integer
        errorPos = InStr(jsonString, """error"": """)
        If errorPos > 0 Then
            Dim errorStr As String
            errorStr = Mid(jsonString, errorPos + 10)
            errorStr = Left(errorStr, InStr(errorStr, """") - 1)
            jsonObject("error") = errorStr
        Else
            jsonObject("error") = "Unknown error"
        End If
    End If
    
    Set ParseJSON = jsonObject
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("ParseJSON", Err.Description, "JSON: " & Left(jsonString, 200))
    Set ParseJSON = Nothing
End Function

' Collect data for multiple stocks using Yahoo Finance API
Public Function CollectMultipleStocksYahoo(stockCodes As String, timeFrame As String, _
                                          startDate As Date, endDate As Date, _
                                          Optional outputPath As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim stockCodeArray As Variant
    Dim i As Integer
    Dim currentStock As String
    Dim successCount As Integer
    
    ' Split stock codes
    stockCodeArray = Split(stockCodes, ",")
    successCount = 0
    
    Call LogMessage(LOG_INFO, "Multiple stocks data collection start: " & stockCodes)
    
    ' Process each stock
    For i = 0 To UBound(stockCodeArray)
        currentStock = Trim(stockCodeArray(i))
        
        If currentStock <> "" Then
            Call LogMessage(LOG_INFO, "Processing stock: " & currentStock)
            
            result = CollectStockDataYahoo(currentStock, timeFrame, startDate, endDate, outputPath)
            
            If result Then
                successCount = successCount + 1
                Call LogMessage(LOG_INFO, "Successfully processed: " & currentStock)
            Else
                Call LogMessage(LOG_ERROR, "Failed to process: " & currentStock)
            End If
        End If
    Next i
    
    Call LogMessage(LOG_INFO, "Multiple stocks collection completed: " & successCount & " of " & (UBound(stockCodeArray) + 1) & " stocks processed")
    
    CollectMultipleStocksYahoo = (successCount > 0)
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CollectMultipleStocksYahoo", Err.Description, "StockCodes: " & stockCodes)
    CollectMultipleStocksYahoo = False
End Function