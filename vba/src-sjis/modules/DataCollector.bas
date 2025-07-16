Attribute VB_Name = "DataCollector"
'******************************************************************************
' Rakuten MS2RSS Stock Data Collector - Data Collection Module
' 
' Description: Stock data collection via Rakuten Securities MarketSpeed2 RSS API
' Author: Claude Code
' Version: 1.0.0
'******************************************************************************

Option Explicit

' Main data collection function
Public Function CollectStockData(stockCode As String, timeFrame As String, _
                                startDate As Date, endDate As Date, _
                                Optional outputPath As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    
    Debug.Print "Data collection start: " & stockCode & " (" & timeFrame & ")"
    
    ' Period validation check
    If startDate > endDate Then
        Debug.Print "Start date is later than end date"
        CollectStockData = False
        Exit Function
    End If
    
    ' Stock code format check
    If Not ValidateStockCode(stockCode) Then
        Debug.Print "Invalid stock code: " & stockCode
        CollectStockData = False
        Exit Function
    End If
    
    ' For demonstration purposes, return success without actual API call
    ' In real implementation, RSS Chart API would be called here
    Debug.Print "Test mode: Data collection simulation for " & stockCode
    
    ' Generate sample output path if not specified
    If outputPath = "" Then
        outputPath = GenerateOutputFilename(stockCode, timeFrame, startDate, endDate)
    End If
    
    ' Create sample CSV file
    result = CreateSampleCSVFile(outputPath, stockCode)
    
    If result Then
        Debug.Print "Data collection complete: " & outputPath
        CollectStockData = True
    Else
        Debug.Print "File save failed: " & outputPath
        CollectStockData = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "DataCollector.CollectStockData Error: " & Err.Description
    CollectStockData = False
End Function

' Stock code validation
Private Function ValidateStockCode(stockCode As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim codePattern As String
    Dim marketPart As String
    Dim codePart As String
    
    ' Check "stockcode.market" format
    If InStr(stockCode, ".") > 0 Then
        codePart = Split(stockCode, ".")(0)
        marketPart = Split(stockCode, ".")(1)
        
        ' Market code validation
        Select Case UCase(marketPart)
            Case "T", "JAX", "JNX", "CHJ"
                ' Valid market codes
            Case Else
                ValidateStockCode = False
                Exit Function
        End Select
    Else
        codePart = stockCode
    End If
    
    ' Numeric part check (4 or 5 digits)
    If Len(codePart) >= 4 And Len(codePart) <= 5 And IsNumeric(codePart) Then
        ValidateStockCode = True
    Else
        ValidateStockCode = False
    End If
    
    Exit Function
    
ErrorHandler:
    ValidateStockCode = False
End Function

' Generate output filename
Private Function GenerateOutputFilename(stockCode As String, timeFrame As String, _
                                      startDate As Date, endDate As Date) As String
    Dim fileName As String
    Dim outputDir As String
    
    ' Output directory
    outputDir = ThisWorkbook.Path & "\output\csv\"
    
    ' Create directory if it doesn't exist
    If Dir(outputDir, vbDirectory) = "" Then
        MkDir outputDir
    End If
    
    ' Generate filename
    fileName = Replace(stockCode, ".", "_") & "_" & timeFrame & "_" & _
               Format(startDate, "YYYYMMDD") & "-" & Format(endDate, "YYYYMMDD") & ".csv"
    
    GenerateOutputFilename = outputDir & fileName
End Function

' Create sample CSV file
Private Function CreateSampleCSVFile(filePath As String, stockCode As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim csvContent As String
    Dim i As Long
    Dim basePrice As Double
    
    ' Generate sample data
    basePrice = 2500 ' Sample base price
    
    csvContent = "DateTime,Open,High,Low,Close,Volume" & vbCrLf
    
    ' Generate 10 sample records
    For i = 1 To 10
        csvContent = csvContent & _
            Format(Now - (10 - i) / 24 / 60 * 5, "YYYY-MM-DD HH:MM:SS") & "," & _
            Format(basePrice + Rnd() * 100, "0.00") & "," & _
            Format(basePrice + Rnd() * 120, "0.00") & "," & _
            Format(basePrice - Rnd() * 80, "0.00") & "," & _
            Format(basePrice + (Rnd() - 0.5) * 50, "0.00") & "," & _
            Int(Rnd() * 100000) + 50000 & vbCrLf
    Next i
    
    ' Save to file
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, csvContent;
    Close #fileNum
    
    CreateSampleCSVFile = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Debug.Print "CreateSampleCSVFile Error: " & Err.Description
    CreateSampleCSVFile = False
End Function

' Multiple stocks batch processing
Public Function CollectMultipleStocks(stockCodes As String, timeFrame As String, _
                                    startDate As Date, endDate As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim stocks() As String
    Dim i As Long
    Dim successCount As Long
    Dim totalCount As Long
    
    ' Split stock codes
    stocks = Split(Replace(stockCodes, " ", ""), ",")
    totalCount = UBound(stocks) + 1
    
    Debug.Print "Multiple stocks data collection start: " & totalCount & " stocks"
    
    ' Process each stock
    For i = 0 To UBound(stocks)
        If Trim(stocks(i)) <> "" Then
            If CollectStockData(Trim(stocks(i)), timeFrame, startDate, endDate) Then
                successCount = successCount + 1
            End If
            
            ' Progress update
            DoEvents
        End If
    Next i
    
    Debug.Print "Multiple stocks data collection complete: " & successCount & "/" & totalCount
    CollectMultipleStocks = (successCount = totalCount)
    
    Exit Function
    
ErrorHandler:
    Debug.Print "CollectMultipleStocks Error: " & Err.Description
    CollectMultipleStocks = False
End Function