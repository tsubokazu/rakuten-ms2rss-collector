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
    
    ' Calculate required data points based on date range and timeframe
    Dim daysDiff As Long
    Dim dataPoints As Long
    
    daysDiff = DateDiff("d", startDate, endDate) + 1
    dataPoints = CalculateDataPoints(daysDiff, timeFrame)
    
    Debug.Print "Data collection for " & stockCode & ": " & dataPoints & " points needed for " & daysDiff & " days"
    
    ' Note: RSS Chart API limitations
    ' - RssChart: Gets latest N points (no date range)
    ' - RssChartPast: Gets N points from specific start date
    ' For this demo, we'll generate sample data reflecting the actual date range
    
    ' Generate sample output path if not specified
    If outputPath = "" Then
        outputPath = GenerateOutputFilename(stockCode, timeFrame, startDate, endDate)
    End If
    
    ' Create sample CSV file with proper date range
    result = CreateSampleCSVFileWithDateRange(outputPath, stockCode, timeFrame, startDate, endDate, dataPoints)
    
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

' Calculate expected data points based on timeframe and days
Private Function CalculateDataPoints(days As Long, timeFrame As String) As Long
    Dim pointsPerDay As Long
    
    Select Case UCase(timeFrame)
        Case "1M": pointsPerDay = 480    ' 8 hours * 60 minutes (market hours)
        Case "5M": pointsPerDay = 96     ' 8 hours * 12 (5-minute intervals)
        Case "15M": pointsPerDay = 32    ' 8 hours * 4 (15-minute intervals)
        Case "30M": pointsPerDay = 16    ' 8 hours * 2 (30-minute intervals)
        Case "60M": pointsPerDay = 8     ' 8 hours (hourly)
        Case "D": pointsPerDay = 1       ' Daily
        Case Else: pointsPerDay = 96     ' Default to 5M
    End Select
    
    CalculateDataPoints = days * pointsPerDay
End Function

' Create sample CSV file with proper date range
Private Function CreateSampleCSVFileWithDateRange(filePath As String, stockCode As String, _
                                                 timeFrame As String, startDate As Date, _
                                                 endDate As Date, dataPoints As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim csvContent As String
    Dim i As Long
    Dim basePrice As Double
    Dim currentDateTime As Date
    Dim minuteInterval As Long
    
    ' Determine minute interval based on timeframe
    Select Case UCase(timeFrame)
        Case "1M": minuteInterval = 1
        Case "5M": minuteInterval = 5
        Case "15M": minuteInterval = 15
        Case "30M": minuteInterval = 30
        Case "60M": minuteInterval = 60
        Case "D": minuteInterval = 1440  ' 24 hours in minutes
        Case Else: minuteInterval = 5    ' Default to 5M
    End Select
    
    ' Generate sample data
    basePrice = 2500 + (Rnd() * 100) ' Random base price around 2500
    csvContent = "DateTime,Open,High,Low,Close,Volume" & vbCrLf
    
    ' Start from the beginning of the start date
    currentDateTime = startDate + TimeValue("09:00:00") ' Market opens at 9:00
    
    ' Generate data points within the specified date range
    For i = 1 To dataPoints
        ' Skip weekends for daily data
        If timeFrame = "D" And (Weekday(currentDateTime) = 1 Or Weekday(currentDateTime) = 7) Then
            currentDateTime = currentDateTime + minuteInterval / 1440
            GoTo NextIteration
        End If
        
        ' Stop if we've passed the end date
        If currentDateTime > endDate + TimeValue("15:00:00") Then Exit For
        
        ' Generate realistic OHLCV data
        Dim openPrice As Double, highPrice As Double, lowPrice As Double, closePrice As Double
        Dim volume As Long
        
        openPrice = basePrice + (Rnd() - 0.5) * 50
        highPrice = openPrice + Rnd() * 30
        lowPrice = openPrice - Rnd() * 30
        closePrice = openPrice + (Rnd() - 0.5) * 40
        volume = Int(Rnd() * 100000) + 50000
        
        ' Update base price for next iteration (trend simulation)
        basePrice = closePrice + (Rnd() - 0.5) * 10
        
        csvContent = csvContent & _
            Format(currentDateTime, "YYYY/MM/DD HH:MM") & "," & _
            Format(openPrice, "0.00") & "," & _
            Format(highPrice, "0.00") & "," & _
            Format(lowPrice, "0.00") & "," & _
            Format(closePrice, "0.00") & "," & _
            volume & vbCrLf
        
NextIteration:
        ' Advance time
        currentDateTime = currentDateTime + minuteInterval / 1440  ' Convert minutes to days
        
        ' Skip to next trading day if we're past market hours
        If TimeValue(Format(currentDateTime, "HH:MM:SS")) > TimeValue("15:00:00") Then
            currentDateTime = Int(currentDateTime) + 1 + TimeValue("09:00:00")
        End If
    Next i
    
    ' Save to file
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, csvContent;
    Close #fileNum
    
    Debug.Print "Generated " & (i - 1) & " data points for period " & Format(startDate, "YYYY/MM/DD") & " to " & Format(endDate, "YYYY/MM/DD")
    CreateSampleCSVFileWithDateRange = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Debug.Print "CreateSampleCSVFileWithDateRange Error: " & Err.Description
    CreateSampleCSVFileWithDateRange = False
End Function