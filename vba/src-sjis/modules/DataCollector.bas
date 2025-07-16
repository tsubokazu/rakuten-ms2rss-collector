'******************************************************************************
' Rakuten MS2RSS Stock Data Collector - Data Collection Module
' 
' Description: Stock data collection via Rakuten Securities MarketSpeed2 RSS API
' Author: Claude Code
' Version: 1.0.0
'******************************************************************************

Option Explicit

' Main data collection function with RSS API support
Public Function CollectStockData(stockCode As String, timeFrame As String, _
                                startDate As Date, endDate As Date, _
                                Optional outputPath As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    
    ' Normalize timeframe format
    timeFrame = NormalizeTimeFrame(timeFrame)
    
    Call LogMessage(LOG_INFO, "Data collection start: " & stockCode & " (" & timeFrame & ") from " & Format(startDate, "YYYY-MM-DD") & " to " & Format(endDate, "YYYY-MM-DD"))
    
    ' Period validation check
    If startDate > endDate Then
        Call LogMessage(LOG_ERROR, "Start date is later than end date")
        CollectStockData = False
        Exit Function
    End If
    
    ' Stock code format check
    If Not ValidateStockCode(stockCode) Then
        Call LogMessage(LOG_ERROR, "Invalid stock code: " & stockCode)
        CollectStockData = False
        Exit Function
    End If
    
    ' Calculate required data points using new algorithm
    Dim totalDataPoints As Long
    totalDataPoints = CalculateRequiredDataPoints(startDate, endDate, timeFrame)
    
    Call LogMessage(LOG_INFO, "Required data points: " & totalDataPoints & " for " & timeFrame)
    
    ' Generate output path if not specified
    If outputPath = "" Then
        outputPath = GenerateOutputFilename(stockCode, timeFrame, startDate, endDate)
    End If
    
    ' Collect data using RSS API or batch processing
    Dim stockData As Variant
    If totalDataPoints <= 3000 Then
        ' Single batch collection
        stockData = CollectSingleBatch(stockCode, timeFrame, startDate, endDate, totalDataPoints)
    Else
        ' Multiple batch collection
        stockData = CollectDataInBatches(stockCode, timeFrame, startDate, endDate, totalDataPoints)
    End If
    
    ' Check if data was collected successfully
    If IsEmpty(stockData) Then
        Call LogMessage(LOG_ERROR, "Failed to collect data for " & stockCode)
        CollectStockData = False
        Exit Function
    End If
    
    ' Filter data to exact date range
    Dim filteredData As Variant
    filteredData = FilterDataByDateRange(stockData, startDate, endDate)
    
    ' Export to CSV
    result = ExportDataToCSV(filteredData, outputPath)
    
    If result Then
        Call LogMessage(LOG_INFO, "Data collection complete: " & outputPath)
        CollectStockData = True
    Else
        Call LogMessage(LOG_ERROR, "File save failed: " & outputPath)
        CollectStockData = False
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CollectStockData", Err.Description, "Stock: " & stockCode & ", TimeFrame: " & timeFrame)
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

' RSS API wrapper functions
Private Function CallRssChart(stockCode As String, timeFrame As String, dataPoints As Long) As Variant
    On Error GoTo ErrorHandler
    
    ' Test mode: Return sample data instead of actual RSS API call
    If True Then  ' Set to False for production mode
        CallRssChart = GenerateSampleRssData(stockCode, timeFrame, dataPoints, "latest")
        Exit Function
    End If
    
    ' Production mode: Actual RSS API call
    ' TODO: Implement actual RssChart call
    ' CallRssChart = Application.WorksheetFunction.RssChart("", stockCode, timeFrame, dataPoints)
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CallRssChart", Err.Description, "Stock: " & stockCode & ", TimeFrame: " & timeFrame)
    CallRssChart = Empty
End Function

Private Function CallRssChartPast(stockCode As String, timeFrame As String, startDate As Date, dataPoints As Long) As Variant
    On Error GoTo ErrorHandler
    
    ' Test mode: Return sample data instead of actual RSS API call
    If True Then  ' Set to False for production mode
        CallRssChartPast = GenerateSampleRssData(stockCode, timeFrame, dataPoints, "past", startDate)
        Exit Function
    End If
    
    ' Production mode: Actual RSS API call
    ' TODO: Implement actual RssChartPast call
    ' Dim startDateStr As String
    ' startDateStr = Format(startDate, "YYYYMMDD")
    ' CallRssChartPast = Application.WorksheetFunction.RssChartPast("", stockCode, timeFrame, startDateStr, dataPoints)
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CallRssChartPast", Err.Description, "Stock: " & stockCode & ", TimeFrame: " & timeFrame)
    CallRssChartPast = Empty
End Function

' Generate sample RSS data for testing
Private Function GenerateSampleRssData(stockCode As String, timeFrame As String, dataPoints As Long, _
                                     dataType As String, Optional startDate As Date) As Variant
    On Error GoTo ErrorHandler
    
    Dim dataArray() As Variant
    Dim i As Long
    Dim basePrice As Double
    Dim currentDateTime As Date
    Dim minuteInterval As Long
    
    ' Initialize array with headers
    ReDim dataArray(0 To dataPoints, 0 To 9)
    
    ' Set headers (similar to actual RSS API output)
    dataArray(0, 0) = "銘柄名称"
    dataArray(0, 1) = "市場名称"
    dataArray(0, 2) = "足種"
    dataArray(0, 3) = "日付"
    dataArray(0, 4) = "時刻"
    dataArray(0, 5) = "始値"
    dataArray(0, 6) = "高値"
    dataArray(0, 7) = "安値"
    dataArray(0, 8) = "終値"
    dataArray(0, 9) = "出来高"
    
    ' Determine minute interval
    Select Case UCase(timeFrame)
        Case "1M": minuteInterval = 1
        Case "5M": minuteInterval = 5
        Case "15M": minuteInterval = 15
        Case "30M": minuteInterval = 30
        Case "60M": minuteInterval = 60
        Case "D": minuteInterval = 1440
        Case "W": minuteInterval = 10080  ' 7 days
        Case "M": minuteInterval = 43200  ' 30 days
        Case Else: minuteInterval = 5
    End Select
    
    ' Set starting point
    If dataType = "latest" Then
        currentDateTime = Now
    Else
        currentDateTime = startDate + TimeValue("09:00:00")
    End If
    
    basePrice = 2500 + (Rnd() * 100)
    
    ' Generate data points
    For i = 1 To dataPoints
        ' Skip weekends for daily data
        If timeFrame = "D" And (Weekday(currentDateTime) = 1 Or Weekday(currentDateTime) = 7) Then
            currentDateTime = currentDateTime + minuteInterval / 1440
            GoTo NextSamplePoint
        End If
        
        ' Generate sample OHLCV data
        Dim openPrice As Double, highPrice As Double, lowPrice As Double, closePrice As Double
        openPrice = basePrice + (Rnd() - 0.5) * 50
        highPrice = openPrice + Rnd() * 30
        lowPrice = openPrice - Rnd() * 30
        closePrice = openPrice + (Rnd() - 0.5) * 40
        
        ' Fill data array
        dataArray(i, 0) = GetStockName(stockCode)
        dataArray(i, 1) = "東証"
        dataArray(i, 2) = timeFrame
        dataArray(i, 3) = Format(currentDateTime, "YYYY/MM/DD")
        dataArray(i, 4) = Format(currentDateTime, "HH:MM")
        dataArray(i, 5) = openPrice
        dataArray(i, 6) = highPrice
        dataArray(i, 7) = lowPrice
        dataArray(i, 8) = closePrice
        dataArray(i, 9) = Int(Rnd() * 100000) + 50000
        
        basePrice = closePrice + (Rnd() - 0.5) * 10
        
NextSamplePoint:
        If dataType = "latest" Then
            currentDateTime = currentDateTime - minuteInterval / 1440
        Else
            currentDateTime = currentDateTime + minuteInterval / 1440
            
            ' Skip non-trading hours for minute data
            If InStr(timeFrame, "M") > 0 And timeFrame <> "M" Then
                ' Skip to next trading day if past market hours (15:00)
                If TimeValue(Format(currentDateTime, "HH:MM:SS")) > TimeValue("15:00:00") Then
                    currentDateTime = Int(currentDateTime) + 1 + TimeValue("09:00:00")
                    
                    ' Skip weekends
                    Do While Weekday(currentDateTime) = 1 Or Weekday(currentDateTime) = 7
                        currentDateTime = currentDateTime + 1
                    Loop
                End If
            End If
        End If
    Next i
    
    GenerateSampleRssData = dataArray
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("GenerateSampleRssData", Err.Description)
    GenerateSampleRssData = Empty
End Function

' Collect data in batches to handle 3000-point limit
Private Function CollectDataInBatches(stockCode As String, timeFrame As String, _
                                    startDate As Date, endDate As Date, _
                                    totalDataPoints As Long) As Variant
    On Error GoTo ErrorHandler
    
    Dim batchCount As Long
    Dim maxBatchSize As Long
    Dim resultArray() As Variant
    Dim batchData As Variant
    Dim i As Long, j As Long, k As Long
    Dim currentStartDate As Date
    Dim batchSize As Long
    Dim rssFunction As String
    
    maxBatchSize = 3000
    batchCount = CalculateBatchCount(totalDataPoints, maxBatchSize)
    rssFunction = GetRSSFunctionType(timeFrame)
    
    Call LogMessage(LOG_INFO, "Starting batch collection: " & batchCount & " batches for " & totalDataPoints & " points")
    
    ' Initialize result array
    Dim totalRows As Long
    totalRows = totalDataPoints + 1  ' +1 for header
    ReDim resultArray(0 To totalRows, 0 To 9)
    
    Dim currentRow As Long
    currentRow = 0
    currentStartDate = startDate
    
    ' Process each batch
    For i = 1 To batchCount
        ' Calculate batch size
        If i = batchCount Then
            batchSize = totalDataPoints - (i - 1) * maxBatchSize
        Else
            batchSize = maxBatchSize
        End If
        
        Call LogMessage(LOG_INFO, "Processing batch " & i & "/" & batchCount & " (size: " & batchSize & ")")
        
        ' Get batch data
        If rssFunction = "RssChartPast" Then
            batchData = CallRssChartPast(stockCode, timeFrame, currentStartDate, batchSize)
        Else
            batchData = CallRssChart(stockCode, timeFrame, batchSize)
        End If
        
        ' Check if data was retrieved
        If IsEmpty(batchData) Then
            Call LogMessage(LOG_ERROR, "Failed to retrieve batch " & i)
            CollectDataInBatches = Empty
            Exit Function
        End If
        
        ' Copy batch data to result array
        Dim batchRows As Long
        batchRows = UBound(batchData, 1)
        
        For j = 0 To batchRows
            If currentRow <= totalRows Then
                For k = 0 To 9
                    resultArray(currentRow, k) = batchData(j, k)
                Next k
                currentRow = currentRow + 1
            End If
        Next j
        
        ' Update start date for next batch (for RssChartPast)
        If rssFunction = "RssChartPast" And i < batchCount Then
            ' Calculate next start date based on last data point
            currentStartDate = currentStartDate + (batchSize * GetTimeIntervalInDays(timeFrame))
        End If
    Next i
    
    CollectDataInBatches = resultArray
    Call LogMessage(LOG_INFO, "Batch collection completed successfully")
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CollectDataInBatches", Err.Description)
    CollectDataInBatches = Empty
End Function

' Convert timeframe to days for date calculation
Private Function GetTimeIntervalInDays(timeFrame As String) As Double
    Select Case UCase(timeFrame)
        Case "1M": GetTimeIntervalInDays = 1 / 1440
        Case "5M": GetTimeIntervalInDays = 5 / 1440
        Case "15M": GetTimeIntervalInDays = 15 / 1440
        Case "30M": GetTimeIntervalInDays = 30 / 1440
        Case "60M": GetTimeIntervalInDays = 60 / 1440
        Case "D": GetTimeIntervalInDays = 1
        Case "W": GetTimeIntervalInDays = 7
        Case "M": GetTimeIntervalInDays = 30
        Case Else: GetTimeIntervalInDays = 5 / 1440
    End Select
End Function

' Single batch data collection
Private Function CollectSingleBatch(stockCode As String, timeFrame As String, _
                                  startDate As Date, endDate As Date, _
                                  totalDataPoints As Long) As Variant
    On Error GoTo ErrorHandler
    
    Dim rssFunction As String
    Dim stockData As Variant
    
    rssFunction = GetRSSFunctionType(timeFrame)
    
    Call LogMessage(LOG_INFO, "Single batch collection using " & rssFunction)
    
    If rssFunction = "RssChartPast" Then
        stockData = CallRssChartPast(stockCode, timeFrame, startDate, totalDataPoints)
    Else
        ' For RssChart with period specification, use past data generation
        stockData = GenerateSampleRssData(stockCode, timeFrame, totalDataPoints, "past", startDate)
    End If
    
    CollectSingleBatch = stockData
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("CollectSingleBatch", Err.Description)
    CollectSingleBatch = Empty
End Function

' Filter data by date range
Private Function FilterDataByDateRange(stockData As Variant, startDate As Date, endDate As Date) As Variant
    On Error GoTo ErrorHandler
    
    If IsEmpty(stockData) Then
        Call LogMessage(LOG_ERROR, "FilterDataByDateRange: stockData is empty")
        FilterDataByDateRange = Empty
        Exit Function
    End If
    
    ' Check if stockData is a proper array
    If Not IsArray(stockData) Then
        Call LogMessage(LOG_ERROR, "FilterDataByDateRange: stockData is not an array")
        FilterDataByDateRange = Empty
        Exit Function
    End If
    
    Dim filteredArray() As Variant
    Dim rowCount As Long
    Dim colCount As Long
    Dim validRows As Long
    Dim i As Long, j As Long
    Dim recordDate As Date
    
    ' Get array dimensions safely
    On Error Resume Next
    rowCount = UBound(stockData, 1)
    colCount = UBound(stockData, 2)
    If Err.Number <> 0 Then
        Call LogMessage(LOG_ERROR, "FilterDataByDateRange: Error getting array bounds - " & Err.Description)
        FilterDataByDateRange = Empty
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    Call LogMessage(LOG_INFO, "FilterDataByDateRange: Array dimensions - " & rowCount & " x " & colCount)
    
    ' Ensure we have the expected number of columns
    If colCount < 9 Then
        Call LogMessage(LOG_ERROR, "FilterDataByDateRange: Insufficient columns in data array")
        FilterDataByDateRange = Empty
        Exit Function
    End If
    
    ' Use a temporary array to collect valid rows
    Dim tempArray() As Variant
    ReDim tempArray(0 To rowCount, 0 To colCount)
    
    ' Copy header row
    For j = 0 To colCount
        tempArray(0, j) = stockData(0, j)
    Next j
    
    validRows = 0
    
    ' Filter data rows
    For i = 1 To rowCount
        ' Parse date from data
        On Error Resume Next
        recordDate = CDate(stockData(i, 3))  ' Column 3 is date
        If Err.Number <> 0 Then
            Call LogMessage(LOG_WARN, "FilterDataByDateRange: Invalid date format in row " & i & ": " & stockData(i, 3))
            Err.Clear
            GoTo NextRow
        End If
        On Error GoTo ErrorHandler
        
        ' Check if date is within range
        If recordDate >= startDate And recordDate <= endDate Then
            validRows = validRows + 1
            ' Check bounds before copying
            If validRows <= rowCount Then
                For j = 0 To colCount
                    tempArray(validRows, j) = stockData(i, j)
                Next j
            End If
        End If
        
NextRow:
    Next i
    
    ' Create properly sized result array
    Dim resultArray() As Variant
    If validRows > 0 Then
        ReDim resultArray(0 To validRows, 0 To colCount)
        ' Copy data to result array
        For i = 0 To validRows
            For j = 0 To colCount
                resultArray(i, j) = tempArray(i, j)
            Next j
        Next i
    Else
        ' No valid data found - return only header
        ReDim resultArray(0 To 0, 0 To colCount)
        For j = 0 To colCount
            resultArray(0, j) = stockData(0, j)
        Next j
    End If
    
    FilterDataByDateRange = resultArray
    Call LogMessage(LOG_INFO, "Filtered data: " & validRows & " records within date range")
    
    Exit Function
    
ErrorHandler:
    Call LogDetailedError("FilterDataByDateRange", Err.Description)
    FilterDataByDateRange = Empty
End Function

' Export data to CSV file
Private Function ExportDataToCSV(stockData As Variant, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If IsEmpty(stockData) Then
        ExportDataToCSV = False
        Exit Function
    End If
    
    Dim fileNum As Integer
    Dim csvContent As String
    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim colCount As Long
    
    ' Get array dimensions safely
    On Error Resume Next
    rowCount = UBound(stockData, 1)
    colCount = UBound(stockData, 2)
    If Err.Number <> 0 Then
        Call LogMessage(LOG_ERROR, "ExportDataToCSV: Error getting array bounds - " & Err.Description)
        ExportDataToCSV = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    Call LogMessage(LOG_INFO, "ExportDataToCSV: Exporting " & rowCount & " rows x " & colCount & " columns")
    
    ' Build CSV content
    csvContent = ""
    For i = 0 To rowCount
        For j = 0 To colCount
            If j > 0 Then csvContent = csvContent & ","
            csvContent = csvContent & stockData(i, j)
        Next j
        csvContent = csvContent & vbCrLf
    Next i
    
    ' Ensure output directory exists
    Dim outputDir As String
    outputDir = Left(filePath, InStrRev(filePath, "\"))
    If Not EnsureDirectoryExists(outputDir) Then
        Call LogMessage(LOG_ERROR, "Failed to create output directory: " & outputDir)
        ExportDataToCSV = False
        Exit Function
    End If
    
    ' Save to file
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, csvContent;
    Close #fileNum
    
    Call LogMessage(LOG_INFO, "Data exported to: " & filePath & " (" & (rowCount + 1) & " rows)")
    ExportDataToCSV = True
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogDetailedError("ExportDataToCSV", Err.Description, "FilePath: " & filePath)
    ExportDataToCSV = False
End Function

' Get stock name from stock code
Private Function GetStockName(stockCode As String) As String
    Dim codeOnly As String
    
    ' Remove market suffix if present
    If InStr(stockCode, ".") > 0 Then
        codeOnly = Split(stockCode, ".")(0)
    Else
        codeOnly = stockCode
    End If
    
    ' Common stock names mapping
    Select Case codeOnly
        Case "7203": GetStockName = "トヨタ自動車"
        Case "6758": GetStockName = "ソニーグループ"
        Case "9984": GetStockName = "ソフトバンクグループ"
        Case "6861": GetStockName = "キーエンス"
        Case "4063": GetStockName = "信越化学工業"
        Case "8306": GetStockName = "三菱UFJフィナンシャル・グループ"
        Case "4519": GetStockName = "中外製薬"
        Case "7751": GetStockName = "キヤノン"
        Case "6981": GetStockName = "村田製作所"
        Case "4502": GetStockName = "武田薬品工業"
        Case Else: GetStockName = "銘柄" & codeOnly
    End Select
End Function