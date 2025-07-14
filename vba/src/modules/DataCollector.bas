Attribute VB_Name = "DataCollector"
'******************************************************************************
' 楽天MS2RSS株価データコレクター - データ取得モジュール
' 
' 説明: 楽天証券MarketSpeed2のRSS API経由で株価データを取得
' 作成者: Claude Code
' バージョン: 1.0.0
'******************************************************************************

Option Explicit

' データ取得の主関数
Public Function CollectStockData(stockCode As String, timeFrame As String, _
                                startDate As Date, endDate As Date, _
                                Optional outputPath As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim totalBars As Long
    Dim currentDate As Date
    Dim dataArray As Variant
    Dim csvData As String
    
    ' ログ出力
    Call LogMessage("INFO", "データ取得開始: " & stockCode & " (" & timeFrame & ")")
    
    ' 期間の妥当性チェック
    If startDate > endDate Then
        Call LogMessage("ERROR", "開始日が終了日より後です")
        CollectStockData = False
        Exit Function
    End If
    
    ' 銘柄コードの形式チェック
    If Not ValidateStockCode(stockCode) Then
        Call LogMessage("ERROR", "無効な銘柄コード: " & stockCode)
        CollectStockData = False
        Exit Function
    End If
    
    ' RSS Chart APIでデータ取得
    dataArray = GetChartDataFromAPI(stockCode, timeFrame, startDate, endDate)
    
    If IsEmpty(dataArray) Then
        Call LogMessage("ERROR", "データ取得に失敗しました: " & stockCode)
        CollectStockData = False
        Exit Function
    End If
    
    ' CSVデータに変換
    csvData = ConvertToCSV(dataArray)
    
    ' ファイル出力
    If outputPath = "" Then
        outputPath = GenerateOutputFilename(stockCode, timeFrame, startDate, endDate)
    End If
    
    result = SaveCSVFile(csvData, outputPath)
    
    If result Then
        Call LogMessage("INFO", "データ取得完了: " & outputPath)
        CollectStockData = True
    Else
        Call LogMessage("ERROR", "ファイル保存に失敗しました: " & outputPath)
        CollectStockData = False
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "DataCollector.CollectStockData: " & Err.Description)
    CollectStockData = False
End Function

' RSS Chart APIからデータを取得
Private Function GetChartDataFromAPI(stockCode As String, timeFrame As String, _
                                   startDate As Date, endDate As Date) As Variant
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim headerRange As Range
    Dim result As Variant
    Dim tempSheetName As String
    Dim maxBars As Long
    
    ' 一時ワークシートを作成
    tempSheetName = "TempData_" & Format(Now, "hhmmss")
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = tempSheetName
    
    ' ヘッダー行を設定
    ws.Range("A1:F1").Value = Array("DateTime", "Open", "High", "Low", "Close", "Volume")
    Set headerRange = ws.Range("A1:F1")
    
    ' API制限を考慮した最大取得本数
    maxBars = 3000
    
    ' RSS Chart API呼び出し（VBA版）
    ' 過去データ取得の場合はRssChartPast_vを使用
    If startDate < Date Then
        result = Application.WorksheetFunction.RssChartPast_v( _
            headerRange, stockCode, timeFrame, Format(startDate, "YYYYMMDD"), maxBars)
    Else
        result = Application.WorksheetFunction.RssChart_v( _
            headerRange, stockCode, timeFrame, maxBars)
    End If
    
    ' 結果をコピー
    GetChartDataFromAPI = result
    
    ' 一時ワークシートを削除
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "GetChartDataFromAPI: " & Err.Description)
    
    ' エラー時も一時ワークシートを削除
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    GetChartDataFromAPI = Empty
End Function

' 銘柄コードの妥当性チェック
Private Function ValidateStockCode(stockCode As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim codePattern As String
    Dim marketPart As String
    Dim codePart As String
    
    ' 銘柄コード.市場 の形式をチェック
    If InStr(stockCode, ".") > 0 Then
        codePart = Split(stockCode, ".")(0)
        marketPart = Split(stockCode, ".")(1)
        
        ' 市場コードの妥当性チェック
        Select Case UCase(marketPart)
            Case "T", "JAX", "JNX", "CHJ"
                ' 有効な市場コード
            Case Else
                ValidateStockCode = False
                Exit Function
        End Select
    Else
        codePart = stockCode
    End If
    
    ' 数値部分のチェック（4桁または5桁）
    If Len(codePart) >= 4 And Len(codePart) <= 5 And IsNumeric(codePart) Then
        ValidateStockCode = True
    Else
        ValidateStockCode = False
    End If
    
    Exit Function
    
ErrorHandler:
    ValidateStockCode = False
End Function

' 配列データをCSV形式に変換
Private Function ConvertToCSV(dataArray As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim csvString As String
    Dim i As Long, j As Long
    Dim rowData As String
    
    ' ヘッダー行
    csvString = "DateTime,Open,High,Low,Close,Volume" & vbCrLf
    
    ' データ行
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        rowData = ""
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            If j > LBound(dataArray, 2) Then rowData = rowData & ","
            rowData = rowData & CStr(dataArray(i, j))
        Next j
        csvString = csvString & rowData & vbCrLf
    Next i
    
    ConvertToCSV = csvString
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "ConvertToCSV: " & Err.Description)
    ConvertToCSV = ""
End Function

' 出力ファイル名を生成
Private Function GenerateOutputFilename(stockCode As String, timeFrame As String, _
                                      startDate As Date, endDate As Date) As String
    Dim fileName As String
    Dim outputDir As String
    
    ' 出力ディレクトリ
    outputDir = ThisWorkbook.Path & "\output\csv\"
    
    ' ディレクトリが存在しない場合は作成
    If Dir(outputDir, vbDirectory) = "" Then
        MkDir outputDir
    End If
    
    ' ファイル名生成
    fileName = Replace(stockCode, ".", "_") & "_" & timeFrame & "_" & _
               Format(startDate, "YYYYMMDD") & "-" & Format(endDate, "YYYYMMDD") & ".csv"
    
    GenerateOutputFilename = outputDir & fileName
End Function

' CSVファイルを保存
Private Function SaveCSVFile(csvData As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, csvData;
    Close #fileNum
    
    SaveCSVFile = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage("ERROR", "SaveCSVFile: " & Err.Description)
    SaveCSVFile = False
End Function

' 複数銘柄の一括処理
Public Function CollectMultipleStocks(stockCodes As String, timeFrame As String, _
                                    startDate As Date, endDate As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim stocks() As String
    Dim i As Long
    Dim successCount As Long
    Dim totalCount As Long
    
    ' 銘柄コードを分割
    stocks = Split(Replace(stockCodes, " ", ""), ",")
    totalCount = UBound(stocks) + 1
    
    Call LogMessage("INFO", "複数銘柄データ取得開始: " & totalCount & "銘柄")
    
    ' 各銘柄を処理
    For i = 0 To UBound(stocks)
        If Trim(stocks(i)) <> "" Then
            If CollectStockData(Trim(stocks(i)), timeFrame, startDate, endDate) Then
                successCount = successCount + 1
            End If
            
            ' 進捗更新（後でフォームから呼び出し可能にする）
            DoEvents
        End If
    Next i
    
    Call LogMessage("INFO", "複数銘柄データ取得完了: " & successCount & "/" & totalCount)
    CollectMultipleStocks = (successCount = totalCount)
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "CollectMultipleStocks: " & Err.Description)
    CollectMultipleStocks = False
End Function