Attribute VB_Name = "CSVExporter"
'******************************************************************************
' 楽天MS2RSS株価データコレクター - CSV出力モジュール
' 
' 説明: 株価データのCSV形式出力機能
' 作成者: Claude Code
' バージョン: 1.0.0
'******************************************************************************

Option Explicit

' CSV出力設定
Private Type CSVConfig
    IncludeHeader As Boolean
    DateTimeFormat As String
    DecimalPlaces As Integer
    Encoding As String
    Delimiter As String
End Type

Private config As CSVConfig

' モジュール初期化
Private Sub InitializeCSVConfig()
    config.IncludeHeader = True
    config.DateTimeFormat = "YYYY-MM-DD HH:MM:SS"
    config.DecimalPlaces = 2
    config.Encoding = "UTF-8"
    config.Delimiter = ","
End Sub

' 株価データをCSV形式で出力
Public Function ExportStockDataToCSV(stockData As Variant, filePath As String, _
                                    Optional includeHeader As Boolean = True) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim csvContent As String
    Dim fileNum As Integer
    
    Call InitializeCSVConfig
    config.IncludeHeader = includeHeader
    
    ' CSV内容を生成
    csvContent = GenerateCSVContent(stockData)
    
    If csvContent = "" Then
        Call LogMessage("ERROR", "CSV内容の生成に失敗しました")
        ExportStockDataToCSV = False
        Exit Function
    End If
    
    ' ファイルに書き込み
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, csvContent;
    Close #fileNum
    
    Call LogMessage("INFO", "CSVファイル出力完了: " & filePath)
    ExportStockDataToCSV = True
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage("ERROR", "ExportStockDataToCSV: " & Err.Description)
    ExportStockDataToCSV = False
End Function

' CSV内容を生成
Private Function GenerateCSVContent(stockData As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim content As String
    Dim i As Long, j As Long
    Dim rowContent As String
    Dim cellValue As String
    
    ' ヘッダー行を追加
    If config.IncludeHeader Then
        content = GenerateCSVHeader() & vbCrLf
    End If
    
    ' データ行を処理
    If IsArray(stockData) Then
        For i = LBound(stockData, 1) To UBound(stockData, 1)
            rowContent = ""
            For j = LBound(stockData, 2) To UBound(stockData, 2)
                If j > LBound(stockData, 2) Then
                    rowContent = rowContent & config.Delimiter
                End If
                
                cellValue = FormatCellValue(stockData(i, j), j)
                rowContent = rowContent & cellValue
            Next j
            content = content & rowContent & vbCrLf
        Next i
    End If
    
    GenerateCSVContent = content
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "GenerateCSVContent: " & Err.Description)
    GenerateCSVContent = ""
End Function

' CSVヘッダーを生成
Private Function GenerateCSVHeader() As String
    Dim header As String
    
    header = "DateTime" & config.Delimiter & _
             "Open" & config.Delimiter & _
             "High" & config.Delimiter & _
             "Low" & config.Delimiter & _
             "Close" & config.Delimiter & _
             "Volume"
    
    GenerateCSVHeader = header
End Function

' セル値をフォーマット
Private Function FormatCellValue(value As Variant, columnIndex As Integer) As String
    On Error GoTo ErrorHandler
    
    Dim formattedValue As String
    
    Select Case columnIndex
        Case 0 ' DateTime列
            If IsDate(value) Then
                formattedValue = Format(value, "YYYY-MM-DD HH:MM:SS")
            Else
                formattedValue = CStr(value)
            End If
            
        Case 1 To 4 ' OHLC列（数値）
            If IsNumeric(value) Then
                formattedValue = Format(value, "0." & String(config.DecimalPlaces, "0"))
            Else
                formattedValue = CStr(value)
            End If
            
        Case 5 ' Volume列（整数）
            If IsNumeric(value) Then
                formattedValue = Format(value, "0")
            Else
                formattedValue = CStr(value)
            End If
            
        Case Else
            formattedValue = CStr(value)
    End Select
    
    ' CSVエスケープ処理
    If InStr(formattedValue, config.Delimiter) > 0 Or _
       InStr(formattedValue, """") > 0 Or _
       InStr(formattedValue, vbCrLf) > 0 Then
        formattedValue = """" & Replace(formattedValue, """", """""") & """"
    End If
    
    FormatCellValue = formattedValue
    Exit Function
    
ErrorHandler:
    FormatCellValue = CStr(value)
End Function

' バッチCSV出力（複数銘柄対応）
Public Function ExportMultipleStocksToCSV(stockDataArray As Variant, _
                                         baseFilePath As String, _
                                         stockCodes As Variant) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim filePath As String
    Dim fileName As String
    Dim fileExtension As String
    Dim baseName As String
    Dim successCount As Long
    
    If Not IsArray(stockDataArray) Or Not IsArray(stockCodes) Then
        Call LogMessage("ERROR", "入力データが配列ではありません")
        ExportMultipleStocksToCSV = False
        Exit Function
    End If
    
    ' ファイルパスを分解
    fileExtension = ".csv"
    If InStr(baseFilePath, ".") > 0 Then
        baseName = Left(baseFilePath, InStrRev(baseFilePath, ".") - 1)
        fileExtension = Right(baseFilePath, Len(baseFilePath) - InStrRev(baseFilePath, ".") + 1)
    Else
        baseName = baseFilePath
    End If
    
    ' 各銘柄のデータを出力
    For i = LBound(stockCodes) To UBound(stockCodes)
        fileName = baseName & "_" & Replace(stockCodes(i), ".", "_") & fileExtension
        
        If ExportStockDataToCSV(stockDataArray(i), fileName) Then
            successCount = successCount + 1
        End If
    Next i
    
    Call LogMessage("INFO", "バッチCSV出力完了: " & successCount & "/" & (UBound(stockCodes) - LBound(stockCodes) + 1))
    ExportMultipleStocksToCSV = (successCount = (UBound(stockCodes) - LBound(stockCodes) + 1))
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "ExportMultipleStocksToCSV: " & Err.Description)
    ExportMultipleStocksToCSV = False
End Function

' CSV設定を変更
Public Sub SetCSVConfig(Optional includeHeader As Boolean = True, _
                       Optional dateTimeFormat As String = "YYYY-MM-DD HH:MM:SS", _
                       Optional decimalPlaces As Integer = 2, _
                       Optional delimiter As String = ",")
    
    config.IncludeHeader = includeHeader
    config.DateTimeFormat = dateTimeFormat
    config.DecimalPlaces = decimalPlaces
    config.Delimiter = delimiter
End Sub

' ファイルサイズを取得（MB単位）
Public Function GetFileSize(filePath As String) As Double
    On Error GoTo ErrorHandler
    
    Dim fileSize As Long
    fileSize = FileLen(filePath)
    GetFileSize = fileSize / 1024 / 1024 ' MB換算
    
    Exit Function
    
ErrorHandler:
    GetFileSize = 0
End Function

' CSVファイルの妥当性チェック
Public Function ValidateCSVFile(filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim firstLine As String
    Dim expectedHeader As String
    
    If Dir(filePath) = "" Then
        Call LogMessage("ERROR", "ファイルが存在しません: " & filePath)
        ValidateCSVFile = False
        Exit Function
    End If
    
    ' ファイルサイズチェック（空ファイルでないか）
    If GetFileSize(filePath) = 0 Then
        Call LogMessage("ERROR", "ファイルが空です: " & filePath)
        ValidateCSVFile = False
        Exit Function
    End If
    
    ' ヘッダー行をチェック
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Line Input #fileNum, firstLine
    Close #fileNum
    
    expectedHeader = GenerateCSVHeader()
    If firstLine <> expectedHeader Then
        Call LogMessage("WARN", "ヘッダー行が期待値と異なります: " & filePath)
    End If
    
    ValidateCSVFile = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage("ERROR", "ValidateCSVFile: " & Err.Description)
    ValidateCSVFile = False
End Function