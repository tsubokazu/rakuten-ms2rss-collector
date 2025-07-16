Attribute VB_Name = "DataCollector"
'******************************************************************************
' 楽天MS2RSS株価データコレクター - データ取得モジュール（簡易版）
' 
' 説明: 楽天証券MarketSpeed2のRSS API経由で株価データを取得
' 作成者: Claude Code
' バージョン: 1.0.0
'******************************************************************************

Option Explicit

' データ取得の主関数（簡易版）
Public Function CollectStockData(stockCode As String, timeFrame As String, _
                                startDate As Date, endDate As Date) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 基本的な妥当性チェック
    If stockCode = "" Or timeFrame = "" Then
        MsgBox "銘柄コードまたは足種が指定されていません", vbExclamation
        CollectStockData = False
        Exit Function
    End If
    
    ' 簡易テスト（実際のRSS APIは呼び出さない）
    MsgBox "データ取得テスト実行中..." & vbCrLf & _
           "銘柄: " & stockCode & vbCrLf & _
           "足種: " & timeFrame & vbCrLf & _
           "期間: " & Format(startDate, "MM/DD") & " - " & Format(endDate, "MM/DD"), _
           vbInformation, "テスト実行"
    
    ' 成功として扱う
    CollectStockData = True
    Exit Function
    
ErrorHandler:
    MsgBox "CollectStockData エラー: " & Err.Description, vbCritical
    CollectStockData = False
End Function

' 複数銘柄の一括処理（簡易版）
Public Function CollectMultipleStocks(stockCodes As String, timeFrame As String, _
                                    startDate As Date, endDate As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim stocks() As String
    Dim i As Long
    Dim successCount As Long
    
    ' 銘柄コードを分割
    stocks = Split(Replace(stockCodes, " ", ""), ",")
    
    ' 各銘柄を処理
    For i = 0 To UBound(stocks)
        If Trim(stocks(i)) <> "" Then
            If CollectStockData(Trim(stocks(i)), timeFrame, startDate, endDate) Then
                successCount = successCount + 1
            End If
        End If
    Next i
    
    MsgBox "処理完了: " & successCount & "/" & (UBound(stocks) + 1) & " 銘柄", vbInformation
    CollectMultipleStocks = True
    
    Exit Function
    
ErrorHandler:
    MsgBox "CollectMultipleStocks エラー: " & Err.Description, vbCritical
    CollectMultipleStocks = False
End Function