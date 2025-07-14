VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StockData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' 楽天MS2RSS株価データコレクター - 株価データクラス
' 
' 説明: 株価データの構造定義とメソッド
' 作成者: Claude Code
' バージョン: 1.0.0
'******************************************************************************

Option Explicit

' プライベート変数
Private m_StockCode As String
Private m_StockName As String
Private m_Market As String
Private m_TimeFrame As String
Private m_StartDate As Date
Private m_EndDate As Date
Private m_DataArray As Variant
Private m_RecordCount As Long
Private m_LastUpdate As Date

' プロパティ - 銘柄コード
Public Property Get StockCode() As String
    StockCode = m_StockCode
End Property

Public Property Let StockCode(value As String)
    m_StockCode = UCase(Trim(value))
    
    ' 市場コードを分離
    If InStr(m_StockCode, ".") > 0 Then
        m_Market = Split(m_StockCode, ".")(1)
        m_StockCode = Split(m_StockCode, ".")(0)
    Else
        m_Market = "T" ' デフォルトは東証
    End If
End Property

' プロパティ - 銘柄名
Public Property Get StockName() As String
    StockName = m_StockName
End Property

Public Property Let StockName(value As String)
    m_StockName = value
End Property

' プロパティ - 市場
Public Property Get Market() As String
    Market = m_Market
End Property

Public Property Let Market(value As String)
    m_Market = UCase(value)
End Property

' プロパティ - 足種
Public Property Get TimeFrame() As String
    TimeFrame = m_TimeFrame
End Property

Public Property Let TimeFrame(value As String)
    m_TimeFrame = UCase(value)
End Property

' プロパティ - 開始日
Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Property Let StartDate(value As Date)
    m_StartDate = value
End Property

' プロパティ - 終了日
Public Property Get EndDate() As Date
    EndDate = m_EndDate
End Property

Public Property Let EndDate(value As Date)
    m_EndDate = value
End Property

' プロパティ - データ配列
Public Property Get DataArray() As Variant
    DataArray = m_DataArray
End Property

Public Property Let DataArray(value As Variant)
    m_DataArray = value
    If IsArray(value) Then
        m_RecordCount = UBound(value, 1) - LBound(value, 1) + 1
    Else
        m_RecordCount = 0
    End If
    m_LastUpdate = Now
End Property

' プロパティ - レコード数
Public Property Get RecordCount() As Long
    RecordCount = m_RecordCount
End Property

' プロパティ - 最終更新日時
Public Property Get LastUpdate() As Date
    LastUpdate = m_LastUpdate
End Property

' 完全な銘柄コード（市場付き）を取得
Public Function GetFullStockCode() As String
    If m_Market <> "" And m_Market <> "T" Then
        GetFullStockCode = m_StockCode & "." & m_Market
    Else
        GetFullStockCode = m_StockCode
    End If
End Function

' データの妥当性チェック
Public Function IsValid() As Boolean
    On Error GoTo ErrorHandler
    
    ' 必須項目チェック
    If m_StockCode = "" Then
        Call LogMessage(LOG_ERROR, "銘柄コードが設定されていません")
        IsValid = False
        Exit Function
    End If
    
    If m_TimeFrame = "" Then
        Call LogMessage(LOG_ERROR, "足種が設定されていません")
        IsValid = False
        Exit Function
    End If
    
    ' 日付範囲チェック
    If m_StartDate > m_EndDate Then
        Call LogMessage(LOG_ERROR, "開始日が終了日より後です")
        IsValid = False
        Exit Function
    End If
    
    ' 足種の妥当性チェック
    If Not ValidateTimeFrame(m_TimeFrame) Then
        Call LogMessage(LOG_ERROR, "無効な足種です: " & m_TimeFrame)
        IsValid = False
        Exit Function
    End If
    
    IsValid = True
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "StockData.IsValid: " & Err.Description)
    IsValid = False
End Function

' データをクリア
Public Sub ClearData()
    m_DataArray = Empty
    m_RecordCount = 0
    m_LastUpdate = 0
    Call LogMessage(LOG_DEBUG, "データをクリアしました: " & GetFullStockCode())
End Sub

' 統計情報を取得
Public Function GetStatistics() As String
    On Error GoTo ErrorHandler
    
    Dim stats As String
    Dim minPrice As Double, maxPrice As Double
    Dim totalVolume As Long
    Dim i As Long
    
    If Not IsArray(m_DataArray) Or m_RecordCount = 0 Then
        GetStatistics = "データがありません"
        Exit Function
    End If
    
    ' 初期値設定
    minPrice = m_DataArray(LBound(m_DataArray, 1), 3) ' Low
    maxPrice = m_DataArray(LBound(m_DataArray, 1), 2) ' High
    totalVolume = 0
    
    ' 統計計算
    For i = LBound(m_DataArray, 1) To UBound(m_DataArray, 1)
        If m_DataArray(i, 3) < minPrice Then minPrice = m_DataArray(i, 3) ' Low
        If m_DataArray(i, 2) > maxPrice Then maxPrice = m_DataArray(i, 2) ' High
        totalVolume = totalVolume + m_DataArray(i, 5) ' Volume
    Next i
    
    ' 統計文字列作成
    stats = "【" & GetFullStockCode() & " 統計情報】" & vbCrLf
    stats = stats & "期間: " & Format(m_StartDate, "YYYY/MM/DD") & " - " & Format(m_EndDate, "YYYY/MM/DD") & vbCrLf
    stats = stats & "足種: " & m_TimeFrame & vbCrLf
    stats = stats & "レコード数: " & Format(m_RecordCount, "#,##0") & vbCrLf
    stats = stats & "最高値: " & Format(maxPrice, "#,##0.00") & vbCrLf
    stats = stats & "最安値: " & Format(minPrice, "#,##0.00") & vbCrLf
    stats = stats & "総出来高: " & Format(totalVolume, "#,##0") & vbCrLf
    stats = stats & "最終更新: " & Format(m_LastUpdate, "YYYY/MM/DD HH:MM:SS")
    
    GetStatistics = stats
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "StockData.GetStatistics: " & Err.Description)
    GetStatistics = "統計情報の取得でエラーが発生しました"
End Function

' 指定された日付範囲のデータを抽出
Public Function ExtractDateRange(extractStart As Date, extractEnd As Date) As StockData
    On Error GoTo ErrorHandler
    
    Dim newStockData As StockData
    Dim extractedData() As Variant
    Dim i As Long, j As Long
    Dim dataDate As Date
    Dim extractCount As Long
    
    Set newStockData = New StockData
    
    If Not IsArray(m_DataArray) Or m_RecordCount = 0 Then
        Set ExtractDateRange = newStockData
        Exit Function
    End If
    
    ' 抽出データの件数をカウント
    extractCount = 0
    For i = LBound(m_DataArray, 1) To UBound(m_DataArray, 1)
        dataDate = CDate(m_DataArray(i, 0)) ' DateTime列
        If dataDate >= extractStart And dataDate <= extractEnd Then
            extractCount = extractCount + 1
        End If
    Next i
    
    If extractCount = 0 Then
        Set ExtractDateRange = newStockData
        Exit Function
    End If
    
    ' 抽出データ配列を初期化
    ReDim extractedData(0 To extractCount - 1, 0 To UBound(m_DataArray, 2))
    
    ' データを抽出
    j = 0
    For i = LBound(m_DataArray, 1) To UBound(m_DataArray, 1)
        dataDate = CDate(m_DataArray(i, 0))
        If dataDate >= extractStart And dataDate <= extractEnd Then
            Dim k As Long
            For k = 0 To UBound(m_DataArray, 2)
                extractedData(j, k) = m_DataArray(i, k)
            Next k
            j = j + 1
        End If
    Next i
    
    ' 新しいStockDataオブジェクトを設定
    With newStockData
        .StockCode = Me.GetFullStockCode()
        .StockName = m_StockName
        .TimeFrame = m_TimeFrame
        .StartDate = extractStart
        .EndDate = extractEnd
        .DataArray = extractedData
    End With
    
    Set ExtractDateRange = newStockData
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "StockData.ExtractDateRange: " & Err.Description)
    Set ExtractDateRange = New StockData
End Function

' JSON形式でエクスポート（簡易版）
Public Function ToJSON() As String
    On Error GoTo ErrorHandler
    
    Dim json As String
    Dim i As Long
    
    json = "{" & vbCrLf
    json = json & "  ""stockCode"": """ & GetFullStockCode() & """," & vbCrLf
    json = json & "  ""stockName"": """ & m_StockName & """," & vbCrLf
    json = json & "  ""timeFrame"": """ & m_TimeFrame & """," & vbCrLf
    json = json & "  ""startDate"": """ & Format(m_StartDate, "YYYY-MM-DD") & """," & vbCrLf
    json = json & "  ""endDate"": """ & Format(m_EndDate, "YYYY-MM-DD") & """," & vbCrLf
    json = json & "  ""recordCount"": " & m_RecordCount & "," & vbCrLf
    json = json & "  ""lastUpdate"": """ & Format(m_LastUpdate, "YYYY-MM-DD HH:MM:SS") & """," & vbCrLf
    json = json & "  ""data"": [" & vbCrLf
    
    If IsArray(m_DataArray) And m_RecordCount > 0 Then
        For i = LBound(m_DataArray, 1) To UBound(m_DataArray, 1)
            If i > LBound(m_DataArray, 1) Then json = json & "," & vbCrLf
            
            json = json & "    {" & vbCrLf
            json = json & "      ""dateTime"": """ & Format(CDate(m_DataArray(i, 0)), "YYYY-MM-DD HH:MM:SS") & """," & vbCrLf
            json = json & "      ""open"": " & m_DataArray(i, 1) & "," & vbCrLf
            json = json & "      ""high"": " & m_DataArray(i, 2) & "," & vbCrLf
            json = json & "      ""low"": " & m_DataArray(i, 3) & "," & vbCrLf
            json = json & "      ""close"": " & m_DataArray(i, 4) & "," & vbCrLf
            json = json & "      ""volume"": " & m_DataArray(i, 5) & vbCrLf
            json = json & "    }"
        Next i
    End If
    
    json = json & vbCrLf & "  ]" & vbCrLf
    json = json & "}"
    
    ToJSON = json
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "StockData.ToJSON: " & Err.Description)
    ToJSON = "{""error"": ""JSON生成でエラーが発生しました""}"
End Function

' クラス終了時
Private Sub Class_Terminate()
    Call ClearData()
End Sub