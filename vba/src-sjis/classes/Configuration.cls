VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
' 楽天MS2RSS株価データコレクター - 設定管理クラス
' 
' 説明: アプリケーション設定の管理
' 作成者: Claude Code
' バージョン: 1.0.0
'******************************************************************************

Option Explicit

' 設定項目
Private m_MaxBarsPerRequest As Long
Private m_DefaultTimeFrame As String
Private m_DefaultOutputPath As String
Private m_LogEnabled As Boolean
Private m_LogLevel As String
Private m_MaxLogFiles As Integer
Private m_MaxFileSizeMB As Integer
Private m_DateTimeFormat As String
Private m_DecimalPlaces As Integer
Private m_CSVEncoding As String
Private m_MaxStocksBatch As Integer

' 設定ファイルパス
Private m_ConfigFilePath As String

' プロパティ - API制限
Public Property Get MaxBarsPerRequest() As Long
    MaxBarsPerRequest = m_MaxBarsPerRequest
End Property

Public Property Let MaxBarsPerRequest(value As Long)
    If value > 0 And value <= 3000 Then
        m_MaxBarsPerRequest = value
    Else
        m_MaxBarsPerRequest = 3000
    End If
End Property

' プロパティ - デフォルト足種
Public Property Get DefaultTimeFrame() As String
    DefaultTimeFrame = m_DefaultTimeFrame
End Property

Public Property Let DefaultTimeFrame(value As String)
    If ValidateTimeFrame(value) Then
        m_DefaultTimeFrame = UCase(value)
    Else
        m_DefaultTimeFrame = "5M"
    End If
End Property

' プロパティ - デフォルト出力パス
Public Property Get DefaultOutputPath() As String
    DefaultOutputPath = m_DefaultOutputPath
End Property

Public Property Let DefaultOutputPath(value As String)
    m_DefaultOutputPath = value
End Property

' プロパティ - ログ有効/無効
Public Property Get LogEnabled() As Boolean
    LogEnabled = m_LogEnabled
End Property

Public Property Let LogEnabled(value As Boolean)
    m_LogEnabled = value
End Property

' プロパティ - ログレベル
Public Property Get LogLevel() As String
    LogLevel = m_LogLevel
End Property

Public Property Let LogLevel(value As String)
    Select Case UCase(value)
        Case "DEBUG", "INFO", "WARN", "ERROR"
            m_LogLevel = UCase(value)
        Case Else
            m_LogLevel = "INFO"
    End Select
End Property

' プロパティ - 最大ログファイル数
Public Property Get MaxLogFiles() As Integer
    MaxLogFiles = m_MaxLogFiles
End Property

Public Property Let MaxLogFiles(value As Integer)
    If value > 0 And value <= 100 Then
        m_MaxLogFiles = value
    Else
        m_MaxLogFiles = 10
    End If
End Property

' プロパティ - ログファイル最大サイズ
Public Property Get MaxFileSizeMB() As Integer
    MaxFileSizeMB = m_MaxFileSizeMB
End Property

Public Property Let MaxFileSizeMB(value As Integer)
    If value > 0 And value <= 1000 Then
        m_MaxFileSizeMB = value
    Else
        m_MaxFileSizeMB = 10
    End If
End Property

' プロパティ - 日時フォーマット
Public Property Get DateTimeFormat() As String
    DateTimeFormat = m_DateTimeFormat
End Property

Public Property Let DateTimeFormat(value As String)
    m_DateTimeFormat = value
End Property

' プロパティ - 小数点以下桁数
Public Property Get DecimalPlaces() As Integer
    DecimalPlaces = m_DecimalPlaces
End Property

Public Property Let DecimalPlaces(value As Integer)
    If value >= 0 And value <= 8 Then
        m_DecimalPlaces = value
    Else
        m_DecimalPlaces = 2
    End If
End Property

' プロパティ - CSV文字エンコーディング
Public Property Get CSVEncoding() As String
    CSVEncoding = m_CSVEncoding
End Property

Public Property Let CSVEncoding(value As String)
    m_CSVEncoding = value
End Property

' プロパティ - バッチ処理最大銘柄数
Public Property Get MaxStocksBatch() As Integer
    MaxStocksBatch = m_MaxStocksBatch
End Property

Public Property Let MaxStocksBatch(value As Integer)
    If value > 0 And value <= 100 Then
        m_MaxStocksBatch = value
    Else
        m_MaxStocksBatch = 20
    End If
End Property

' クラス初期化
Private Sub Class_Initialize()
    Call LoadDefaultSettings()
    m_ConfigFilePath = ThisWorkbook.Path & "\config\settings.json"
End Sub

' デフォルト設定の読み込み
Private Sub LoadDefaultSettings()
    m_MaxBarsPerRequest = 3000
    m_DefaultTimeFrame = "5M"
    m_DefaultOutputPath = ThisWorkbook.Path & "\output\csv\"
    m_LogEnabled = True
    m_LogLevel = "INFO"
    m_MaxLogFiles = 10
    m_MaxFileSizeMB = 10
    m_DateTimeFormat = "YYYY-MM-DD HH:MM:SS"
    m_DecimalPlaces = 2
    m_CSVEncoding = "UTF-8"
    m_MaxStocksBatch = 20
End Sub

' 設定ファイルから読み込み
Public Function LoadFromFile(Optional filePath As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim configPath As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim key As String
    Dim value As String
    
    If filePath = "" Then
        configPath = m_ConfigFilePath
    Else
        configPath = filePath
    End If
    
    ' ファイル存在チェック
    If Dir(configPath) = "" Then
        Call LogMessage(LOG_WARN, "設定ファイルが見つかりません: " & configPath)
        LoadFromFile = False
        Exit Function
    End If
    
    ' ファイル読み込み（簡易JSON解析）
    fileNum = FreeFile
    Open configPath For Input As #fileNum
    fileContent = Input(LOF(fileNum), fileNum)
    Close #fileNum
    
    ' 基本的なJSON解析（キー:値のペアを抽出）
    Call ParseJSONSettings(fileContent)
    
    Call LogMessage(LOG_INFO, "設定ファイルを読み込みました: " & configPath)
    LoadFromFile = True
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage(LOG_ERROR, "LoadFromFile: " & Err.Description)
    LoadFromFile = False
End Function

' 簡易JSON解析
Private Sub ParseJSONSettings(jsonContent As String)
    On Error Resume Next
    
    ' max_bars_per_request
    m_MaxBarsPerRequest = ExtractNumberFromJSON(jsonContent, "max_bars_per_request", 3000)
    
    ' default_timeframe
    m_DefaultTimeFrame = ExtractStringFromJSON(jsonContent, "default_timeframe", "5M")
    
    ' decimal_places
    m_DecimalPlaces = ExtractNumberFromJSON(jsonContent, "decimal_places", 2)
    
    ' max_stocks_batch
    m_MaxStocksBatch = ExtractNumberFromJSON(jsonContent, "max_stocks_batch", 20)
    
    ' logging enabled
    m_LogEnabled = ExtractBooleanFromJSON(jsonContent, "enabled", True)
    
    ' log level
    m_LogLevel = ExtractStringFromJSON(jsonContent, "level", "INFO")
    
    ' max_log_files
    m_MaxLogFiles = ExtractNumberFromJSON(jsonContent, "max_log_files", 10)
    
    ' max_file_size_mb
    m_MaxFileSizeMB = ExtractNumberFromJSON(jsonContent, "max_file_size_mb", 10)
End Sub

' JSONから文字列値を抽出
Private Function ExtractStringFromJSON(jsonContent As String, key As String, defaultValue As String) As String
    Dim pattern As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String
    
    pattern = """" & key & """:"
    startPos = InStr(jsonContent, pattern)
    
    If startPos > 0 Then
        startPos = InStr(startPos, jsonContent, """") + 1
        endPos = InStr(startPos, jsonContent, """")
        If endPos > startPos Then
            value = Mid(jsonContent, startPos, endPos - startPos)
            ExtractStringFromJSON = value
            Exit Function
        End If
    End If
    
    ExtractStringFromJSON = defaultValue
End Function

' JSONから数値を抽出
Private Function ExtractNumberFromJSON(jsonContent As String, key As String, defaultValue As Long) As Long
    Dim pattern As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String
    
    pattern = """" & key & """:"
    startPos = InStr(jsonContent, pattern)
    
    If startPos > 0 Then
        startPos = InStr(startPos, jsonContent, ":") + 1
        
        ' 空白をスキップ
        While Mid(jsonContent, startPos, 1) = " "
            startPos = startPos + 1
        Wend
        
        endPos = startPos
        While IsNumeric(Mid(jsonContent, endPos, 1)) Or Mid(jsonContent, endPos, 1) = "."
            endPos = endPos + 1
        Wend
        
        If endPos > startPos Then
            value = Mid(jsonContent, startPos, endPos - startPos)
            If IsNumeric(value) Then
                ExtractNumberFromJSON = CLng(value)
                Exit Function
            End If
        End If
    End If
    
    ExtractNumberFromJSON = defaultValue
End Function

' JSONからブール値を抽出
Private Function ExtractBooleanFromJSON(jsonContent As String, key As String, defaultValue As Boolean) As Boolean
    Dim pattern As String
    Dim startPos As Long
    Dim value As String
    
    pattern = """" & key & """:"
    startPos = InStr(jsonContent, pattern)
    
    If startPos > 0 Then
        startPos = InStr(startPos, jsonContent, ":") + 1
        
        ' 空白をスキップ
        While Mid(jsonContent, startPos, 1) = " "
            startPos = startPos + 1
        Wend
        
        If Mid(jsonContent, startPos, 4) = "true" Then
            ExtractBooleanFromJSON = True
            Exit Function
        ElseIf Mid(jsonContent, startPos, 5) = "false" Then
            ExtractBooleanFromJSON = False
            Exit Function
        End If
    End If
    
    ExtractBooleanFromJSON = defaultValue
End Function

' 設定ファイルに保存
Public Function SaveToFile(Optional filePath As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim configPath As String
    Dim fileNum As Integer
    Dim jsonContent As String
    
    If filePath = "" Then
        configPath = m_ConfigFilePath
    Else
        configPath = filePath
    End If
    
    ' ディレクトリ作成
    Call EnsureDirectoryExists(Left(configPath, InStrRev(configPath, "\")))
    
    ' JSON形式で設定内容を生成
    jsonContent = GenerateJSONSettings()
    
    ' ファイル保存
    fileNum = FreeFile
    Open configPath For Output As #fileNum
    Print #fileNum, jsonContent
    Close #fileNum
    
    Call LogMessage(LOG_INFO, "設定ファイルを保存しました: " & configPath)
    SaveToFile = True
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage(LOG_ERROR, "SaveToFile: " & Err.Description)
    SaveToFile = False
End Function

' JSON形式の設定内容を生成
Private Function GenerateJSONSettings() As String
    Dim json As String
    
    json = "{" & vbCrLf
    json = json & "  ""application"": {" & vbCrLf
    json = json & "    ""name"": ""楽天MS2RSS株価データコレクター""," & vbCrLf
    json = json & "    ""version"": ""1.0.0""," & vbCrLf
    json = json & "    ""author"": ""Claude Code""" & vbCrLf
    json = json & "  }," & vbCrLf
    json = json & "  ""data_collection"": {" & vbCrLf
    json = json & "    ""max_bars_per_request"": " & m_MaxBarsPerRequest & "," & vbCrLf
    json = json & "    ""default_timeframe"": """ & m_DefaultTimeFrame & """," & vbCrLf
    json = json & "    ""default_market"": ""T""" & vbCrLf
    json = json & "  }," & vbCrLf
    json = json & "  ""csv_output"": {" & vbCrLf
    json = json & "    ""datetime_format"": """ & m_DateTimeFormat & """," & vbCrLf
    json = json & "    ""decimal_places"": " & m_DecimalPlaces & "," & vbCrLf
    json = json & "    ""include_header"": true," & vbCrLf
    json = json & "    ""encoding"": """ & m_CSVEncoding & """" & vbCrLf
    json = json & "  }," & vbCrLf
    json = json & "  ""ui_settings"": {" & vbCrLf
    json = json & "    ""max_stocks_batch"": " & m_MaxStocksBatch & vbCrLf
    json = json & "  }," & vbCrLf
    json = json & "  ""logging"": {" & vbCrLf
    json = json & "    ""enabled"": " & LCase(CStr(m_LogEnabled)) & "," & vbCrLf
    json = json & "    ""level"": """ & m_LogLevel & """," & vbCrLf
    json = json & "    ""max_log_files"": " & m_MaxLogFiles & "," & vbCrLf
    json = json & "    ""max_file_size_mb"": " & m_MaxFileSizeMB & vbCrLf
    json = json & "  }" & vbCrLf
    json = json & "}"
    
    GenerateJSONSettings = json
End Function

' 設定の妥当性チェック
Public Function ValidateSettings() As Boolean
    On Error GoTo ErrorHandler
    
    Dim isValid As Boolean
    isValid = True
    
    ' 数値範囲チェック
    If m_MaxBarsPerRequest <= 0 Or m_MaxBarsPerRequest > 3000 Then
        Call LogMessage(LOG_ERROR, "無効なMaxBarsPerRequest: " & m_MaxBarsPerRequest)
        isValid = False
    End If
    
    If m_DecimalPlaces < 0 Or m_DecimalPlaces > 8 Then
        Call LogMessage(LOG_ERROR, "無効なDecimalPlaces: " & m_DecimalPlaces)
        isValid = False
    End If
    
    If m_MaxStocksBatch <= 0 Or m_MaxStocksBatch > 100 Then
        Call LogMessage(LOG_ERROR, "無効なMaxStocksBatch: " & m_MaxStocksBatch)
        isValid = False
    End If
    
    ' 足種チェック
    If Not ValidateTimeFrame(m_DefaultTimeFrame) Then
        Call LogMessage(LOG_ERROR, "無効なDefaultTimeFrame: " & m_DefaultTimeFrame)
        isValid = False
    End If
    
    ' パスチェック
    If m_DefaultOutputPath = "" Then
        Call LogMessage(LOG_ERROR, "DefaultOutputPathが設定されていません")
        isValid = False
    End If
    
    ValidateSettings = isValid
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "ValidateSettings: " & Err.Description)
    ValidateSettings = False
End Function

' 設定のリセット
Public Sub ResetToDefaults()
    Call LoadDefaultSettings()
    Call LogMessage(LOG_INFO, "設定をデフォルトにリセットしました")
End Sub

' 設定の表示用文字列を生成
Public Function ToString() As String
    Dim result As String
    
    result = "【現在の設定】" & vbCrLf
    result = result & "API最大取得本数: " & m_MaxBarsPerRequest & vbCrLf
    result = result & "デフォルト足種: " & m_DefaultTimeFrame & vbCrLf
    result = result & "デフォルト出力パス: " & m_DefaultOutputPath & vbCrLf
    result = result & "ログ有効: " & m_LogEnabled & vbCrLf
    result = result & "ログレベル: " & m_LogLevel & vbCrLf
    result = result & "CSV小数点桁数: " & m_DecimalPlaces & vbCrLf
    result = result & "バッチ最大銘柄数: " & m_MaxStocksBatch
    
    ToString = result
End Function