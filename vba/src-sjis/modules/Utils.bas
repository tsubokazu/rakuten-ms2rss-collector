Attribute VB_Name = "Utils"
'******************************************************************************
' 楽天MS2RSS株価データコレクター - ユーティリティモジュール
' 
' 説明: 共通ユーティリティ関数・ログ機能
' 作成者: Claude Code
' バージョン: 1.0.0
'******************************************************************************

Option Explicit

' ログレベル定数
Public Const LOG_DEBUG As String = "DEBUG"
Public Const LOG_INFO As String = "INFO"
Public Const LOG_WARN As String = "WARN"
Public Const LOG_ERROR As String = "ERROR"

' ログ設定
Private Type LogConfig
    Enabled As Boolean
    Level As String
    MaxFiles As Integer
    MaxFileSizeMB As Integer
    LogDirectory As String
End Type

Private logConfig As LogConfig
Private logInitialized As Boolean

' ログ機能初期化
Private Sub InitializeLogging()
    If logInitialized Then Exit Sub
    
    logConfig.Enabled = True
    logConfig.Level = LOG_INFO
    logConfig.MaxFiles = 10
    logConfig.MaxFileSizeMB = 10
    logConfig.LogDirectory = ThisWorkbook.Path & "\output\logs\"
    
    ' ログディレクトリ作成
    If Dir(logConfig.LogDirectory, vbDirectory) = "" Then
        MkDir logConfig.LogDirectory
    End If
    
    logInitialized = True
End Sub

' ログメッセージ出力
Public Sub LogMessage(level As String, message As String)
    On Error GoTo ErrorHandler
    
    Call InitializeLogging
    
    If Not logConfig.Enabled Then Exit Sub
    
    Dim logFilePath As String
    Dim fileNum As Integer
    Dim timestamp As String
    Dim logLine As String
    
    ' タイムスタンプ生成
    timestamp = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    ' ログ行作成
    logLine = timestamp & " [" & level & "] " & message
    
    ' コンソール出力（イミディエイトウィンドウ）
    Debug.Print logLine
    
    ' ファイル出力
    logFilePath = GetCurrentLogFilePath()
    
    ' ファイルサイズチェック
    If Dir(logFilePath) <> "" And FileLen(logFilePath) > logConfig.MaxFileSizeMB * 1024 * 1024 Then
        Call RotateLogFiles()
        logFilePath = GetCurrentLogFilePath()
    End If
    
    ' ログファイルに書き込み
    fileNum = FreeFile
    Open logFilePath For Append As #fileNum
    Print #fileNum, logLine
    Close #fileNum
    
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Debug.Print "LogMessage Error: " & Err.Description
End Sub

' 現在のログファイルパスを取得
Private Function GetCurrentLogFilePath() As String
    Dim fileName As String
    fileName = "ms2rss_collector_" & Format(Date, "YYYYMMDD") & ".log"
    GetCurrentLogFilePath = logConfig.LogDirectory & fileName
End Function

' ログファイルローテーション
Private Sub RotateLogFiles()
    On Error Resume Next
    
    Dim i As Integer
    Dim oldFile As String
    Dim newFile As String
    Dim baseFileName As String
    
    baseFileName = "ms2rss_collector_" & Format(Date, "YYYYMMDD")
    
    ' 古いファイルを削除
    For i = logConfig.MaxFiles To 1 Step -1
        oldFile = logConfig.LogDirectory & baseFileName & "_" & Format(i, "00") & ".log"
        If Dir(oldFile) <> "" Then
            If i = logConfig.MaxFiles Then
                Kill oldFile
            Else
                newFile = logConfig.LogDirectory & baseFileName & "_" & Format(i + 1, "00") & ".log"
                Name oldFile As newFile
            End If
        End If
    Next i
    
    ' 現在のファイルを_01にリネーム
    oldFile = GetCurrentLogFilePath()
    If Dir(oldFile) <> "" Then
        newFile = logConfig.LogDirectory & baseFileName & "_01.log"
        Name oldFile As newFile
    End If
End Sub

' 現在時刻の文字列取得
Public Function GetCurrentTimestamp() As String
    GetCurrentTimestamp = Format(Now, "YYYY-MM-DD HH:MM:SS")
End Function

' ディレクトリ存在チェック・作成
Public Function EnsureDirectoryExists(dirPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Dir(dirPath, vbDirectory) = "" Then
        MkDir dirPath
        Call LogMessage(LOG_INFO, "ディレクトリを作成しました: " & dirPath)
    End If
    
    EnsureDirectoryExists = True
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "ディレクトリ作成エラー: " & dirPath & " - " & Err.Description)
    EnsureDirectoryExists = False
End Function

' ファイル存在チェック
Public Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

' 安全なファイル名生成（無効文字を除去）
Public Function SafeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim char As String
    Dim result As String
    
    invalidChars = "\/:*?""<>|"
    result = fileName
    
    For i = 1 To Len(invalidChars)
        char = Mid(invalidChars, i, 1)
        result = Replace(result, char, "_")
    Next i
    
    SafeFileName = result
End Function

' 進捗表示（0-100%）
Public Sub ShowProgress(currentValue As Long, maxValue As Long, Optional message As String = "")
    Dim percentage As Integer
    Dim progressBar As String
    Dim i As Integer
    
    If maxValue = 0 Then Exit Sub
    
    percentage = Int((currentValue / maxValue) * 100)
    
    ' プログレスバー文字列作成
    progressBar = "["
    For i = 1 To 50
        If i <= Int(percentage / 2) Then
            progressBar = progressBar & "="
        Else
            progressBar = progressBar & " "
        End If
    Next i
    progressBar = progressBar & "] " & percentage & "%"
    
    If message <> "" Then
        progressBar = message & " " & progressBar
    End If
    
    ' ステータスバーに表示
    Application.StatusBar = progressBar
    
    ' ログにも出力（10%刻み）
    If percentage Mod 10 = 0 And currentValue > 0 Then
        Call LogMessage(LOG_INFO, "進捗: " & percentage & "% (" & currentValue & "/" & maxValue & ")")
    End If
    
    DoEvents
End Sub

' 進捗表示クリア
Public Sub ClearProgress()
    Application.StatusBar = False
End Sub

' 設定値の検証
Public Function ValidateTimeFrame(timeFrame As String) As Boolean
    Dim validFrames As Variant
    Dim i As Integer
    
    validFrames = Array("T", "1M", "2M", "3M", "4M", "5M", "10M", "15M", "30M", "60M", _
                       "2H", "4H", "8H", "D", "W", "M")
    
    For i = 0 To UBound(validFrames)
        If UCase(timeFrame) = UCase(validFrames(i)) Then
            ValidateTimeFrame = True
            Exit Function
        End If
    Next i
    
    ValidateTimeFrame = False
End Function

' 日付範囲の検証
Public Function ValidateDateRange(startDate As Date, endDate As Date) As Boolean
    If startDate > endDate Then
        Call LogMessage(LOG_ERROR, "開始日が終了日より後です")
        ValidateDateRange = False
        Exit Function
    End If
    
    If startDate > Date Then
        Call LogMessage(LOG_WARN, "開始日が未来日です")
    End If
    
    If DateDiff("d", startDate, endDate) > 365 Then
        Call LogMessage(LOG_WARN, "取得期間が1年を超えています")
    End If
    
    ValidateDateRange = True
End Function

' メモリ使用量チェック（簡易版）
Public Function GetMemoryUsage() As String
    On Error Resume Next
    
    Dim memInfo As String
    memInfo = "Excel: " & Format(Application.MemoryUsed / 1024, "0.0") & "MB"
    
    GetMemoryUsage = memInfo
End Function

' アプリケーション設定の復元
Public Sub RestoreApplicationSettings()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub

' アプリケーション設定の最適化
Public Sub OptimizeApplicationSettings()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

' バージョン情報取得
Public Function GetApplicationVersion() As String
    GetApplicationVersion = "楽天MS2RSS株価データコレクター v1.0.0"
End Function

' システム情報取得
Public Function GetSystemInfo() As String
    Dim info As String
    
    info = "Excel: " & Application.Version & vbCrLf
    info = info & "OS: " & Application.OperatingSystem & vbCrLf
    info = info & "ユーザー: " & Application.UserName & vbCrLf
    info = info & "実行日時: " & GetCurrentTimestamp()
    
    GetSystemInfo = info
End Function

' エラー情報の詳細ログ
Public Sub LogDetailedError(functionName As String, errorDescription As String, _
                          Optional additionalInfo As String = "")
    
    Dim errorMessage As String
    
    errorMessage = "関数: " & functionName & vbCrLf
    errorMessage = errorMessage & "エラー: " & errorDescription & vbCrLf
    
    If additionalInfo <> "" Then
        errorMessage = errorMessage & "追加情報: " & additionalInfo & vbCrLf
    End If
    
    errorMessage = errorMessage & "時刻: " & GetCurrentTimestamp() & vbCrLf
    errorMessage = errorMessage & GetMemoryUsage()
    
    Call LogMessage(LOG_ERROR, errorMessage)
End Sub