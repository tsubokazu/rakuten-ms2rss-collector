Attribute VB_Name = "Utils"
'******************************************************************************
' 楽天MS2RSS株価データコレクター - ユーティリティモジュール（簡易版）
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

' ログメッセージ出力（簡易版）
Public Sub LogMessage(level As String, message As String)
    On Error Resume Next
    
    Dim logLine As String
    Dim timestamp As String
    
    ' タイムスタンプ生成
    timestamp = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    ' ログ行作成
    logLine = timestamp & " [" & level & "] " & message
    
    ' コンソール出力（イミディエイトウィンドウ）
    Debug.Print logLine
End Sub

' ディレクトリ存在チェック・作成（簡易版）
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

' エラー情報の詳細ログ（簡易版）
Public Sub LogDetailedError(functionName As String, errorDescription As String, _
                          Optional additionalInfo As String = "")
    
    Dim errorMessage As String
    
    errorMessage = "関数: " & functionName & " / エラー: " & errorDescription
    
    If additionalInfo <> "" Then
        errorMessage = errorMessage & " / 追加情報: " & additionalInfo
    End If
    
    Call LogMessage(LOG_ERROR, errorMessage)
End Sub