Attribute VB_Name = "MainModule"
'******************************************************************************
' 楽天MS2RSS株価データコレクター - メインモジュール
' 
' 説明: アプリケーションのエントリーポイントとメイン制御
' 作成者: Claude Code
' バージョン: 1.0.0
'******************************************************************************

Option Explicit

' アプリケーション情報
Public Const APP_NAME As String = "楽天MS2RSS株価データコレクター"
Public Const APP_VERSION As String = "1.0.0"
Public Const BUILD_DATE As String = "2025-01-16"

' メインフォームを表示
Public Sub ShowMainForm()
    On Error GoTo ErrorHandler
    
    ' ログ初期化
    Call LogMessage("INFO", "アプリケーション開始: " & APP_NAME & " v" & APP_VERSION)
    
    ' 初期設定チェック
    If Not CheckInitialSetup() Then
        MsgBox "初期設定に問題があります。ログを確認してください。", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' メインフォーム表示
    Load MainForm
    MainForm.Show vbModal
    
    ' フォームが閉じられた後のクリーンアップ
    Unload MainForm
    Set MainForm = Nothing
    
    Call LogMessage("INFO", "アプリケーション終了")
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("ShowMainForm", Err.Description)
    MsgBox "アプリケーションの起動でエラーが発生しました: " & Err.Description, vbCritical, APP_NAME
End Sub

' 初期設定チェック
Private Function CheckInitialSetup() As Boolean
    On Error GoTo ErrorHandler
    
    Dim setupOK As Boolean
    setupOK = True
    
    ' 1. 出力ディレクトリの存在確認・作成
    If Not EnsureDirectoryExists(ThisWorkbook.Path & "\output\csv\") Then
        Call LogMessage("ERROR", "CSV出力ディレクトリの作成に失敗しました")
        setupOK = False
    End If
    
    If Not EnsureDirectoryExists(ThisWorkbook.Path & "\output\logs\") Then
        Call LogMessage("ERROR", "ログディレクトリの作成に失敗しました")
        setupOK = False
    End If
    
    CheckInitialSetup = setupOK
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "CheckInitialSetup: " & Err.Description)
    CheckInitialSetup = False
End Function

' 簡単テスト実行
Public Sub QuickTest()
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim testStockCode As String
    Dim testTimeFrame As String
    Dim testStartDate As Date
    Dim testEndDate As Date
    
    ' テストパラメータ
    testStockCode = "7203"  ' トヨタ自動車
    testTimeFrame = "5M"    ' 5分足
    testStartDate = Date - 1  ' 昨日
    testEndDate = Date        ' 今日
    
    Call LogMessage("INFO", "クイックテスト開始: " & testStockCode)
    
    ' データ取得テスト
    result = CollectStockData(testStockCode, testTimeFrame, testStartDate, testEndDate)
    
    If result Then
        MsgBox "クイックテスト成功！" & vbCrLf & _
               "銘柄: " & testStockCode & vbCrLf & _
               "足種: " & testTimeFrame & vbCrLf & _
               "期間: " & Format(testStartDate, "MM/DD") & " - " & Format(testEndDate, "MM/DD"), _
               vbInformation, "テスト結果"
        Call LogMessage("INFO", "クイックテスト成功")
    Else
        MsgBox "クイックテスト失敗。ログを確認してください。", vbExclamation, "テスト結果"
        Call LogMessage("ERROR", "クイックテスト失敗")
    End If
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("QuickTest", Err.Description, "株価: " & testStockCode)
    MsgBox "テスト実行中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' バージョン情報を表示
Public Sub ShowAbout()
    Dim aboutMessage As String
    
    aboutMessage = APP_NAME & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "バージョン: " & APP_VERSION & vbCrLf
    aboutMessage = aboutMessage & "ビルド日: " & BUILD_DATE & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "楽天証券MarketSpeed2のRSS APIを使用して" & vbCrLf
    aboutMessage = aboutMessage & "株価データを取得し、CSV形式で出力します。" & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "Created with Claude Code"
    
    MsgBox aboutMessage, vbInformation, "このアプリケーションについて"
End Sub