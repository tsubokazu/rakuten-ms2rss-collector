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
    Call LogMessage(LOG_INFO, "アプリケーション開始: " & APP_NAME & " v" & APP_VERSION)
    
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
    
    Call LogMessage(LOG_INFO, "アプリケーション終了")
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
        Call LogMessage(LOG_ERROR, "CSV出力ディレクトリの作成に失敗しました")
        setupOK = False
    End If
    
    If Not EnsureDirectoryExists(ThisWorkbook.Path & "\output\logs\") Then
        Call LogMessage(LOG_ERROR, "ログディレクトリの作成に失敗しました")
        setupOK = False
    End If
    
    ' 2. 設定ファイルの確認
    Dim config As Configuration
    Set config = New Configuration
    
    If Not config.LoadFromFile() Then
        Call LogMessage(LOG_WARN, "設定ファイルの読み込みに失敗、デフォルト設定を使用します")
        config.SaveToFile ' デフォルト設定を保存
    End If
    
    If Not config.ValidateSettings() Then
        Call LogMessage(LOG_ERROR, "設定値に問題があります")
        setupOK = False
    End If
    
    Set config = Nothing
    
    ' 3. MarketSpeed2接続テスト
    If Not TestMS2Connection() Then
        Call LogMessage(LOG_WARN, "MarketSpeed2への接続テストに失敗しました")
        ' 警告のみで続行（オフラインでもVBAコードは確認可能）
    End If
    
    CheckInitialSetup = setupOK
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "CheckInitialSetup: " & Err.Description)
    CheckInitialSetup = False
End Function

' MarketSpeed2接続テスト
Private Function TestMS2Connection() As Boolean
    On Error GoTo ErrorHandler
    
    Dim testResult As Variant
    
    ' 簡易接続テスト（日経平均の現在値を取得）
    testResult = Application.WorksheetFunction.RssIndexMarket("0000", "現在値")
    
    If IsError(testResult) Then
        Call LogMessage(LOG_WARN, "MS2接続テスト失敗: RSS関数がエラーを返しました")
        TestMS2Connection = False
    Else
        Call LogMessage(LOG_INFO, "MS2接続テスト成功: " & testResult)
        TestMS2Connection = True
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_WARN, "MS2接続テスト例外: " & Err.Description)
    TestMS2Connection = False
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
    
    Call LogMessage(LOG_INFO, "クイックテスト開始: " & testStockCode)
    
    ' データ取得テスト
    result = CollectStockData(testStockCode, testTimeFrame, testStartDate, testEndDate)
    
    If result Then
        MsgBox "クイックテスト成功！" & vbCrLf & _
               "銘柄: " & testStockCode & vbCrLf & _
               "足種: " & testTimeFrame & vbCrLf & _
               "期間: " & Format(testStartDate, "MM/DD") & " - " & Format(testEndDate, "MM/DD"), _
               vbInformation, "テスト結果"
        Call LogMessage(LOG_INFO, "クイックテスト成功")
    Else
        MsgBox "クイックテスト失敗。ログを確認してください。", vbExclamation, "テスト結果"
        Call LogMessage(LOG_ERROR, "クイックテスト失敗")
    End If
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("QuickTest", Err.Description, "株価: " & testStockCode)
    MsgBox "テスト実行中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' 現在の設定を表示
Public Sub ShowCurrentConfig()
    On Error GoTo ErrorHandler
    
    Dim config As Configuration
    Set config = New Configuration
    
    If config.LoadFromFile() Then
        MsgBox config.ToString(), vbInformation, "現在の設定"
    Else
        MsgBox "設定ファイルの読み込みに失敗しました。", vbExclamation, "設定エラー"
    End If
    
    Set config = Nothing
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "ShowCurrentConfig: " & Err.Description)
    MsgBox "設定表示でエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ログビューアーを表示
Public Sub ShowLogViewer()
    On Error GoTo ErrorHandler
    
    Dim logFilePath As String
    Dim logContent As String
    Dim fileNum As Integer
    
    ' 今日のログファイルパス
    logFilePath = ThisWorkbook.Path & "\output\logs\ms2rss_collector_" & Format(Date, "YYYYMMDD") & ".log"
    
    ' ログファイル存在チェック
    If Dir(logFilePath) = "" Then
        MsgBox "本日のログファイルが見つかりません。" & vbCrLf & logFilePath, vbInformation, "ログビューアー"
        Exit Sub
    End If
    
    ' ログファイル読み込み
    fileNum = FreeFile
    Open logFilePath For Input As #fileNum
    logContent = Input(LOF(fileNum), fileNum)
    Close #fileNum
    
    ' ログ内容をメッセージボックスで表示（簡易版）
    If Len(logContent) > 1000 Then
        logContent = "... (省略) ..." & vbCrLf & Right(logContent, 800)
    End If
    
    MsgBox "【本日のログ】" & vbCrLf & vbCrLf & logContent, vbInformation, "ログビューアー"
    
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage(LOG_ERROR, "ShowLogViewer: " & Err.Description)
    MsgBox "ログ表示でエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' バージョン情報を表示
Public Sub ShowAbout()
    Dim aboutMessage As String
    
    aboutMessage = APP_NAME & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "バージョン: " & APP_VERSION & vbCrLf
    aboutMessage = aboutMessage & "ビルド日: " & BUILD_DATE & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "楽天証券MarketSpeed2のRSS APIを使用して" & vbCrLf
    aboutMessage = aboutMessage & "株価データを取得し、CSV形式で出力します。" & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "Created with Claude Code" & vbCrLf
    aboutMessage = aboutMessage & "https://github.com/tsubokazu/rakuten-ms2rss-collector"
    
    MsgBox aboutMessage, vbInformation, "このアプリケーションについて"
End Sub

' アプリケーション終了時のクリーンアップ
Public Sub CleanupApplication()
    On Error Resume Next
    
    ' 進捗表示をクリア
    Call ClearProgress()
    
    ' アプリケーション設定を復元
    Call RestoreApplicationSettings()
    
    ' ログメッセージ
    Call LogMessage(LOG_INFO, "アプリケーションクリーンアップ完了")
End Sub

' 自動実行用サンプル
Public Sub AutoCollectSample()
    On Error GoTo ErrorHandler
    
    Dim stocks As String
    Dim timeFrame As String
    Dim startDate As Date
    Dim endDate As Date
    Dim result As Boolean
    
    ' サンプル設定
    stocks = "7203,6758,9984"    ' トヨタ、ソニー、ソフトバンク
    timeFrame = "5M"             ' 5分足
    startDate = Date - 7         ' 1週間前
    endDate = Date               ' 今日
    
    ' 確認メッセージ
    If MsgBox("以下の設定で自動データ取得を実行しますか？" & vbCrLf & vbCrLf & _
              "銘柄: " & stocks & vbCrLf & _
              "足種: " & timeFrame & vbCrLf & _
              "期間: " & Format(startDate, "YYYY/MM/DD") & " - " & Format(endDate, "YYYY/MM/DD"), _
              vbYesNo + vbQuestion, "自動実行確認") = vbNo Then
        Exit Sub
    End If
    
    Call LogMessage(LOG_INFO, "自動データ取得開始")
    
    ' バッチ処理実行
    result = CollectMultipleStocks(stocks, timeFrame, startDate, endDate)
    
    If result Then
        MsgBox "自動データ取得が完了しました。", vbInformation, "完了"
        Call LogMessage(LOG_INFO, "自動データ取得完了")
    Else
        MsgBox "自動データ取得で一部エラーが発生しました。ログを確認してください。", vbExclamation, "一部エラー"
        Call LogMessage(LOG_WARN, "自動データ取得で一部エラー")
    End If
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("AutoCollectSample", Err.Description)
    MsgBox "自動実行でエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' Excelブック終了時の処理
Public Sub Workbook_BeforeClose()
    Call CleanupApplication()
End Sub