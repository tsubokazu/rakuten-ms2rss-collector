Attribute VB_Name = "WorksheetMacros"
'******************************************************************************
' 楽天MS2RSS株価データコレクター - ワークシートマクロ
' 
' 説明: Excelワークシート上のボタンから呼び出されるマクロ
' 作成者: Claude Code
' バージョン: 1.0.0
'******************************************************************************

Option Explicit

' メインフォーム表示ボタン用マクロ
Public Sub StartDataCollection()
    Call ShowMainForm
End Sub

' クイックテストボタン用マクロ
Public Sub RunQuickTest()
    Call QuickTest
End Sub

' 設定表示ボタン用マクロ
Public Sub DisplaySettings()
    Call ShowCurrentConfig
End Sub

' ログ表示ボタン用マクロ
Public Sub ViewLogs()
    Call ShowLogViewer
End Sub

' バージョン情報ボタン用マクロ
Public Sub AboutApp()
    Call ShowAbout
End Sub

' 自動実行サンプルボタン用マクロ
Public Sub RunAutoSample()
    Call AutoCollectSample
End Sub

' 出力フォルダを開く
Public Sub OpenOutputFolder()
    On Error GoTo ErrorHandler
    
    Dim outputPath As String
    outputPath = ThisWorkbook.Path & "\output\csv\"
    
    ' フォルダが存在しない場合は作成
    If Not EnsureDirectoryExists(outputPath) Then
        MsgBox "出力フォルダの作成に失敗しました: " & outputPath, vbCritical
        Exit Sub
    End If
    
    ' フォルダを開く
    Shell "explorer.exe " & Chr(34) & outputPath & Chr(34), vbNormalFocus
    Call LogMessage(LOG_INFO, "出力フォルダを開きました: " & outputPath)
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "OpenOutputFolder: " & Err.Description)
    MsgBox "フォルダを開くことができませんでした: " & Err.Description, vbCritical
End Sub

' ログフォルダを開く
Public Sub OpenLogFolder()
    On Error GoTo ErrorHandler
    
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\output\logs\"
    
    ' フォルダが存在しない場合は作成
    If Not EnsureDirectoryExists(logPath) Then
        MsgBox "ログフォルダの作成に失敗しました: " & logPath, vbCritical
        Exit Sub
    End If
    
    ' フォルダを開く
    Shell "explorer.exe " & Chr(34) & logPath & Chr(34), vbNormalFocus
    Call LogMessage(LOG_INFO, "ログフォルダを開きました: " & logPath)
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "OpenLogFolder: " & Err.Description)
    MsgBox "フォルダを開くことができませんでした: " & Err.Description, vbCritical
End Sub

' MarketSpeed2接続テスト
Public Sub TestConnection()
    On Error GoTo ErrorHandler
    
    Dim testResult As Variant
    
    ' 日経平均現在値で接続テスト
    testResult = Application.WorksheetFunction.RssIndexMarket("0000", "現在値")
    
    If IsError(testResult) Then
        MsgBox "MarketSpeed2への接続に失敗しました。" & vbCrLf & vbCrLf & _
               "以下を確認してください：" & vbCrLf & _
               "1. MarketSpeed2が起動している" & vbCrLf & _
               "2. RSS機能が有効になっている" & vbCrLf & _
               "3. ログイン状態が正常である", vbExclamation, "接続テスト結果"
        Call LogMessage(LOG_ERROR, "MS2接続テスト失敗")
    Else
        MsgBox "MarketSpeed2への接続に成功しました！" & vbCrLf & vbCrLf & _
               "日経平均現在値: " & testResult, vbInformation, "接続テスト結果"
        Call LogMessage(LOG_INFO, "MS2接続テスト成功: " & testResult)
    End If
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "TestConnection: " & Err.Description)
    MsgBox "接続テストでエラーが発生しました: " & Err.Description, vbCritical, "接続テストエラー"
End Sub

' 設定ファイルの初期化
Public Sub InitializeSettings()
    On Error GoTo ErrorHandler
    
    Dim config As Configuration
    Set config = New Configuration
    
    ' デフォルト設定でファイル作成
    If config.SaveToFile() Then
        MsgBox "設定ファイルを初期化しました。" & vbCrLf & _
               "ファイル: " & ThisWorkbook.Path & "\config\settings.json", vbInformation, "設定初期化"
        Call LogMessage(LOG_INFO, "設定ファイル初期化完了")
    Else
        MsgBox "設定ファイルの初期化に失敗しました。", vbCritical, "設定初期化エラー"
    End If
    
    Set config = Nothing
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "InitializeSettings: " & Err.Description)
    MsgBox "設定初期化でエラーが発生しました: " & Err.Description, vbCritical, "初期化エラー"
End Sub

' システム情報表示
Public Sub ShowSystemInfo()
    Dim info As String
    
    info = GetSystemInfo() & vbCrLf & vbCrLf
    info = info & GetApplicationVersion() & vbCrLf
    info = info & "メモリ使用量: " & GetMemoryUsage()
    
    MsgBox info, vbInformation, "システム情報"
End Sub

' ヘルプ表示
Public Sub ShowHelp()
    Dim helpMessage As String
    
    helpMessage = "【楽天MS2RSS株価データコレクター ヘルプ】" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "■ 基本的な使い方" & vbCrLf
    helpMessage = helpMessage & "1. 「データ収集開始」ボタンをクリック" & vbCrLf
    helpMessage = helpMessage & "2. 銘柄コード、期間、足種を設定" & vbCrLf
    helpMessage = helpMessage & "3. 出力先フォルダを指定" & vbCrLf
    helpMessage = helpMessage & "4. 「実行」ボタンでデータ取得開始" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "■ 銘柄コード形式" & vbCrLf
    helpMessage = helpMessage & "• 単一銘柄: 7203" & vbCrLf
    helpMessage = helpMessage & "• 複数銘柄: 7203,6758,9984" & vbCrLf
    helpMessage = helpMessage & "• 市場指定: 7203.T, 7203.JAX" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "■ 対応足種" & vbCrLf
    helpMessage = helpMessage & "1M, 5M, 15M, 30M, 60M, D（日足）" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "■ 注意事項" & vbCrLf
    helpMessage = helpMessage & "• MarketSpeed2が起動している必要があります" & vbCrLf
    helpMessage = helpMessage & "• RSS機能を有効にしてください" & vbCrLf
    helpMessage = helpMessage & "• 大量データ取得時は時間がかかります" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "詳細は docs/vba-guide.md を参照してください。"
    
    MsgBox helpMessage, vbInformation, "ヘルプ"
End Sub

' サンプルデータ作成（テスト用）
Public Sub CreateSampleData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sampleData(1 To 10, 1 To 6) As Variant
    Dim i As Long
    
    ' 新しいワークシートを作成
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "サンプルデータ_" & Format(Now, "HHMMSS")
    
    ' ヘッダー設定
    ws.Range("A1:F1").Value = Array("DateTime", "Open", "High", "Low", "Close", "Volume")
    
    ' サンプルデータ生成
    For i = 1 To 10
        sampleData(i, 1) = Format(Now - (10 - i) / 24 / 60 * 5, "YYYY-MM-DD HH:MM:SS") ' 5分間隔
        sampleData(i, 2) = 2500 + Rnd() * 100 ' Open
        sampleData(i, 3) = sampleData(i, 2) + Rnd() * 50 ' High
        sampleData(i, 4) = sampleData(i, 2) - Rnd() * 50 ' Low
        sampleData(i, 5) = sampleData(i, 2) + (Rnd() - 0.5) * 30 ' Close
        sampleData(i, 6) = Int(Rnd() * 100000) + 50000 ' Volume
    Next i
    
    ' データをワークシートに設定
    ws.Range("A2:F11").Value = sampleData
    
    ' 書式設定
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("B2:E11").NumberFormat = "0.00"
    ws.Range("F2:F11").NumberFormat = "#,##0"
    ws.Columns.AutoFit
    
    MsgBox "サンプルデータを作成しました: " & ws.Name, vbInformation, "サンプルデータ作成"
    Call LogMessage(LOG_INFO, "サンプルデータ作成: " & ws.Name)
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "CreateSampleData: " & Err.Description)
    MsgBox "サンプルデータ作成でエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' 全ワークシートマクロのリスト表示
Public Sub ShowMacroList()
    Dim macroList As String
    
    macroList = "【利用可能なマクロ一覧】" & vbCrLf & vbCrLf
    macroList = macroList & "■ データ操作" & vbCrLf
    macroList = macroList & "• StartDataCollection - データ収集開始" & vbCrLf
    macroList = macroList & "• RunQuickTest - クイックテスト実行" & vbCrLf
    macroList = macroList & "• RunAutoSample - 自動実行サンプル" & vbCrLf & vbCrLf
    macroList = macroList & "■ 設定・情報" & vbCrLf
    macroList = macroList & "• DisplaySettings - 設定表示" & vbCrLf
    macroList = macroList & "• InitializeSettings - 設定初期化" & vbCrLf
    macroList = macroList & "• ShowSystemInfo - システム情報" & vbCrLf & vbCrLf
    macroList = macroList & "■ ログ・デバッグ" & vbCrLf
    macroList = macroList & "• ViewLogs - ログ表示" & vbCrLf
    macroList = macroList & "• TestConnection - 接続テスト" & vbCrLf
    macroList = macroList & "• CreateSampleData - サンプルデータ作成" & vbCrLf & vbCrLf
    macroList = macroList & "■ ユーティリティ" & vbCrLf
    macroList = macroList & "• OpenOutputFolder - 出力フォルダを開く" & vbCrLf
    macroList = macroList & "• OpenLogFolder - ログフォルダを開く" & vbCrLf
    macroList = macroList & "• AboutApp - バージョン情報" & vbCrLf
    macroList = macroList & "• ShowHelp - ヘルプ表示"
    
    MsgBox macroList, vbInformation, "マクロ一覧"
End Sub