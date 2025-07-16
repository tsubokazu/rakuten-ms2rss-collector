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

' バージョン情報ボタン用マクロ
Public Sub AboutApp()
    Call ShowAbout
End Sub

' 出力フォルダを開く
Public Sub OpenOutputFolder()
    On Error GoTo ErrorHandler
    
    Dim outputPath As String
    outputPath = ThisWorkbook.Path & "\output\csv\"
    
    ' フォルダが存在しない場合は作成
    If Dir(outputPath, vbDirectory) = "" Then
        MkDir outputPath
    End If
    
    ' フォルダを開く
    Shell "explorer.exe " & Chr(34) & outputPath & Chr(34), vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "フォルダを開くことができませんでした: " & Err.Description, vbCritical
End Sub

' MarketSpeed2接続テスト
Public Sub TestConnection()
    On Error GoTo ErrorHandler
    
    Dim testResult As Variant
    
    ' 日経平均現在値で接続テスト
    testResult = Application.WorksheetFunction.RssIndexMarket("0000", "現在値")
    
    If IsError(testResult) Then
        MsgBox "MarketSpeed2への接続に失敗しました。" & vbCrLf & _
               "MarketSpeed2が起動しているか確認してください。", vbExclamation
    Else
        MsgBox "MarketSpeed2への接続に成功しました！" & vbCrLf & _
               "日経平均現在値: " & testResult, vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "接続テストでエラーが発生しました: " & Err.Description, vbCritical
End Sub

' ヘルプ表示
Public Sub ShowHelp()
    Dim helpMessage As String
    
    helpMessage = "【楽天MS2RSS株価データコレクター ヘルプ】" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "■ 基本的な使い方" & vbCrLf
    helpMessage = helpMessage & "1. 「データ収集開始」ボタンをクリック" & vbCrLf
    helpMessage = helpMessage & "2. 銘柄コードを入力（例：7203,6758,9984）" & vbCrLf
    helpMessage = helpMessage & "3. 「実行」ボタンでデータ取得開始" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "■ 注意事項" & vbCrLf
    helpMessage = helpMessage & "• MarketSpeed2が起動している必要があります" & vbCrLf
    helpMessage = helpMessage & "• RSS機能を有効にしてください"
    
    MsgBox helpMessage, vbInformation, "ヘルプ"
End Sub