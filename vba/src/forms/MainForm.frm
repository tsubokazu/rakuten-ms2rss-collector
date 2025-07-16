VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "楽天MS2RSS株価データコレクター v1.0"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnHelp 
      Caption         =   "ヘルプ"
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   5400
      Width           =   1200
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "中止"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   5400
      Width           =   1200
   End
   Begin VB.CommandButton btnExecute 
      Caption         =   "データ取得開始"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   5400
      Width           =   1200
   End
   Begin VB.CommandButton btnSelectOutput 
      Caption         =   "参照..."
      Height          =   315
      Left            =   7680
      TabIndex        =   12
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtOutputPath 
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      Top             =   4440
      Width           =   5760
   End
   Begin VB.ComboBox cmbTimeFrame 
      Height          =   315
      Left            =   1800
      Style           =   2  'DropDown List
      TabIndex        =   9
      Top             =   3720
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   5400
      TabIndex        =   7
      Top             =   3000
      Width           =   1815
      Format          =   "yyyy/MM/dd"
      CurrentDate     =   44576
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
      Format          =   "yyyy/MM/dd"
      CurrentDate     =   44576
   End
   Begin VB.TextBox txtStockCodes 
      Height          =   720
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1800
      Width           =   5760
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   8760
      Appearance      =   1
      Max             =   100
      Scrolling       =   1
   End
   Begin VB.Label lblStatus 
      Caption         =   "待機中"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   8760
   End
   Begin VB.Label Label6 
      Caption         =   "出力先フォルダ:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4480
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "足種:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3760
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "終了日:"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   3040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "開始日:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3040
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "銘柄コード:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "楽天証券MarketSpeed2のRSS APIを使用して株価データを取得し、CSV形式で出力します。"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label lblInstructions 
      Caption         =   "銘柄コードをカンマ区切りで入力してください。例: 7203,6758,9984"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   8655
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' 楽天MS2RSS株価データコレクター - メインフォーム
' 
' 説明: ユーザーインターフェースとメイン制御
' 作成者: Claude Code
' バージョン: 1.0.0
'******************************************************************************

Option Explicit

' フォーム内の制御変数
Private isProcessing As Boolean
Private currentProgress As Long
Private totalProgress As Long

' フォーム初期化
Private Sub UserForm_Initialize()
    Call InitializeForm
    Call LoadDefaultSettings
    Call LogMessage(LOG_INFO, "メインフォームを初期化しました")
End Sub

' フォーム設定初期化
Private Sub InitializeForm()
    ' フォームのサイズと位置
    Me.Width = 450
    Me.Height = 600
    
    ' 足種コンボボックスの初期化
    With Me.cmbTimeFrame
        .Clear
        .AddItem "1M (1分足)"
        .AddItem "5M (5分足)"
        .AddItem "15M (15分足)"
        .AddItem "30M (30分足)"
        .AddItem "60M (60分足)"
        .AddItem "D (日足)"
        .ListIndex = 1 ' デフォルトは5分足
    End With
    
    ' 日付の初期化
    Me.dtpStartDate.Value = DateAdd("d", -30, Date) ' 30日前
    Me.dtpEndDate.Value = Date ' 今日
    
    ' プログレスバーの初期化
    Me.pgbProgress.Value = 0
    
    ' ボタンの初期状態
    Me.btnExecute.Caption = "データ取得開始"
    Me.btnCancel.Enabled = False
    
    isProcessing = False
End Sub

' デフォルト設定の読み込み
Private Sub LoadDefaultSettings()
    On Error Resume Next
    
    ' 前回の設定を復元（レジストリまたは設定ファイルから）
    Me.txtStockCodes.Text = GetSetting("MS2RSSCollector", "Settings", "LastStockCodes", "7203,6758,9984")
    Me.txtOutputPath.Text = GetSetting("MS2RSSCollector", "Settings", "LastOutputPath", ThisWorkbook.Path & "\output\csv\")
    
    ' ディレクトリ存在確認
    If Not EnsureDirectoryExists(Me.txtOutputPath.Text) Then
        Me.txtOutputPath.Text = ThisWorkbook.Path & "\output\csv\"
        Call EnsureDirectoryExists(Me.txtOutputPath.Text)
    End If
End Sub

' データ取得開始ボタン
Private Sub btnExecute_Click()
    If isProcessing Then
        Call StopProcessing()
    Else
        Call StartProcessing()
    End If
End Sub

' データ取得処理開始
Private Sub StartProcessing()
    On Error GoTo ErrorHandler
    
    ' 入力値検証
    If Not ValidateInputs() Then
        Exit Sub
    End If
    
    ' 設定保存
    Call SaveCurrentSettings()
    
    ' UI状態更新
    isProcessing = True
    Me.btnExecute.Caption = "処理中止"
    Me.btnCancel.Enabled = True
    Me.pgbProgress.Value = 0
    
    ' アプリケーション最適化
    Call OptimizeApplicationSettings()
    
    ' メイン処理実行
    Call ExecuteDataCollection()
    
    ' 完了処理
    Call CompleteProcessing()
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("StartProcessing", Err.Description)
    Call CompleteProcessing()
    MsgBox "処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 入力値の検証
Private Function ValidateInputs() As Boolean
    On Error GoTo ErrorHandler
    
    ' 銘柄コードチェック
    If Trim(Me.txtStockCodes.Text) = "" Then
        MsgBox "銘柄コードを入力してください。", vbExclamation
        Me.txtStockCodes.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    ' 日付範囲チェック
    If Not ValidateDateRange(Me.dtpStartDate.Value, Me.dtpEndDate.Value) Then
        MsgBox "日付範囲が正しくありません。", vbExclamation
        Me.dtpStartDate.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    ' 足種チェック
    If Me.cmbTimeFrame.ListIndex = -1 Then
        MsgBox "足種を選択してください。", vbExclamation
        Me.cmbTimeFrame.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    ' 出力パスチェック
    If Trim(Me.txtOutputPath.Text) = "" Then
        MsgBox "出力先フォルダを指定してください。", vbExclamation
        Me.txtOutputPath.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    ValidateInputs = True
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "ValidateInputs: " & Err.Description)
    ValidateInputs = False
End Function

' メインデータ収集処理
Private Sub ExecuteDataCollection()
    On Error GoTo ErrorHandler
    
    Dim stockCodes As String
    Dim timeFrame As String
    Dim startDate As Date
    Dim endDate As Date
    Dim outputPath As String
    Dim stocks() As String
    Dim i As Long
    Dim successCount As Long
    
    ' パラメータ取得
    stockCodes = Trim(Me.txtStockCodes.Text)
    timeFrame = GetSelectedTimeFrame()
    startDate = Me.dtpStartDate.Value
    endDate = Me.dtpEndDate.Value
    outputPath = Trim(Me.txtOutputPath.Text)
    
    ' 銘柄分割
    stocks = Split(Replace(stockCodes, " ", ""), ",")
    totalProgress = UBound(stocks) + 1
    currentProgress = 0
    
    Call LogMessage(LOG_INFO, "データ収集開始: " & totalProgress & "銘柄")
    
    ' 各銘柄を処理
    For i = 0 To UBound(stocks)
        If Not isProcessing Then Exit For ' 中止チェック
        
        If Trim(stocks(i)) <> "" Then
            ' 進捗更新
            currentProgress = i + 1
            Call UpdateProgress(currentProgress, totalProgress, "処理中: " & stocks(i))
            
            ' データ収集実行
            If CollectStockData(Trim(stocks(i)), timeFrame, startDate, endDate, _
                              outputPath & GenerateFileName(stocks(i), timeFrame, startDate, endDate)) Then
                successCount = successCount + 1
                Call LogMessage(LOG_INFO, "成功: " & stocks(i))
            Else
                Call LogMessage(LOG_ERROR, "失敗: " & stocks(i))
            End If
            
            DoEvents
        End If
    Next i
    
    ' 結果表示
    Call ShowResults(successCount, totalProgress)
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("ExecuteDataCollection", Err.Description)
End Sub

' ファイル名生成
Private Function GenerateFileName(stockCode As String, timeFrame As String, _
                                startDate As Date, endDate As Date) As String
    
    Dim fileName As String
    fileName = SafeFileName(Replace(stockCode, ".", "_")) & "_" & timeFrame & "_" & _
               Format(startDate, "YYYYMMDD") & "-" & Format(endDate, "YYYYMMDD") & ".csv"
    
    GenerateFileName = fileName
End Function

' 選択された足種を取得
Private Function GetSelectedTimeFrame() As String
    Dim selectedText As String
    selectedText = Me.cmbTimeFrame.Text
    
    ' "1M (1分足)" から "1M" を抽出
    GetSelectedTimeFrame = Left(selectedText, InStr(selectedText, " ") - 1)
End Function

' 進捗更新
Private Sub UpdateProgress(current As Long, total As Long, Optional message As String = "")
    Dim percentage As Integer
    
    If total = 0 Then Exit Sub
    
    percentage = Int((current / total) * 100)
    Me.pgbProgress.Value = percentage
    
    If message <> "" Then
        Me.lblStatus.Caption = message
    End If
    
    DoEvents
End Sub

' 結果表示
Private Sub ShowResults(successCount As Long, totalCount As Long)
    Dim message As String
    
    message = "処理完了" & vbCrLf & vbCrLf
    message = message & "成功: " & successCount & "銘柄" & vbCrLf
    message = message & "失敗: " & (totalCount - successCount) & "銘柄" & vbCrLf
    message = message & "合計: " & totalCount & "銘柄"
    
    If successCount = totalCount Then
        MsgBox message, vbInformation, "処理完了"
    Else
        MsgBox message, vbExclamation, "処理完了（一部エラー）"
    End If
    
    Call LogMessage(LOG_INFO, "処理完了: " & successCount & "/" & totalCount)
End Sub

' 処理停止
Private Sub StopProcessing()
    isProcessing = False
    Call LogMessage(LOG_INFO, "ユーザーによって処理が中止されました")
End Sub

' 処理完了時の後処理
Private Sub CompleteProcessing()
    ' UI状態復元
    isProcessing = False
    Me.btnExecute.Caption = "データ取得開始"
    Me.btnCancel.Enabled = False
    Me.lblStatus.Caption = "待機中"
    
    ' アプリケーション設定復元
    Call RestoreApplicationSettings()
    Call ClearProgress()
End Sub

' 現在の設定を保存
Private Sub SaveCurrentSettings()
    On Error Resume Next
    
    SaveSetting "MS2RSSCollector", "Settings", "LastStockCodes", Me.txtStockCodes.Text
    SaveSetting "MS2RSSCollector", "Settings", "LastOutputPath", Me.txtOutputPath.Text
End Sub

' 出力フォルダ選択ボタン
Private Sub btnSelectOutput_Click()
    Dim selectedPath As String
    
    selectedPath = SelectFolder("出力先フォルダを選択してください")
    
    If selectedPath <> "" Then
        Me.txtOutputPath.Text = selectedPath
        If Right(selectedPath, 1) <> "\" Then
            Me.txtOutputPath.Text = selectedPath & "\"
        End If
    End If
End Sub

' フォルダ選択ダイアログ
Private Function SelectFolder(title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .title = title
        .InitialFileName = Me.txtOutputPath.Text
        
        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        Else
            SelectFolder = ""
        End If
    End With
    
    Set fd = Nothing
End Function

' キャンセルボタン
Private Sub btnCancel_Click()
    Call StopProcessing()
End Sub

' フォームクローズ時
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If isProcessing Then
        If MsgBox("処理中です。本当に終了しますか？", vbYesNo + vbQuestion) = vbNo Then
            Cancel = True
            Exit Sub
        End If
        Call StopProcessing()
    End If
    
    Call SaveCurrentSettings()
    Call LogMessage(LOG_INFO, "メインフォームを終了しました")
End Sub

' ヘルプボタン
Private Sub btnHelp_Click()
    Dim helpMessage As String
    
    helpMessage = "楽天MS2RSS株価データコレクター" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "【使用方法】" & vbCrLf
    helpMessage = helpMessage & "1. 銘柄コードをカンマ区切りで入力" & vbCrLf
    helpMessage = helpMessage & "   例: 7203,6758,9984" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "2. 取得期間と足種を選択" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "3. 出力先フォルダを指定" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "4. 「データ取得開始」をクリック" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "【対応市場】" & vbCrLf
    helpMessage = helpMessage & "T:東証, JAX:JAX, JNX:JNX, CHJ:Chi-X" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "詳細はドキュメントを参照してください。"
    
    MsgBox helpMessage, vbInformation, "ヘルプ"
End Sub