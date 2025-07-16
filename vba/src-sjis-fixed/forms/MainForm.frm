VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "楽天MS2RSS株価データコレクター v1.0"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExecute 
      Caption         =   "データ取得開始"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   5400
      Width           =   1200
   End
   Begin VB.TextBox txtStockCodes 
      Height          =   720
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   5760
   End
   Begin VB.Label Label1 
      Caption         =   "楽天証券MarketSpeed2のRSS APIを使用して株価データを取得し、CSV形式で出力します。"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label Label2 
      Caption         =   "銘柄コード:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1840
      Width           =   1575
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' フォーム初期化
Private Sub UserForm_Initialize()
    Me.txtStockCodes.Text = "7203,6758,9984"
End Sub

' 実行ボタン
Private Sub btnExecute_Click()
    Dim stockCodes As String
    Dim result As Boolean
    
    stockCodes = Trim(Me.txtStockCodes.Text)
    
    If stockCodes = "" Then
        MsgBox "銘柄コードを入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 簡単なデータ取得テスト
    result = CollectMultipleStocks(stockCodes, "5M", Date - 1, Date)
    
    If result Then
        MsgBox "データ取得完了", vbInformation
    Else
        MsgBox "データ取得に失敗しました", vbExclamation
    End If
End Sub