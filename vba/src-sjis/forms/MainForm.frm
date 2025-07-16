VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Rakuten MS2RSS Stock Data Collector v1.0"
   ClientHeight    =   4000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExecute 
      Caption         =   "Start Data Collection"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   3360
      Width           =   1500
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   3360
      Width           =   1200
   End
   Begin VB.TextBox txtStockCodes 
      Height          =   720
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   4500
   End
   Begin VB.ComboBox cmbTimeFrame 
      Height          =   315
      Left            =   1200
      Style           =   2  'DropDown List
      TabIndex        =   1
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label3 
      Caption         =   "Time Frame:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2440
      Width           =   1000
   End
   Begin VB.Label Label2 
      Caption         =   "Stock Codes:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1240
      Width           =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Collect stock data using Rakuten Securities MarketSpeed2 RSS API and output as CSV."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Enter stock codes separated by commas. Example: 7203,6758,9984"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   5655
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
' Rakuten MS2RSS Stock Data Collector - Main Form
' 
' Description: User interface and main control
' Author: Claude Code
' Version: 1.0.0
'******************************************************************************

Option Explicit

' Form initialization
Private Sub UserForm_Initialize()
    Call InitializeForm
    Call LogMessage(LOG_INFO, "Main form initialized")
End Sub

' Initialize form settings
Private Sub InitializeForm()
    ' Form size and position
    Me.Width = 300
    Me.Height = 250
    
    ' Initialize timeframe combobox
    With Me.cmbTimeFrame
        .Clear
        .AddItem "1M (1 minute)"
        .AddItem "5M (5 minutes)"
        .AddItem "15M (15 minutes)"
        .AddItem "30M (30 minutes)"
        .AddItem "60M (60 minutes)"
        .AddItem "D (Daily)"
        .ListIndex = 1 ' Default to 5 minutes
    End With
    
    ' Set default stock codes
    Me.txtStockCodes.Text = "7203,6758,9984"
End Sub

' Execute button click
Private Sub btnExecute_Click()
    Call StartProcessing()
End Sub

' Start data collection process
Private Sub StartProcessing()
    On Error GoTo ErrorHandler
    
    ' Input validation
    If Not ValidateInputs() Then
        Exit Sub
    End If
    
    ' Execute data collection
    Call ExecuteDataCollection()
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("StartProcessing", Err.Description)
    MsgBox "Error occurred during processing: " & Err.Description, vbCritical
End Sub

' Validate user inputs
Private Function ValidateInputs() As Boolean
    On Error GoTo ErrorHandler
    
    ' Check stock codes
    If Trim(Me.txtStockCodes.Text) = "" Then
        MsgBox "Please enter stock codes.", vbExclamation
        Me.txtStockCodes.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    ' Check timeframe
    If Me.cmbTimeFrame.ListIndex = -1 Then
        MsgBox "Please select a timeframe.", vbExclamation
        Me.cmbTimeFrame.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    ValidateInputs = True
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "ValidateInputs: " & Err.Description)
    ValidateInputs = False
End Function

' Execute data collection
Private Sub ExecuteDataCollection()
    On Error GoTo ErrorHandler
    
    Dim stockCodes As String
    Dim timeFrame As String
    Dim stocks() As String
    Dim i As Long
    Dim successCount As Long
    
    ' Get parameters
    stockCodes = Trim(Me.txtStockCodes.Text)
    timeFrame = GetSelectedTimeFrame()
    
    ' Split stock codes
    stocks = Split(Replace(stockCodes, " ", ""), ",")
    
    Call LogMessage(LOG_INFO, "Data collection started: " & (UBound(stocks) + 1) & " stocks")
    
    ' Process each stock
    For i = 0 To UBound(stocks)
        If Trim(stocks(i)) <> "" Then
            ' Execute data collection
            If CollectStockData(Trim(stocks(i)), timeFrame, Date - 1, Date) Then
                successCount = successCount + 1
                Call LogMessage(LOG_INFO, "Success: " & stocks(i))
            Else
                Call LogMessage(LOG_ERROR, "Failed: " & stocks(i))
            End If
        End If
    Next i
    
    ' Show results
    Call ShowResults(successCount, UBound(stocks) + 1)
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("ExecuteDataCollection", Err.Description)
End Sub

' Get selected timeframe
Private Function GetSelectedTimeFrame() As String
    Dim selectedText As String
    selectedText = Me.cmbTimeFrame.Text
    
    ' Extract "1M" from "1M (1 minute)"
    GetSelectedTimeFrame = Left(selectedText, InStr(selectedText, " ") - 1)
End Function

' Show results
Private Sub ShowResults(successCount As Long, totalCount As Long)
    Dim message As String
    
    message = "Processing completed" & vbCrLf & vbCrLf
    message = message & "Success: " & successCount & " stocks" & vbCrLf
    message = message & "Failed: " & (totalCount - successCount) & " stocks" & vbCrLf
    message = message & "Total: " & totalCount & " stocks"
    
    If successCount = totalCount Then
        MsgBox message, vbInformation, "Completed"
    Else
        MsgBox message, vbExclamation, "Completed (with errors)"
    End If
    
    Call LogMessage(LOG_INFO, "Processing completed: " & successCount & "/" & totalCount)
End Sub

' Cancel button click
Private Sub btnCancel_Click()
    Unload Me
End Sub

' Form close
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call LogMessage(LOG_INFO, "Main form closed")
End Sub