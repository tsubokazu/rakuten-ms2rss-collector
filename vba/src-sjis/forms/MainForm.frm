VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "�y�VMS2RSS�����f�[�^�R���N�^�[ v1.0"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnHelp 
      Caption         =   "�w���v"
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   5400
      Width           =   1200
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "���~"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   5400
      Width           =   1200
   End
   Begin VB.CommandButton btnExecute 
      Caption         =   "�f�[�^�擾�J�n"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   5400
      Width           =   1200
   End
   Begin VB.CommandButton btnSelectOutput 
      Caption         =   "�Q��..."
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
      Caption         =   "�ҋ@��"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   8760
   End
   Begin VB.Label Label6 
      Caption         =   "�o�͐�t�H���_:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4480
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "����:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3760
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "�I����:"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   3040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "�J�n��:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3040
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "�����R�[�h:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "�y�V�،�MarketSpeed2��RSS API���g�p���Ċ����f�[�^���擾���ACSV�`���ŏo�͂��܂��B"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label lblInstructions 
      Caption         =   "�����R�[�h���J���}��؂�œ��͂��Ă��������B��: 7203,6758,9984"
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
' �y�VMS2RSS�����f�[�^�R���N�^�[ - ���C���t�H�[��
' 
' ����: ���[�U�[�C���^�[�t�F�[�X�ƃ��C������
' �쐬��: Claude Code
' �o�[�W����: 1.0.0
'******************************************************************************

Option Explicit

' �t�H�[�����̐���ϐ�
Private isProcessing As Boolean
Private currentProgress As Long
Private totalProgress As Long

' �t�H�[��������
Private Sub UserForm_Initialize()
    Call InitializeForm
    Call LoadDefaultSettings
    Call LogMessage(LOG_INFO, "���C���t�H�[�������������܂���")
End Sub

' �t�H�[���ݒ菉����
Private Sub InitializeForm()
    ' �t�H�[���̃T�C�Y�ƈʒu
    Me.Width = 450
    Me.Height = 600
    
    ' ����R���{�{�b�N�X�̏�����
    With Me.cmbTimeFrame
        .Clear
        .AddItem "1M (1����)"
        .AddItem "5M (5����)"
        .AddItem "15M (15����)"
        .AddItem "30M (30����)"
        .AddItem "60M (60����)"
        .AddItem "D (����)"
        .ListIndex = 1 ' �f�t�H���g��5����
    End With
    
    ' ���t�̏�����
    Me.dtpStartDate.Value = DateAdd("d", -30, Date) ' 30���O
    Me.dtpEndDate.Value = Date ' ����
    
    ' �v���O���X�o�[�̏�����
    Me.pgbProgress.Value = 0
    
    ' �{�^���̏������
    Me.btnExecute.Caption = "�f�[�^�擾�J�n"
    Me.btnCancel.Enabled = False
    
    isProcessing = False
End Sub

' �f�t�H���g�ݒ�̓ǂݍ���
Private Sub LoadDefaultSettings()
    On Error Resume Next
    
    ' �O��̐ݒ�𕜌��i���W�X�g���܂��͐ݒ�t�@�C������j
    Me.txtStockCodes.Text = GetSetting("MS2RSSCollector", "Settings", "LastStockCodes", "7203,6758,9984")
    Me.txtOutputPath.Text = GetSetting("MS2RSSCollector", "Settings", "LastOutputPath", ThisWorkbook.Path & "\output\csv\")
    
    ' �f�B���N�g�����݊m�F
    If Not EnsureDirectoryExists(Me.txtOutputPath.Text) Then
        Me.txtOutputPath.Text = ThisWorkbook.Path & "\output\csv\"
        Call EnsureDirectoryExists(Me.txtOutputPath.Text)
    End If
End Sub

' �f�[�^�擾�J�n�{�^��
Private Sub btnExecute_Click()
    If isProcessing Then
        Call StopProcessing()
    Else
        Call StartProcessing()
    End If
End Sub

' �f�[�^�擾�����J�n
Private Sub StartProcessing()
    On Error GoTo ErrorHandler
    
    ' ���͒l����
    If Not ValidateInputs() Then
        Exit Sub
    End If
    
    ' �ݒ�ۑ�
    Call SaveCurrentSettings()
    
    ' UI��ԍX�V
    isProcessing = True
    Me.btnExecute.Caption = "�������~"
    Me.btnCancel.Enabled = True
    Me.pgbProgress.Value = 0
    
    ' �A�v���P�[�V�����œK��
    Call OptimizeApplicationSettings()
    
    ' ���C���������s
    Call ExecuteDataCollection()
    
    ' ��������
    Call CompleteProcessing()
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("StartProcessing", Err.Description)
    Call CompleteProcessing()
    MsgBox "�������ɃG���[���������܂���: " & Err.Description, vbCritical
End Sub

' ���͒l�̌���
Private Function ValidateInputs() As Boolean
    On Error GoTo ErrorHandler
    
    ' �����R�[�h�`�F�b�N
    If Trim(Me.txtStockCodes.Text) = "" Then
        MsgBox "�����R�[�h����͂��Ă��������B", vbExclamation
        Me.txtStockCodes.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    ' ���t�͈̓`�F�b�N
    If Not ValidateDateRange(Me.dtpStartDate.Value, Me.dtpEndDate.Value) Then
        MsgBox "���t�͈͂�����������܂���B", vbExclamation
        Me.dtpStartDate.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    ' ����`�F�b�N
    If Me.cmbTimeFrame.ListIndex = -1 Then
        MsgBox "�����I�����Ă��������B", vbExclamation
        Me.cmbTimeFrame.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    ' �o�̓p�X�`�F�b�N
    If Trim(Me.txtOutputPath.Text) = "" Then
        MsgBox "�o�͐�t�H���_���w�肵�Ă��������B", vbExclamation
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

' ���C���f�[�^���W����
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
    
    ' �p�����[�^�擾
    stockCodes = Trim(Me.txtStockCodes.Text)
    timeFrame = GetSelectedTimeFrame()
    startDate = Me.dtpStartDate.Value
    endDate = Me.dtpEndDate.Value
    outputPath = Trim(Me.txtOutputPath.Text)
    
    ' ��������
    stocks = Split(Replace(stockCodes, " ", ""), ",")
    totalProgress = UBound(stocks) + 1
    currentProgress = 0
    
    Call LogMessage(LOG_INFO, "�f�[�^���W�J�n: " & totalProgress & "����")
    
    ' �e����������
    For i = 0 To UBound(stocks)
        If Not isProcessing Then Exit For ' ���~�`�F�b�N
        
        If Trim(stocks(i)) <> "" Then
            ' �i���X�V
            currentProgress = i + 1
            Call UpdateProgress(currentProgress, totalProgress, "������: " & stocks(i))
            
            ' �f�[�^���W���s
            If CollectStockData(Trim(stocks(i)), timeFrame, startDate, endDate, _
                              outputPath & GenerateFileName(stocks(i), timeFrame, startDate, endDate)) Then
                successCount = successCount + 1
                Call LogMessage(LOG_INFO, "����: " & stocks(i))
            Else
                Call LogMessage(LOG_ERROR, "���s: " & stocks(i))
            End If
            
            DoEvents
        End If
    Next i
    
    ' ���ʕ\��
    Call ShowResults(successCount, totalProgress)
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("ExecuteDataCollection", Err.Description)
End Sub

' �t�@�C��������
Private Function GenerateFileName(stockCode As String, timeFrame As String, _
                                startDate As Date, endDate As Date) As String
    
    Dim fileName As String
    fileName = SafeFileName(Replace(stockCode, ".", "_")) & "_" & timeFrame & "_" & _
               Format(startDate, "YYYYMMDD") & "-" & Format(endDate, "YYYYMMDD") & ".csv"
    
    GenerateFileName = fileName
End Function

' �I�����ꂽ������擾
Private Function GetSelectedTimeFrame() As String
    Dim selectedText As String
    selectedText = Me.cmbTimeFrame.Text
    
    ' "1M (1����)" ���� "1M" �𒊏o
    GetSelectedTimeFrame = Left(selectedText, InStr(selectedText, " ") - 1)
End Function

' �i���X�V
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

' ���ʕ\��
Private Sub ShowResults(successCount As Long, totalCount As Long)
    Dim message As String
    
    message = "��������" & vbCrLf & vbCrLf
    message = message & "����: " & successCount & "����" & vbCrLf
    message = message & "���s: " & (totalCount - successCount) & "����" & vbCrLf
    message = message & "���v: " & totalCount & "����"
    
    If successCount = totalCount Then
        MsgBox message, vbInformation, "��������"
    Else
        MsgBox message, vbExclamation, "���������i�ꕔ�G���[�j"
    End If
    
    Call LogMessage(LOG_INFO, "��������: " & successCount & "/" & totalCount)
End Sub

' ������~
Private Sub StopProcessing()
    isProcessing = False
    Call LogMessage(LOG_INFO, "���[�U�[�ɂ���ď��������~����܂���")
End Sub

' �����������̌㏈��
Private Sub CompleteProcessing()
    ' UI��ԕ���
    isProcessing = False
    Me.btnExecute.Caption = "�f�[�^�擾�J�n"
    Me.btnCancel.Enabled = False
    Me.lblStatus.Caption = "�ҋ@��"
    
    ' �A�v���P�[�V�����ݒ蕜��
    Call RestoreApplicationSettings()
    Call ClearProgress()
End Sub

' ���݂̐ݒ��ۑ�
Private Sub SaveCurrentSettings()
    On Error Resume Next
    
    SaveSetting "MS2RSSCollector", "Settings", "LastStockCodes", Me.txtStockCodes.Text
    SaveSetting "MS2RSSCollector", "Settings", "LastOutputPath", Me.txtOutputPath.Text
End Sub

' �o�̓t�H���_�I���{�^��
Private Sub btnSelectOutput_Click()
    Dim selectedPath As String
    
    selectedPath = SelectFolder("�o�͐�t�H���_��I�����Ă�������")
    
    If selectedPath <> "" Then
        Me.txtOutputPath.Text = selectedPath
        If Right(selectedPath, 1) <> "\" Then
            Me.txtOutputPath.Text = selectedPath & "\"
        End If
    End If
End Sub

' �t�H���_�I���_�C�A���O
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

' �L�����Z���{�^��
Private Sub btnCancel_Click()
    Call StopProcessing()
End Sub

' �t�H�[���N���[�Y��
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If isProcessing Then
        If MsgBox("�������ł��B�{���ɏI�����܂����H", vbYesNo + vbQuestion) = vbNo Then
            Cancel = True
            Exit Sub
        End If
        Call StopProcessing()
    End If
    
    Call SaveCurrentSettings()
    Call LogMessage(LOG_INFO, "���C���t�H�[�����I�����܂���")
End Sub

' �w���v�{�^��
Private Sub btnHelp_Click()
    Dim helpMessage As String
    
    helpMessage = "�y�VMS2RSS�����f�[�^�R���N�^�[" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "�y�g�p���@�z" & vbCrLf
    helpMessage = helpMessage & "1. �����R�[�h���J���}��؂�œ���" & vbCrLf
    helpMessage = helpMessage & "   ��: 7203,6758,9984" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "2. �擾���ԂƑ����I��" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "3. �o�͐�t�H���_���w��" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "4. �u�f�[�^�擾�J�n�v���N���b�N" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "�y�Ή��s��z" & vbCrLf
    helpMessage = helpMessage & "T:����, JAX:JAX, JNX:JNX, CHJ:Chi-X" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "�ڍׂ̓h�L�������g���Q�Ƃ��Ă��������B"
    
    MsgBox helpMessage, vbInformation, "�w���v"
End Sub