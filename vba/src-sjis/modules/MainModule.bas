Attribute VB_Name = "MainModule"
'******************************************************************************
' �y�VMS2RSS�����f�[�^�R���N�^�[ - ���C�����W���[��
' 
' ����: �A�v���P�[�V�����̃G���g���[�|�C���g�ƃ��C������
' �쐬��: Claude Code
' �o�[�W����: 1.0.0
'******************************************************************************

Option Explicit

' �A�v���P�[�V�������
Public Const APP_NAME As String = "�y�VMS2RSS�����f�[�^�R���N�^�["
Public Const APP_VERSION As String = "1.0.0"
Public Const BUILD_DATE As String = "2025-01-16"

' ���C���t�H�[����\��
Public Sub ShowMainForm()
    On Error GoTo ErrorHandler
    
    ' ���O������
    Call LogMessage(LOG_INFO, "�A�v���P�[�V�����J�n: " & APP_NAME & " v" & APP_VERSION)
    
    ' �����ݒ�`�F�b�N
    If Not CheckInitialSetup() Then
        MsgBox "�����ݒ�ɖ�肪����܂��B���O���m�F���Ă��������B", vbCritical, APP_NAME
        Exit Sub
    End If
    
    ' ���C���t�H�[���\��
    Load MainForm
    MainForm.Show vbModal
    
    ' �t�H�[��������ꂽ��̃N���[���A�b�v
    Unload MainForm
    Set MainForm = Nothing
    
    Call LogMessage(LOG_INFO, "�A�v���P�[�V�����I��")
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("ShowMainForm", Err.Description)
    MsgBox "�A�v���P�[�V�����̋N���ŃG���[���������܂���: " & Err.Description, vbCritical, APP_NAME
End Sub

' �����ݒ�`�F�b�N
Private Function CheckInitialSetup() As Boolean
    On Error GoTo ErrorHandler
    
    Dim setupOK As Boolean
    setupOK = True
    
    ' 1. �o�̓f�B���N�g���̑��݊m�F�E�쐬
    If Not EnsureDirectoryExists(ThisWorkbook.Path & "\output\csv\") Then
        Call LogMessage(LOG_ERROR, "CSV�o�̓f�B���N�g���̍쐬�Ɏ��s���܂���")
        setupOK = False
    End If
    
    If Not EnsureDirectoryExists(ThisWorkbook.Path & "\output\logs\") Then
        Call LogMessage(LOG_ERROR, "���O�f�B���N�g���̍쐬�Ɏ��s���܂���")
        setupOK = False
    End If
    
    ' 2. �ݒ�t�@�C���̊m�F
    Dim config As Configuration
    Set config = New Configuration
    
    If Not config.LoadFromFile() Then
        Call LogMessage(LOG_WARN, "�ݒ�t�@�C���̓ǂݍ��݂Ɏ��s�A�f�t�H���g�ݒ���g�p���܂�")
        config.SaveToFile ' �f�t�H���g�ݒ��ۑ�
    End If
    
    If Not config.ValidateSettings() Then
        Call LogMessage(LOG_ERROR, "�ݒ�l�ɖ�肪����܂�")
        setupOK = False
    End If
    
    Set config = Nothing
    
    ' 3. MarketSpeed2�ڑ��e�X�g
    If Not TestMS2Connection() Then
        Call LogMessage(LOG_WARN, "MarketSpeed2�ւ̐ڑ��e�X�g�Ɏ��s���܂���")
        ' �x���݂̂ő��s�i�I�t���C���ł�VBA�R�[�h�͊m�F�\�j
    End If
    
    CheckInitialSetup = setupOK
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "CheckInitialSetup: " & Err.Description)
    CheckInitialSetup = False
End Function

' MarketSpeed2�ڑ��e�X�g
Private Function TestMS2Connection() As Boolean
    On Error GoTo ErrorHandler
    
    Dim testResult As Variant
    
    ' �ȈՐڑ��e�X�g�i���o���ς̌��ݒl���擾�j
    testResult = Application.WorksheetFunction.RssIndexMarket("0000", "���ݒl")
    
    If IsError(testResult) Then
        Call LogMessage(LOG_WARN, "MS2�ڑ��e�X�g���s: RSS�֐����G���[��Ԃ��܂���")
        TestMS2Connection = False
    Else
        Call LogMessage(LOG_INFO, "MS2�ڑ��e�X�g����: " & testResult)
        TestMS2Connection = True
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_WARN, "MS2�ڑ��e�X�g��O: " & Err.Description)
    TestMS2Connection = False
End Function

' �ȒP�e�X�g���s
Public Sub QuickTest()
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim testStockCode As String
    Dim testTimeFrame As String
    Dim testStartDate As Date
    Dim testEndDate As Date
    
    ' �e�X�g�p�����[�^
    testStockCode = "7203"  ' �g���^������
    testTimeFrame = "5M"    ' 5����
    testStartDate = Date - 1  ' ���
    testEndDate = Date        ' ����
    
    Call LogMessage(LOG_INFO, "�N�C�b�N�e�X�g�J�n: " & testStockCode)
    
    ' �f�[�^�擾�e�X�g
    result = CollectStockData(testStockCode, testTimeFrame, testStartDate, testEndDate)
    
    If result Then
        MsgBox "�N�C�b�N�e�X�g�����I" & vbCrLf & _
               "����: " & testStockCode & vbCrLf & _
               "����: " & testTimeFrame & vbCrLf & _
               "����: " & Format(testStartDate, "MM/DD") & " - " & Format(testEndDate, "MM/DD"), _
               vbInformation, "�e�X�g����"
        Call LogMessage(LOG_INFO, "�N�C�b�N�e�X�g����")
    Else
        MsgBox "�N�C�b�N�e�X�g���s�B���O���m�F���Ă��������B", vbExclamation, "�e�X�g����"
        Call LogMessage(LOG_ERROR, "�N�C�b�N�e�X�g���s")
    End If
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("QuickTest", Err.Description, "����: " & testStockCode)
    MsgBox "�e�X�g���s���ɃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
End Sub

' ���݂̐ݒ��\��
Public Sub ShowCurrentConfig()
    On Error GoTo ErrorHandler
    
    Dim config As Configuration
    Set config = New Configuration
    
    If config.LoadFromFile() Then
        MsgBox config.ToString(), vbInformation, "���݂̐ݒ�"
    Else
        MsgBox "�ݒ�t�@�C���̓ǂݍ��݂Ɏ��s���܂����B", vbExclamation, "�ݒ�G���["
    End If
    
    Set config = Nothing
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "ShowCurrentConfig: " & Err.Description)
    MsgBox "�ݒ�\���ŃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
End Sub

' ���O�r���[�A�[��\��
Public Sub ShowLogViewer()
    On Error GoTo ErrorHandler
    
    Dim logFilePath As String
    Dim logContent As String
    Dim fileNum As Integer
    
    ' �����̃��O�t�@�C���p�X
    logFilePath = ThisWorkbook.Path & "\output\logs\ms2rss_collector_" & Format(Date, "YYYYMMDD") & ".log"
    
    ' ���O�t�@�C�����݃`�F�b�N
    If Dir(logFilePath) = "" Then
        MsgBox "�{���̃��O�t�@�C����������܂���B" & vbCrLf & logFilePath, vbInformation, "���O�r���[�A�["
        Exit Sub
    End If
    
    ' ���O�t�@�C���ǂݍ���
    fileNum = FreeFile
    Open logFilePath For Input As #fileNum
    logContent = Input(LOF(fileNum), fileNum)
    Close #fileNum
    
    ' ���O���e�����b�Z�[�W�{�b�N�X�ŕ\���i�ȈՔŁj
    If Len(logContent) > 1000 Then
        logContent = "... (�ȗ�) ..." & vbCrLf & Right(logContent, 800)
    End If
    
    MsgBox "�y�{���̃��O�z" & vbCrLf & vbCrLf & logContent, vbInformation, "���O�r���[�A�["
    
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage(LOG_ERROR, "ShowLogViewer: " & Err.Description)
    MsgBox "���O�\���ŃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
End Sub

' �o�[�W��������\��
Public Sub ShowAbout()
    Dim aboutMessage As String
    
    aboutMessage = APP_NAME & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "�o�[�W����: " & APP_VERSION & vbCrLf
    aboutMessage = aboutMessage & "�r���h��: " & BUILD_DATE & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "�y�V�،�MarketSpeed2��RSS API���g�p����" & vbCrLf
    aboutMessage = aboutMessage & "�����f�[�^���擾���ACSV�`���ŏo�͂��܂��B" & vbCrLf & vbCrLf
    aboutMessage = aboutMessage & "Created with Claude Code" & vbCrLf
    aboutMessage = aboutMessage & "https://github.com/tsubokazu/rakuten-ms2rss-collector"
    
    MsgBox aboutMessage, vbInformation, "���̃A�v���P�[�V�����ɂ���"
End Sub

' �A�v���P�[�V�����I�����̃N���[���A�b�v
Public Sub CleanupApplication()
    On Error Resume Next
    
    ' �i���\�����N���A
    Call ClearProgress()
    
    ' �A�v���P�[�V�����ݒ�𕜌�
    Call RestoreApplicationSettings()
    
    ' ���O���b�Z�[�W
    Call LogMessage(LOG_INFO, "�A�v���P�[�V�����N���[���A�b�v����")
End Sub

' �������s�p�T���v��
Public Sub AutoCollectSample()
    On Error GoTo ErrorHandler
    
    Dim stocks As String
    Dim timeFrame As String
    Dim startDate As Date
    Dim endDate As Date
    Dim result As Boolean
    
    ' �T���v���ݒ�
    stocks = "7203,6758,9984"    ' �g���^�A�\�j�[�A�\�t�g�o���N
    timeFrame = "5M"             ' 5����
    startDate = Date - 7         ' 1�T�ԑO
    endDate = Date               ' ����
    
    ' �m�F���b�Z�[�W
    If MsgBox("�ȉ��̐ݒ�Ŏ����f�[�^�擾�����s���܂����H" & vbCrLf & vbCrLf & _
              "����: " & stocks & vbCrLf & _
              "����: " & timeFrame & vbCrLf & _
              "����: " & Format(startDate, "YYYY/MM/DD") & " - " & Format(endDate, "YYYY/MM/DD"), _
              vbYesNo + vbQuestion, "�������s�m�F") = vbNo Then
        Exit Sub
    End If
    
    Call LogMessage(LOG_INFO, "�����f�[�^�擾�J�n")
    
    ' �o�b�`�������s
    result = CollectMultipleStocks(stocks, timeFrame, startDate, endDate)
    
    If result Then
        MsgBox "�����f�[�^�擾���������܂����B", vbInformation, "����"
        Call LogMessage(LOG_INFO, "�����f�[�^�擾����")
    Else
        MsgBox "�����f�[�^�擾�ňꕔ�G���[���������܂����B���O���m�F���Ă��������B", vbExclamation, "�ꕔ�G���["
        Call LogMessage(LOG_WARN, "�����f�[�^�擾�ňꕔ�G���[")
    End If
    
    Exit Sub
    
ErrorHandler:
    Call LogDetailedError("AutoCollectSample", Err.Description)
    MsgBox "�������s�ŃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
End Sub

' Excel�u�b�N�I�����̏���
Public Sub Workbook_BeforeClose()
    Call CleanupApplication()
End Sub