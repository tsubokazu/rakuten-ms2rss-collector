Attribute VB_Name = "WorksheetMacros"
'******************************************************************************
' �y�VMS2RSS�����f�[�^�R���N�^�[ - ���[�N�V�[�g�}�N��
' 
' ����: Excel���[�N�V�[�g��̃{�^������Ăяo�����}�N��
' �쐬��: Claude Code
' �o�[�W����: 1.0.0
'******************************************************************************

Option Explicit

' ���C���t�H�[���\���{�^���p�}�N��
Public Sub StartDataCollection()
    Call ShowMainForm
End Sub

' �N�C�b�N�e�X�g�{�^���p�}�N��
Public Sub RunQuickTest()
    Call QuickTest
End Sub

' �ݒ�\���{�^���p�}�N��
Public Sub DisplaySettings()
    Call ShowCurrentConfig
End Sub

' ���O�\���{�^���p�}�N��
Public Sub ViewLogs()
    Call ShowLogViewer
End Sub

' �o�[�W�������{�^���p�}�N��
Public Sub AboutApp()
    Call ShowAbout
End Sub

' �������s�T���v���{�^���p�}�N��
Public Sub RunAutoSample()
    Call AutoCollectSample
End Sub

' �o�̓t�H���_���J��
Public Sub OpenOutputFolder()
    On Error GoTo ErrorHandler
    
    Dim outputPath As String
    outputPath = ThisWorkbook.Path & "\output\csv\"
    
    ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬
    If Not EnsureDirectoryExists(outputPath) Then
        MsgBox "�o�̓t�H���_�̍쐬�Ɏ��s���܂���: " & outputPath, vbCritical
        Exit Sub
    End If
    
    ' �t�H���_���J��
    Shell "explorer.exe " & Chr(34) & outputPath & Chr(34), vbNormalFocus
    Call LogMessage(LOG_INFO, "�o�̓t�H���_���J���܂���: " & outputPath)
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "OpenOutputFolder: " & Err.Description)
    MsgBox "�t�H���_���J�����Ƃ��ł��܂���ł���: " & Err.Description, vbCritical
End Sub

' ���O�t�H���_���J��
Public Sub OpenLogFolder()
    On Error GoTo ErrorHandler
    
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\output\logs\"
    
    ' �t�H���_�����݂��Ȃ��ꍇ�͍쐬
    If Not EnsureDirectoryExists(logPath) Then
        MsgBox "���O�t�H���_�̍쐬�Ɏ��s���܂���: " & logPath, vbCritical
        Exit Sub
    End If
    
    ' �t�H���_���J��
    Shell "explorer.exe " & Chr(34) & logPath & Chr(34), vbNormalFocus
    Call LogMessage(LOG_INFO, "���O�t�H���_���J���܂���: " & logPath)
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "OpenLogFolder: " & Err.Description)
    MsgBox "�t�H���_���J�����Ƃ��ł��܂���ł���: " & Err.Description, vbCritical
End Sub

' MarketSpeed2�ڑ��e�X�g
Public Sub TestConnection()
    On Error GoTo ErrorHandler
    
    Dim testResult As Variant
    
    ' ���o���ό��ݒl�Őڑ��e�X�g
    testResult = Application.WorksheetFunction.RssIndexMarket("0000", "���ݒl")
    
    If IsError(testResult) Then
        MsgBox "MarketSpeed2�ւ̐ڑ��Ɏ��s���܂����B" & vbCrLf & vbCrLf & _
               "�ȉ����m�F���Ă��������F" & vbCrLf & _
               "1. MarketSpeed2���N�����Ă���" & vbCrLf & _
               "2. RSS�@�\���L���ɂȂ��Ă���" & vbCrLf & _
               "3. ���O�C����Ԃ�����ł���", vbExclamation, "�ڑ��e�X�g����"
        Call LogMessage(LOG_ERROR, "MS2�ڑ��e�X�g���s")
    Else
        MsgBox "MarketSpeed2�ւ̐ڑ��ɐ������܂����I" & vbCrLf & vbCrLf & _
               "���o���ό��ݒl: " & testResult, vbInformation, "�ڑ��e�X�g����"
        Call LogMessage(LOG_INFO, "MS2�ڑ��e�X�g����: " & testResult)
    End If
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "TestConnection: " & Err.Description)
    MsgBox "�ڑ��e�X�g�ŃG���[���������܂���: " & Err.Description, vbCritical, "�ڑ��e�X�g�G���["
End Sub

' �ݒ�t�@�C���̏�����
Public Sub InitializeSettings()
    On Error GoTo ErrorHandler
    
    Dim config As Configuration
    Set config = New Configuration
    
    ' �f�t�H���g�ݒ�Ńt�@�C���쐬
    If config.SaveToFile() Then
        MsgBox "�ݒ�t�@�C�������������܂����B" & vbCrLf & _
               "�t�@�C��: " & ThisWorkbook.Path & "\config\settings.json", vbInformation, "�ݒ菉����"
        Call LogMessage(LOG_INFO, "�ݒ�t�@�C������������")
    Else
        MsgBox "�ݒ�t�@�C���̏������Ɏ��s���܂����B", vbCritical, "�ݒ菉�����G���["
    End If
    
    Set config = Nothing
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "InitializeSettings: " & Err.Description)
    MsgBox "�ݒ菉�����ŃG���[���������܂���: " & Err.Description, vbCritical, "�������G���["
End Sub

' �V�X�e�����\��
Public Sub ShowSystemInfo()
    Dim info As String
    
    info = GetSystemInfo() & vbCrLf & vbCrLf
    info = info & GetApplicationVersion() & vbCrLf
    info = info & "�������g�p��: " & GetMemoryUsage()
    
    MsgBox info, vbInformation, "�V�X�e�����"
End Sub

' �w���v�\��
Public Sub ShowHelp()
    Dim helpMessage As String
    
    helpMessage = "�y�y�VMS2RSS�����f�[�^�R���N�^�[ �w���v�z" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "�� ��{�I�Ȏg����" & vbCrLf
    helpMessage = helpMessage & "1. �u�f�[�^���W�J�n�v�{�^�����N���b�N" & vbCrLf
    helpMessage = helpMessage & "2. �����R�[�h�A���ԁA�����ݒ�" & vbCrLf
    helpMessage = helpMessage & "3. �o�͐�t�H���_���w��" & vbCrLf
    helpMessage = helpMessage & "4. �u���s�v�{�^���Ńf�[�^�擾�J�n" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "�� �����R�[�h�`��" & vbCrLf
    helpMessage = helpMessage & "? �P�����: 7203" & vbCrLf
    helpMessage = helpMessage & "? ��������: 7203,6758,9984" & vbCrLf
    helpMessage = helpMessage & "? �s��w��: 7203.T, 7203.JAX" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "�� �Ή�����" & vbCrLf
    helpMessage = helpMessage & "1M, 5M, 15M, 30M, 60M, D�i�����j" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "�� ���ӎ���" & vbCrLf
    helpMessage = helpMessage & "? MarketSpeed2���N�����Ă���K�v������܂�" & vbCrLf
    helpMessage = helpMessage & "? RSS�@�\��L���ɂ��Ă�������" & vbCrLf
    helpMessage = helpMessage & "? ��ʃf�[�^�擾���͎��Ԃ�������܂�" & vbCrLf & vbCrLf
    helpMessage = helpMessage & "�ڍׂ� docs/vba-guide.md ���Q�Ƃ��Ă��������B"
    
    MsgBox helpMessage, vbInformation, "�w���v"
End Sub

' �T���v���f�[�^�쐬�i�e�X�g�p�j
Public Sub CreateSampleData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sampleData(1 To 10, 1 To 6) As Variant
    Dim i As Long
    
    ' �V�������[�N�V�[�g���쐬
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "�T���v���f�[�^_" & Format(Now, "HHMMSS")
    
    ' �w�b�_�[�ݒ�
    ws.Range("A1:F1").Value = Array("DateTime", "Open", "High", "Low", "Close", "Volume")
    
    ' �T���v���f�[�^����
    For i = 1 To 10
        sampleData(i, 1) = Format(Now - (10 - i) / 24 / 60 * 5, "YYYY-MM-DD HH:MM:SS") ' 5���Ԋu
        sampleData(i, 2) = 2500 + Rnd() * 100 ' Open
        sampleData(i, 3) = sampleData(i, 2) + Rnd() * 50 ' High
        sampleData(i, 4) = sampleData(i, 2) - Rnd() * 50 ' Low
        sampleData(i, 5) = sampleData(i, 2) + (Rnd() - 0.5) * 30 ' Close
        sampleData(i, 6) = Int(Rnd() * 100000) + 50000 ' Volume
    Next i
    
    ' �f�[�^�����[�N�V�[�g�ɐݒ�
    ws.Range("A2:F11").Value = sampleData
    
    ' �����ݒ�
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("B2:E11").NumberFormat = "0.00"
    ws.Range("F2:F11").NumberFormat = "#,##0"
    ws.Columns.AutoFit
    
    MsgBox "�T���v���f�[�^���쐬���܂���: " & ws.Name, vbInformation, "�T���v���f�[�^�쐬"
    Call LogMessage(LOG_INFO, "�T���v���f�[�^�쐬: " & ws.Name)
    
    Exit Sub
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "CreateSampleData: " & Err.Description)
    MsgBox "�T���v���f�[�^�쐬�ŃG���[���������܂���: " & Err.Description, vbCritical, "�G���["
End Sub

' �S���[�N�V�[�g�}�N���̃��X�g�\��
Public Sub ShowMacroList()
    Dim macroList As String
    
    macroList = "�y���p�\�ȃ}�N���ꗗ�z" & vbCrLf & vbCrLf
    macroList = macroList & "�� �f�[�^����" & vbCrLf
    macroList = macroList & "? StartDataCollection - �f�[�^���W�J�n" & vbCrLf
    macroList = macroList & "? RunQuickTest - �N�C�b�N�e�X�g���s" & vbCrLf
    macroList = macroList & "? RunAutoSample - �������s�T���v��" & vbCrLf & vbCrLf
    macroList = macroList & "�� �ݒ�E���" & vbCrLf
    macroList = macroList & "? DisplaySettings - �ݒ�\��" & vbCrLf
    macroList = macroList & "? InitializeSettings - �ݒ菉����" & vbCrLf
    macroList = macroList & "? ShowSystemInfo - �V�X�e�����" & vbCrLf & vbCrLf
    macroList = macroList & "�� ���O�E�f�o�b�O" & vbCrLf
    macroList = macroList & "? ViewLogs - ���O�\��" & vbCrLf
    macroList = macroList & "? TestConnection - �ڑ��e�X�g" & vbCrLf
    macroList = macroList & "? CreateSampleData - �T���v���f�[�^�쐬" & vbCrLf & vbCrLf
    macroList = macroList & "�� ���[�e�B���e�B" & vbCrLf
    macroList = macroList & "? OpenOutputFolder - �o�̓t�H���_���J��" & vbCrLf
    macroList = macroList & "? OpenLogFolder - ���O�t�H���_���J��" & vbCrLf
    macroList = macroList & "? AboutApp - �o�[�W�������" & vbCrLf
    macroList = macroList & "? ShowHelp - �w���v�\��"
    
    MsgBox macroList, vbInformation, "�}�N���ꗗ"
End Sub