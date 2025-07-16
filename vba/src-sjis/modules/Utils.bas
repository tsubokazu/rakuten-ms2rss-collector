Attribute VB_Name = "Utils"
'******************************************************************************
' �y�VMS2RSS�����f�[�^�R���N�^�[ - ���[�e�B���e�B���W���[��
' 
' ����: ���ʃ��[�e�B���e�B�֐��E���O�@�\
' �쐬��: Claude Code
' �o�[�W����: 1.0.0
'******************************************************************************

Option Explicit

' ���O���x���萔
Public Const LOG_DEBUG As String = "DEBUG"
Public Const LOG_INFO As String = "INFO"
Public Const LOG_WARN As String = "WARN"
Public Const LOG_ERROR As String = "ERROR"

' ���O�ݒ�
Private Type LogConfig
    Enabled As Boolean
    Level As String
    MaxFiles As Integer
    MaxFileSizeMB As Integer
    LogDirectory As String
End Type

Private logConfig As LogConfig
Private logInitialized As Boolean

' ���O�@�\������
Private Sub InitializeLogging()
    If logInitialized Then Exit Sub
    
    logConfig.Enabled = True
    logConfig.Level = LOG_INFO
    logConfig.MaxFiles = 10
    logConfig.MaxFileSizeMB = 10
    logConfig.LogDirectory = ThisWorkbook.Path & "\output\logs\"
    
    ' ���O�f�B���N�g���쐬
    If Dir(logConfig.LogDirectory, vbDirectory) = "" Then
        MkDir logConfig.LogDirectory
    End If
    
    logInitialized = True
End Sub

' ���O���b�Z�[�W�o��
Public Sub LogMessage(level As String, message As String)
    On Error GoTo ErrorHandler
    
    Call InitializeLogging
    
    If Not logConfig.Enabled Then Exit Sub
    
    Dim logFilePath As String
    Dim fileNum As Integer
    Dim timestamp As String
    Dim logLine As String
    
    ' �^�C���X�^���v����
    timestamp = Format(Now, "YYYY-MM-DD HH:MM:SS")
    
    ' ���O�s�쐬
    logLine = timestamp & " [" & level & "] " & message
    
    ' �R���\�[���o�́i�C�~�f�B�G�C�g�E�B���h�E�j
    Debug.Print logLine
    
    ' �t�@�C���o��
    logFilePath = GetCurrentLogFilePath()
    
    ' �t�@�C���T�C�Y�`�F�b�N
    If Dir(logFilePath) <> "" And FileLen(logFilePath) > logConfig.MaxFileSizeMB * 1024 * 1024 Then
        Call RotateLogFiles()
        logFilePath = GetCurrentLogFilePath()
    End If
    
    ' ���O�t�@�C���ɏ�������
    fileNum = FreeFile
    Open logFilePath For Append As #fileNum
    Print #fileNum, logLine
    Close #fileNum
    
    Exit Sub
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Debug.Print "LogMessage Error: " & Err.Description
End Sub

' ���݂̃��O�t�@�C���p�X���擾
Private Function GetCurrentLogFilePath() As String
    Dim fileName As String
    fileName = "ms2rss_collector_" & Format(Date, "YYYYMMDD") & ".log"
    GetCurrentLogFilePath = logConfig.LogDirectory & fileName
End Function

' ���O�t�@�C�����[�e�[�V����
Private Sub RotateLogFiles()
    On Error Resume Next
    
    Dim i As Integer
    Dim oldFile As String
    Dim newFile As String
    Dim baseFileName As String
    
    baseFileName = "ms2rss_collector_" & Format(Date, "YYYYMMDD")
    
    ' �Â��t�@�C�����폜
    For i = logConfig.MaxFiles To 1 Step -1
        oldFile = logConfig.LogDirectory & baseFileName & "_" & Format(i, "00") & ".log"
        If Dir(oldFile) <> "" Then
            If i = logConfig.MaxFiles Then
                Kill oldFile
            Else
                newFile = logConfig.LogDirectory & baseFileName & "_" & Format(i + 1, "00") & ".log"
                Name oldFile As newFile
            End If
        End If
    Next i
    
    ' ���݂̃t�@�C����_01�Ƀ��l�[��
    oldFile = GetCurrentLogFilePath()
    If Dir(oldFile) <> "" Then
        newFile = logConfig.LogDirectory & baseFileName & "_01.log"
        Name oldFile As newFile
    End If
End Sub

' ���ݎ����̕�����擾
Public Function GetCurrentTimestamp() As String
    GetCurrentTimestamp = Format(Now, "YYYY-MM-DD HH:MM:SS")
End Function

' �f�B���N�g�����݃`�F�b�N�E�쐬
Public Function EnsureDirectoryExists(dirPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Dir(dirPath, vbDirectory) = "" Then
        MkDir dirPath
        Call LogMessage(LOG_INFO, "�f�B���N�g�����쐬���܂���: " & dirPath)
    End If
    
    EnsureDirectoryExists = True
    Exit Function
    
ErrorHandler:
    Call LogMessage(LOG_ERROR, "�f�B���N�g���쐬�G���[: " & dirPath & " - " & Err.Description)
    EnsureDirectoryExists = False
End Function

' �t�@�C�����݃`�F�b�N
Public Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

' ���S�ȃt�@�C���������i���������������j
Public Function SafeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim char As String
    Dim result As String
    
    invalidChars = "\/:*?""<>|"
    result = fileName
    
    For i = 1 To Len(invalidChars)
        char = Mid(invalidChars, i, 1)
        result = Replace(result, char, "_")
    Next i
    
    SafeFileName = result
End Function

' �i���\���i0-100%�j
Public Sub ShowProgress(currentValue As Long, maxValue As Long, Optional message As String = "")
    Dim percentage As Integer
    Dim progressBar As String
    Dim i As Integer
    
    If maxValue = 0 Then Exit Sub
    
    percentage = Int((currentValue / maxValue) * 100)
    
    ' �v���O���X�o�[������쐬
    progressBar = "["
    For i = 1 To 50
        If i <= Int(percentage / 2) Then
            progressBar = progressBar & "="
        Else
            progressBar = progressBar & " "
        End If
    Next i
    progressBar = progressBar & "] " & percentage & "%"
    
    If message <> "" Then
        progressBar = message & " " & progressBar
    End If
    
    ' �X�e�[�^�X�o�[�ɕ\��
    Application.StatusBar = progressBar
    
    ' ���O�ɂ��o�́i10%���݁j
    If percentage Mod 10 = 0 And currentValue > 0 Then
        Call LogMessage(LOG_INFO, "�i��: " & percentage & "% (" & currentValue & "/" & maxValue & ")")
    End If
    
    DoEvents
End Sub

' �i���\���N���A
Public Sub ClearProgress()
    Application.StatusBar = False
End Sub

' �ݒ�l�̌���
Public Function ValidateTimeFrame(timeFrame As String) As Boolean
    Dim validFrames As Variant
    Dim i As Integer
    
    validFrames = Array("T", "1M", "2M", "3M", "4M", "5M", "10M", "15M", "30M", "60M", _
                       "2H", "4H", "8H", "D", "W", "M")
    
    For i = 0 To UBound(validFrames)
        If UCase(timeFrame) = UCase(validFrames(i)) Then
            ValidateTimeFrame = True
            Exit Function
        End If
    Next i
    
    ValidateTimeFrame = False
End Function

' ���t�͈͂̌���
Public Function ValidateDateRange(startDate As Date, endDate As Date) As Boolean
    If startDate > endDate Then
        Call LogMessage(LOG_ERROR, "�J�n�����I��������ł�")
        ValidateDateRange = False
        Exit Function
    End If
    
    If startDate > Date Then
        Call LogMessage(LOG_WARN, "�J�n�����������ł�")
    End If
    
    If DateDiff("d", startDate, endDate) > 365 Then
        Call LogMessage(LOG_WARN, "�擾���Ԃ�1�N�𒴂��Ă��܂�")
    End If
    
    ValidateDateRange = True
End Function

' �������g�p�ʃ`�F�b�N�i�ȈՔŁj
Public Function GetMemoryUsage() As String
    On Error Resume Next
    
    Dim memInfo As String
    memInfo = "Excel: " & Format(Application.MemoryUsed / 1024, "0.0") & "MB"
    
    GetMemoryUsage = memInfo
End Function

' �A�v���P�[�V�����ݒ�̕���
Public Sub RestoreApplicationSettings()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub

' �A�v���P�[�V�����ݒ�̍œK��
Public Sub OptimizeApplicationSettings()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

' �o�[�W�������擾
Public Function GetApplicationVersion() As String
    GetApplicationVersion = "�y�VMS2RSS�����f�[�^�R���N�^�[ v1.0.0"
End Function

' �V�X�e�����擾
Public Function GetSystemInfo() As String
    Dim info As String
    
    info = "Excel: " & Application.Version & vbCrLf
    info = info & "OS: " & Application.OperatingSystem & vbCrLf
    info = info & "���[�U�[: " & Application.UserName & vbCrLf
    info = info & "���s����: " & GetCurrentTimestamp()
    
    GetSystemInfo = info
End Function

' �G���[���̏ڍ׃��O
Public Sub LogDetailedError(functionName As String, errorDescription As String, _
                          Optional additionalInfo As String = "")
    
    Dim errorMessage As String
    
    errorMessage = "�֐�: " & functionName & vbCrLf
    errorMessage = errorMessage & "�G���[: " & errorDescription & vbCrLf
    
    If additionalInfo <> "" Then
        errorMessage = errorMessage & "�ǉ����: " & additionalInfo & vbCrLf
    End If
    
    errorMessage = errorMessage & "����: " & GetCurrentTimestamp() & vbCrLf
    errorMessage = errorMessage & GetMemoryUsage()
    
    Call LogMessage(LOG_ERROR, errorMessage)
End Sub