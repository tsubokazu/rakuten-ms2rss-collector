Attribute VB_Name = "CSVExporter"
'******************************************************************************
' �y�VMS2RSS�����f�[�^�R���N�^�[ - CSV�o�̓��W���[��
' 
' ����: �����f�[�^��CSV�`���o�͋@�\
' �쐬��: Claude Code
' �o�[�W����: 1.0.0
'******************************************************************************

Option Explicit

' CSV�o�͐ݒ�
Private Type CSVConfig
    IncludeHeader As Boolean
    DateTimeFormat As String
    DecimalPlaces As Integer
    Encoding As String
    Delimiter As String
End Type

Private config As CSVConfig

' ���W���[��������
Private Sub InitializeCSVConfig()
    config.IncludeHeader = True
    config.DateTimeFormat = "YYYY-MM-DD HH:MM:SS"
    config.DecimalPlaces = 2
    config.Encoding = "UTF-8"
    config.Delimiter = ","
End Sub

' �����f�[�^��CSV�`���ŏo��
Public Function ExportStockDataToCSV(stockData As Variant, filePath As String, _
                                    Optional includeHeader As Boolean = True) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim csvContent As String
    Dim fileNum As Integer
    
    Call InitializeCSVConfig
    config.IncludeHeader = includeHeader
    
    ' CSV���e�𐶐�
    csvContent = GenerateCSVContent(stockData)
    
    If csvContent = "" Then
        Call LogMessage("ERROR", "CSV���e�̐����Ɏ��s���܂���")
        ExportStockDataToCSV = False
        Exit Function
    End If
    
    ' �t�@�C���ɏ�������
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, csvContent;
    Close #fileNum
    
    Call LogMessage("INFO", "CSV�t�@�C���o�͊���: " & filePath)
    ExportStockDataToCSV = True
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage("ERROR", "ExportStockDataToCSV: " & Err.Description)
    ExportStockDataToCSV = False
End Function

' CSV���e�𐶐�
Private Function GenerateCSVContent(stockData As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim content As String
    Dim i As Long, j As Long
    Dim rowContent As String
    Dim cellValue As String
    
    ' �w�b�_�[�s��ǉ�
    If config.IncludeHeader Then
        content = GenerateCSVHeader() & vbCrLf
    End If
    
    ' �f�[�^�s������
    If IsArray(stockData) Then
        For i = LBound(stockData, 1) To UBound(stockData, 1)
            rowContent = ""
            For j = LBound(stockData, 2) To UBound(stockData, 2)
                If j > LBound(stockData, 2) Then
                    rowContent = rowContent & config.Delimiter
                End If
                
                cellValue = FormatCellValue(stockData(i, j), j)
                rowContent = rowContent & cellValue
            Next j
            content = content & rowContent & vbCrLf
        Next i
    End If
    
    GenerateCSVContent = content
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "GenerateCSVContent: " & Err.Description)
    GenerateCSVContent = ""
End Function

' CSV�w�b�_�[�𐶐�
Private Function GenerateCSVHeader() As String
    Dim header As String
    
    header = "DateTime" & config.Delimiter & _
             "Open" & config.Delimiter & _
             "High" & config.Delimiter & _
             "Low" & config.Delimiter & _
             "Close" & config.Delimiter & _
             "Volume"
    
    GenerateCSVHeader = header
End Function

' �Z���l���t�H�[�}�b�g
Private Function FormatCellValue(value As Variant, columnIndex As Integer) As String
    On Error GoTo ErrorHandler
    
    Dim formattedValue As String
    
    Select Case columnIndex
        Case 0 ' DateTime��
            If IsDate(value) Then
                formattedValue = Format(value, "YYYY-MM-DD HH:MM:SS")
            Else
                formattedValue = CStr(value)
            End If
            
        Case 1 To 4 ' OHLC��i���l�j
            If IsNumeric(value) Then
                formattedValue = Format(value, "0." & String(config.DecimalPlaces, "0"))
            Else
                formattedValue = CStr(value)
            End If
            
        Case 5 ' Volume��i�����j
            If IsNumeric(value) Then
                formattedValue = Format(value, "0")
            Else
                formattedValue = CStr(value)
            End If
            
        Case Else
            formattedValue = CStr(value)
    End Select
    
    ' CSV�G�X�P�[�v����
    If InStr(formattedValue, config.Delimiter) > 0 Or _
       InStr(formattedValue, """") > 0 Or _
       InStr(formattedValue, vbCrLf) > 0 Then
        formattedValue = """" & Replace(formattedValue, """", """""") & """"
    End If
    
    FormatCellValue = formattedValue
    Exit Function
    
ErrorHandler:
    FormatCellValue = CStr(value)
End Function

' �o�b�`CSV�o�́i���������Ή��j
Public Function ExportMultipleStocksToCSV(stockDataArray As Variant, _
                                         baseFilePath As String, _
                                         stockCodes As Variant) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim filePath As String
    Dim fileName As String
    Dim fileExtension As String
    Dim baseName As String
    Dim successCount As Long
    
    If Not IsArray(stockDataArray) Or Not IsArray(stockCodes) Then
        Call LogMessage("ERROR", "���̓f�[�^���z��ł͂���܂���")
        ExportMultipleStocksToCSV = False
        Exit Function
    End If
    
    ' �t�@�C���p�X�𕪉�
    fileExtension = ".csv"
    If InStr(baseFilePath, ".") > 0 Then
        baseName = Left(baseFilePath, InStrRev(baseFilePath, ".") - 1)
        fileExtension = Right(baseFilePath, Len(baseFilePath) - InStrRev(baseFilePath, ".") + 1)
    Else
        baseName = baseFilePath
    End If
    
    ' �e�����̃f�[�^���o��
    For i = LBound(stockCodes) To UBound(stockCodes)
        fileName = baseName & "_" & Replace(stockCodes(i), ".", "_") & fileExtension
        
        If ExportStockDataToCSV(stockDataArray(i), fileName) Then
            successCount = successCount + 1
        End If
    Next i
    
    Call LogMessage("INFO", "�o�b�`CSV�o�͊���: " & successCount & "/" & (UBound(stockCodes) - LBound(stockCodes) + 1))
    ExportMultipleStocksToCSV = (successCount = (UBound(stockCodes) - LBound(stockCodes) + 1))
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "ExportMultipleStocksToCSV: " & Err.Description)
    ExportMultipleStocksToCSV = False
End Function

' CSV�ݒ��ύX
Public Sub SetCSVConfig(Optional includeHeader As Boolean = True, _
                       Optional dateTimeFormat As String = "YYYY-MM-DD HH:MM:SS", _
                       Optional decimalPlaces As Integer = 2, _
                       Optional delimiter As String = ",")
    
    config.IncludeHeader = includeHeader
    config.DateTimeFormat = dateTimeFormat
    config.DecimalPlaces = decimalPlaces
    config.Delimiter = delimiter
End Sub

' �t�@�C���T�C�Y���擾�iMB�P�ʁj
Public Function GetFileSize(filePath As String) As Double
    On Error GoTo ErrorHandler
    
    Dim fileSize As Long
    fileSize = FileLen(filePath)
    GetFileSize = fileSize / 1024 / 1024 ' MB���Z
    
    Exit Function
    
ErrorHandler:
    GetFileSize = 0
End Function

' CSV�t�@�C���̑Ó����`�F�b�N
Public Function ValidateCSVFile(filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim firstLine As String
    Dim expectedHeader As String
    
    If Dir(filePath) = "" Then
        Call LogMessage("ERROR", "�t�@�C�������݂��܂���: " & filePath)
        ValidateCSVFile = False
        Exit Function
    End If
    
    ' �t�@�C���T�C�Y�`�F�b�N�i��t�@�C���łȂ����j
    If GetFileSize(filePath) = 0 Then
        Call LogMessage("ERROR", "�t�@�C������ł�: " & filePath)
        ValidateCSVFile = False
        Exit Function
    End If
    
    ' �w�b�_�[�s���`�F�b�N
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Line Input #fileNum, firstLine
    Close #fileNum
    
    expectedHeader = GenerateCSVHeader()
    If firstLine <> expectedHeader Then
        Call LogMessage("WARN", "�w�b�_�[�s�����Ғl�ƈقȂ�܂�: " & filePath)
    End If
    
    ValidateCSVFile = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage("ERROR", "ValidateCSVFile: " & Err.Description)
    ValidateCSVFile = False
End Function