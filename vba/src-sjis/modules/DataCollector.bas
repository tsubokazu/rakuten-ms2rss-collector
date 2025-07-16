Attribute VB_Name = "DataCollector"
'******************************************************************************
' �y�VMS2RSS�����f�[�^�R���N�^�[ - �f�[�^�擾���W���[��
' 
' ����: �y�V�،�MarketSpeed2��RSS API�o�R�Ŋ����f�[�^���擾
' �쐬��: Claude Code
' �o�[�W����: 1.0.0
'******************************************************************************

Option Explicit

' �f�[�^�擾�̎�֐�
Public Function CollectStockData(stockCode As String, timeFrame As String, _
                                startDate As Date, endDate As Date, _
                                Optional outputPath As String = "") As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim totalBars As Long
    Dim currentDate As Date
    Dim dataArray As Variant
    Dim csvData As String
    
    ' ���O�o��
    Call LogMessage("INFO", "�f�[�^�擾�J�n: " & stockCode & " (" & timeFrame & ")")
    
    ' ���Ԃ̑Ó����`�F�b�N
    If startDate > endDate Then
        Call LogMessage("ERROR", "�J�n�����I��������ł�")
        CollectStockData = False
        Exit Function
    End If
    
    ' �����R�[�h�̌`���`�F�b�N
    If Not ValidateStockCode(stockCode) Then
        Call LogMessage("ERROR", "�����Ȗ����R�[�h: " & stockCode)
        CollectStockData = False
        Exit Function
    End If
    
    ' RSS Chart API�Ńf�[�^�擾
    dataArray = GetChartDataFromAPI(stockCode, timeFrame, startDate, endDate)
    
    If IsEmpty(dataArray) Then
        Call LogMessage("ERROR", "�f�[�^�擾�Ɏ��s���܂���: " & stockCode)
        CollectStockData = False
        Exit Function
    End If
    
    ' CSV�f�[�^�ɕϊ�
    csvData = ConvertToCSV(dataArray)
    
    ' �t�@�C���o��
    If outputPath = "" Then
        outputPath = GenerateOutputFilename(stockCode, timeFrame, startDate, endDate)
    End If
    
    result = SaveCSVFile(csvData, outputPath)
    
    If result Then
        Call LogMessage("INFO", "�f�[�^�擾����: " & outputPath)
        CollectStockData = True
    Else
        Call LogMessage("ERROR", "�t�@�C���ۑ��Ɏ��s���܂���: " & outputPath)
        CollectStockData = False
    End If
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "DataCollector.CollectStockData: " & Err.Description)
    CollectStockData = False
End Function

' RSS Chart API����f�[�^���擾
Private Function GetChartDataFromAPI(stockCode As String, timeFrame As String, _
                                   startDate As Date, endDate As Date) As Variant
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim headerRange As Range
    Dim result As Variant
    Dim tempSheetName As String
    Dim maxBars As Long
    
    ' �ꎞ���[�N�V�[�g���쐬
    tempSheetName = "TempData_" & Format(Now, "hhmmss")
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = tempSheetName
    
    ' �w�b�_�[�s��ݒ�
    ws.Range("A1:F1").Value = Array("DateTime", "Open", "High", "Low", "Close", "Volume")
    Set headerRange = ws.Range("A1:F1")
    
    ' API�������l�������ő�擾�{��
    maxBars = 3000
    
    ' RSS Chart API�Ăяo���iVBA�Łj
    ' �ߋ��f�[�^�擾�̏ꍇ��RssChartPast_v���g�p
    If startDate < Date Then
        result = Application.WorksheetFunction.RssChartPast_v( _
            headerRange, stockCode, timeFrame, Format(startDate, "YYYYMMDD"), maxBars)
    Else
        result = Application.WorksheetFunction.RssChart_v( _
            headerRange, stockCode, timeFrame, maxBars)
    End If
    
    ' ���ʂ��R�s�[
    GetChartDataFromAPI = result
    
    ' �ꎞ���[�N�V�[�g���폜
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "GetChartDataFromAPI: " & Err.Description)
    
    ' �G���[�����ꎞ���[�N�V�[�g���폜
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    
    GetChartDataFromAPI = Empty
End Function

' �����R�[�h�̑Ó����`�F�b�N
Private Function ValidateStockCode(stockCode As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim codePattern As String
    Dim marketPart As String
    Dim codePart As String
    
    ' �����R�[�h.�s�� �̌`�����`�F�b�N
    If InStr(stockCode, ".") > 0 Then
        codePart = Split(stockCode, ".")(0)
        marketPart = Split(stockCode, ".")(1)
        
        ' �s��R�[�h�̑Ó����`�F�b�N
        Select Case UCase(marketPart)
            Case "T", "JAX", "JNX", "CHJ"
                ' �L���Ȏs��R�[�h
            Case Else
                ValidateStockCode = False
                Exit Function
        End Select
    Else
        codePart = stockCode
    End If
    
    ' ���l�����̃`�F�b�N�i4���܂���5���j
    If Len(codePart) >= 4 And Len(codePart) <= 5 And IsNumeric(codePart) Then
        ValidateStockCode = True
    Else
        ValidateStockCode = False
    End If
    
    Exit Function
    
ErrorHandler:
    ValidateStockCode = False
End Function

' �z��f�[�^��CSV�`���ɕϊ�
Private Function ConvertToCSV(dataArray As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim csvString As String
    Dim i As Long, j As Long
    Dim rowData As String
    
    ' �w�b�_�[�s
    csvString = "DateTime,Open,High,Low,Close,Volume" & vbCrLf
    
    ' �f�[�^�s
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        rowData = ""
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            If j > LBound(dataArray, 2) Then rowData = rowData & ","
            rowData = rowData & CStr(dataArray(i, j))
        Next j
        csvString = csvString & rowData & vbCrLf
    Next i
    
    ConvertToCSV = csvString
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "ConvertToCSV: " & Err.Description)
    ConvertToCSV = ""
End Function

' �o�̓t�@�C�����𐶐�
Private Function GenerateOutputFilename(stockCode As String, timeFrame As String, _
                                      startDate As Date, endDate As Date) As String
    Dim fileName As String
    Dim outputDir As String
    
    ' �o�̓f�B���N�g��
    outputDir = ThisWorkbook.Path & "\output\csv\"
    
    ' �f�B���N�g�������݂��Ȃ��ꍇ�͍쐬
    If Dir(outputDir, vbDirectory) = "" Then
        MkDir outputDir
    End If
    
    ' �t�@�C��������
    fileName = Replace(stockCode, ".", "_") & "_" & timeFrame & "_" & _
               Format(startDate, "YYYYMMDD") & "-" & Format(endDate, "YYYYMMDD") & ".csv"
    
    GenerateOutputFilename = outputDir & fileName
End Function

' CSV�t�@�C����ۑ�
Private Function SaveCSVFile(csvData As String, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, csvData;
    Close #fileNum
    
    SaveCSVFile = True
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage("ERROR", "SaveCSVFile: " & Err.Description)
    SaveCSVFile = False
End Function

' ���������̈ꊇ����
Public Function CollectMultipleStocks(stockCodes As String, timeFrame As String, _
                                    startDate As Date, endDate As Date) As Boolean
    On Error GoTo ErrorHandler
    
    Dim stocks() As String
    Dim i As Long
    Dim successCount As Long
    Dim totalCount As Long
    
    ' �����R�[�h�𕪊�
    stocks = Split(Replace(stockCodes, " ", ""), ",")
    totalCount = UBound(stocks) + 1
    
    Call LogMessage("INFO", "���������f�[�^�擾�J�n: " & totalCount & "����")
    
    ' �e����������
    For i = 0 To UBound(stocks)
        If Trim(stocks(i)) <> "" Then
            If CollectStockData(Trim(stocks(i)), timeFrame, startDate, endDate) Then
                successCount = successCount + 1
            End If
            
            ' �i���X�V�i��Ńt�H�[������Ăяo���\�ɂ���j
            DoEvents
        End If
    Next i
    
    Call LogMessage("INFO", "���������f�[�^�擾����: " & successCount & "/" & totalCount)
    CollectMultipleStocks = (successCount = totalCount)
    
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "CollectMultipleStocks: " & Err.Description)
    CollectMultipleStocks = False
End Function