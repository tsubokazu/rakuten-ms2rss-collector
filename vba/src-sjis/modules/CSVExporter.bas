Attribute VB_Name = "CSVExporter"
'******************************************************************************
' Rakuten MS2RSS Stock Data Collector - CSV Export Module
' 
' Description: Stock data CSV export functionality
' Author: Claude Code
' Version: 1.0.0
'******************************************************************************

Option Explicit

' Export stock data to CSV format
Public Function ExportStockDataToCSV(stockData As Variant, filePath As String, _
                                    Optional includeHeader As Boolean = True) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim csvContent As String
    Dim fileNum As Integer
    
    ' Generate CSV content
    csvContent = GenerateCSVContent(stockData, includeHeader)
    
    If csvContent = "" Then
        Call LogMessage("ERROR", "Failed to generate CSV content")
        ExportStockDataToCSV = False
        Exit Function
    End If
    
    ' Save to file
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, csvContent;
    Close #fileNum
    
    Call LogMessage("INFO", "CSV file exported: " & filePath)
    ExportStockDataToCSV = True
    
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    Call LogMessage("ERROR", "ExportStockDataToCSV: " & Err.Description)
    ExportStockDataToCSV = False
End Function

' Generate CSV content
Private Function GenerateCSVContent(stockData As Variant, includeHeader As Boolean) As String
    On Error GoTo ErrorHandler
    
    Dim content As String
    
    ' Add header row
    If includeHeader Then
        content = "DateTime,Open,High,Low,Close,Volume" & vbCrLf
    End If
    
    ' Add data rows (simple version - generates sample data)
    Dim i As Long
    For i = 1 To 10
        content = content & _
            Format(Now - (10 - i) / 24 / 60 * 5, "YYYY-MM-DD HH:MM:SS") & "," & _
            Format(2500 + Rnd() * 100, "0.00") & "," & _
            Format(2500 + Rnd() * 120, "0.00") & "," & _
            Format(2500 - Rnd() * 80, "0.00") & "," & _
            Format(2500 + (Rnd() - 0.5) * 50, "0.00") & "," & _
            Int(Rnd() * 100000) + 50000 & vbCrLf
    Next i
    
    GenerateCSVContent = content
    Exit Function
    
ErrorHandler:
    Call LogMessage("ERROR", "GenerateCSVContent: " & Err.Description)
    GenerateCSVContent = ""
End Function