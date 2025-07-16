'******************************************************************************
' Simple Test Module for VBA Functionality
' 
' Description: Basic test functions to verify VBA setup
' Author: Claude Code
' Version: 1.0.0
'******************************************************************************

Option Explicit

' Simple test function
Public Sub TestBasic()
    MsgBox "VBA is working correctly!", vbInformation, "Test Result"
    Debug.Print "Basic test completed at " & Now
End Sub

' Test with stock data collection
Public Sub TestStockCollection()
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim stockCode As String
    
    stockCode = "7203"  ' Toyota Motor
    Debug.Print "Testing stock data collection for: " & stockCode
    
    result = CollectStockData(stockCode, "5M", Date - 1, Date)
    
    If result Then
        MsgBox "Stock data collection test successful for " & stockCode, vbInformation, "Test Success"
    Else
        MsgBox "Stock data collection test failed for " & stockCode, vbExclamation, "Test Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in test: " & Err.Description, vbCritical, "Test Error"
    Debug.Print "Test error: " & Err.Description
End Sub

' Test multiple stocks
Public Sub TestMultipleStocks()
    On Error GoTo ErrorHandler
    
    Dim result As Boolean
    Dim stockCodes As String
    
    stockCodes = "7203,6758"  ' Toyota, Sony
    Debug.Print "Testing multiple stocks: " & stockCodes
    
    result = CollectMultipleStocks(stockCodes, "5M", Date - 1, Date)
    
    If result Then
        MsgBox "Multiple stocks test successful", vbInformation, "Test Success"
    Else
        MsgBox "Multiple stocks test completed with some failures", vbExclamation, "Test Completed"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in multiple stocks test: " & Err.Description, vbCritical, "Test Error"
    Debug.Print "Multiple stocks test error: " & Err.Description
End Sub

' Show main interface using simple input box
Public Sub ShowMainInterface()
    Call ShowMainForm
End Sub