Attribute VB_Name = "Module2"
Option Explicit

Public Declare Function SendMessageBynum Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9

Public Function GetLineNumber(txtTextBox As TextBox) As Long
    Dim lngSelectedText As Long
    Dim lngLineNumber As Long
    

    lngSelectedText = SendMessageBynum&(txtTextBox.hwnd, EM_GETSEL, 0, 0&)
    

    lngLineNumber = SendMessageBynum(txtTextBox.hwnd, EM_LINEFROMCHAR, lngSelectedText, 0&)
    

    GetLineNumber = lngLineNumber
End Function

Public Function GetNumberOfLines(txtTextBox As TextBox) As Long
    Dim lngNumberOfLines As Long

    lngNumberOfLines = SendMessageBynum(txtTextBox.hwnd, EM_GETLINECOUNT, 0, 0&)
    

    GetNumberOfLines = lngNumberOfLines
End Function

