Attribute VB_Name = "Module1"


#If Win32 Then
    Declare Function PutFocus Lib "USER32" Alias "SetFocus" (ByVal hwnd As Long) As Long
    Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

#Else
    Declare Function PutFocus% Lib "user" Alias "SetFocus" (ByVal hWd%)
    Declare Function SendMessage& Lib "user" (ByVal hWd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
#End If

Function ScrollText&(TextBox As Control, vLines As Integer)

    #If Win32 Then
        Dim success As Long
        Dim SavedWnd As Long
        Dim R As Long
    #Else
        Dim success As Integer
        Dim SavedWnd As Integer
        Dim R As Integer
    #End If
    
    Const EM_LINESCROLL = &HB6
    

    
    Lines& = vLines
    

    success = SendMessage(TextBox.hwnd, EM_LINESCROLL, 0, Lines&)
    

    R = PutFocus(SavedWnd)
    

    ScrollText& = success
End Function

