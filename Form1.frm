VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MATRIX Textbox Control"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl1 UserControl11 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   2990
      _ExtentY        =   1931
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   1800
      Picture         =   "Form1.frx":08CA
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "Please rate my code:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------
'| If you dont have the KDC buttons you can
'|download them off http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=27225&lngWId=1
'|Please rate my code.
'-------------------------
Private Declare Function GetDesktopWindow Lib "USER32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

Private Sub Image1_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = "http://www.pscode.com"
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL
End Sub
Public Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  

   hWndDesk = GetDesktopWindow()

   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)


  If success = SE_ERR_NOASSOC Then
    MsgBox "Couldn't load the default application"
    Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
End Sub

