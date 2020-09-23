VERSION 5.00
Object = "{85BBCFB7-F3EF-47D3-A6F8-F5057E73E1FE}#1.0#0"; "BUTTON.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   ScaleHeight     =   735
   ScaleWidth      =   855
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin KDCButton107.KDCButton KDCButton1 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Appearance      =   13
      Caption         =   "/\"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin KDCButton107.KDCButton KDCButton2 
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Appearance      =   13
      Caption         =   "\/"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin KDCButton107.KDCButton KDCButton3 
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Appearance      =   13
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Sub KDCButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num& = ScrollText&(Text1, -1)
End Sub

Private Sub KDCButton2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Num& = ScrollText&(Text1, 1)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Str(GetNumberOfLines(Text1)) = 0 Then
    KDCButton3.Height = KDCButton2.Top
End If
KDCButton3.Height = Str(GetNumberOfLines(Text1)) - 20
End Sub

Private Sub UserControl_Resize()
Text1.Height = UserControl.Height
Text1.Width = UserControl.Width - 240
KDCButton1.Left = Text1.Width
KDCButton3.Left = Text1.Width
KDCButton2.Top = UserControl.Height - 255
KDCButton2.Left = UserControl.Width - 255
End Sub
