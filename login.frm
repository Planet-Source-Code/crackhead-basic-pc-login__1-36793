VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Secure Login"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Pass:"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Logon:"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "The Pass Phrase Is?"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Login:"
      Height          =   2775
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Enter Secure Login"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long

Const SPI_SCREENSAVERRUNNING = 97

Private Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByVal lpvParam As Any, _
    ByVal fuWinIni As Long) As Long
Private Sub DisableCtrlAltDel()

RegisterServiceProcess GetCurrentProcessId, 1

End Sub
Private Sub FUCKYOU()
    MsgBox ("FuckOff!")
    
End Sub

'ENABLE below
'Public Sub EnableCtrlAltDel()
'RegisterServiceProcess GetCurrentProcessId, 0
'End Sub

Private Sub Command1_Click()
    log1 = Text1.Text
    pass1 = Text2.Text
    getid$ = Text1.Text & Text2.Text
    Text3.Text = getid
    
    'Turn On CTRL ALT DELETE! Otherwise, disables function!!
    'Set login and password!!
    If getid$ = "testprogram" Then ' getid$ = Login and Pass added together
    '                   ^----trick, can be all one word in either box! Cool!
    Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, "1", 0)
    MsgBox ("Thank You, Login Correct!")
        End
    Else
        GoTo msg1
       End If
msg1:
           MsgBox ("Invalid Id or Password")
           Exit Sub
           
End Sub

Private Sub Command2_Click()
MsgBox ("Disabled, please login!")
'End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'vbKeyEscape and other escape keys
    If KeyAscii = 27 Then MsgBox ("FuckOff!")
        
End Sub

Private Sub Form_Load()
    Call DisableCtrlAltDel '
    ' ^-----remove from ctrl alt delete window
    
    Call SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, "1", 0) '
    '  ^---Turn Off CTRL ALT DELETE all together along with alt f4, alt tab, etc...
    
    
    
End Sub

