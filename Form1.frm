VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_ed As Boolean
Private Sub Command1_Click()
flag_ed = Not flag_ed
If flag_ed Then
  Command1.Caption = "Disable KeyBoard"
Else
  Command1.Caption = "Enable KeyBoard"
End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If flag_ed Then
Else
KeyCode = 0
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If flag_ed Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If flag_ed Then
Else
KeyCode = 0
End If
End Sub

Private Sub Form_Load()
flag_ed = True
Command1.Caption = "Disable KeyBoard"
End Sub


