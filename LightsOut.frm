VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lights Out"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LightsOut.frx":0000
   ScaleHeight     =   4185
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Win 
      Interval        =   10
      Left            =   2400
      Top             =   240
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   240
      Picture         =   "LightsOut.frx":9664
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3720
      Width           =   3375
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   24
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   24
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   23
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   22
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   22
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   21
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   21
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   20
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   20
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   19
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   19
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   18
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   18
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   17
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   16
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   15
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   14
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   13
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   12
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   11
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   10
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   9
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   8
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   7
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   6
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   5
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   4
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   3
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   2
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox Light 
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Red = &HFF&
Const Black = &H0&
Dim I As Integer
Dim TheNum As Integer
Dim Counter As Integer

Private Sub Command1_Click()
Form_Load
End Sub

Private Sub Form_Load()
For j = 0 To 24
    Randomize
    TheNum = Int(Rnd * 5)
    If TheNum = 1 Or TheNum = 3 Then
        Light(j).BackColor = Red
    End If
    If TheNum = 2 Or TheNum = 4 Or TheNum = 5 Then
        Light(j).BackColor = Black
    End If
Next
End Sub

Private Sub Light_Click(Index As Integer)
I = Index
If Light(I).BackColor = Red Then
    Light(I).BackColor = Black
Else
    Light(I).BackColor = Red
End If

If I = 0 Or I = 5 Or I = 10 Or I = 15 Or I = 20 Then
    GoTo Next1:
End If
If Light(I - 1).BackColor = Red Then
    Light(I - 1).BackColor = Black
Else
    Light(I - 1).BackColor = Red
End If

Next1:
If I = 4 Or I = 9 Or I = 14 Or I = 19 Or I = 24 Then
    GoTo Next2:
End If
If Light(I + 1).BackColor = Red Then
    Light(I + 1).BackColor = Black
Else
    Light(I + 1).BackColor = Red
End If


Next2:
If I = 0 Or I = 1 Or I = 2 Or I = 3 Or I = 4 Then
    GoTo Next3:
End If
If Light(I - 5).BackColor = Red Then
    Light(I - 5).BackColor = Black
Else
    Light(I - 5).BackColor = Red
End If

Next3:
On Error Resume Next
If Light(I + 5).BackColor = Red Then
    Light(I + 5).BackColor = Black
Else
    Light(I + 5).BackColor = Red
End If

End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Form1.Caption = "By"
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
Form1.Caption = "Joe Fitz."
Timer5.Enabled = True
End Sub

Private Sub Timer5_Timer()
Timer5.Enabled = False
Form1.Caption = "Lights Out"
Timer2.Enabled = True

End Sub

Private Sub Win_Timer()
If Light(0).BackColor = Black And Light(1).BackColor = Black And Light(2).BackColor = Black And Light(3).BackColor = Black And Light(4).BackColor = Black And Light(5).BackColor = Black And Light(6).BackColor = Black And Light(7).BackColor = Black And Light(8).BackColor = Black And Light(9).BackColor = Black And Light(10).BackColor = Black And Light(11).BackColor = Black And Light(12).BackColor = Black And Light(13).BackColor = Black And Light(14).BackColor = Black And Light(15).BackColor = Black And Light(16).BackColor = Black And Light(17).BackColor = Black And Light(18).BackColor = Black And Light(19).BackColor = Black And Light(20).BackColor = Black And Light(21).BackColor = Black And Light(22).BackColor = Black And Light(23).BackColor = Black And Light(24).BackColor = Black Then
    MsgBox "Lights Out!", vbOKOnly, "By Joe Fitz."
    Form_Load
End If
End Sub
