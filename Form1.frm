VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parking System"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   480
      Width           =   3135
   End
   Begin VB.PictureBox Picture6 
      Height          =   975
      Left            =   4080
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   13
      Top             =   4320
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      Height          =   975
      Left            =   4080
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   12
      Top             =   3000
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      Height          =   975
      Left            =   4080
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   1080
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   1080
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   1080
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Timer Timer7 
      Interval        =   10000
      Left            =   2880
      Top             =   1440
   End
   Begin VB.Timer Timer6 
      Interval        =   500
      Left            =   9960
      Top             =   2760
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   2520
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   480
      Top             =   4440
   End
   Begin VB.CommandButton Command7 
      Caption         =   "START"
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Text            =   "2"
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   855
      Left            =   14520
      TabIndex        =   4
      Text            =   "Text8"
      Top             =   10320
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   14400
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   10080
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   14640
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   9840
      Width           =   1455
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   360
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   720
      Top             =   0
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   11160
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Jog Mode"
      Height          =   735
      Left            =   10800
      TabIndex        =   0
      Top             =   10800
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1920
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      OutBufferSize   =   1
   End
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   21
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   3960
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   19
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "COMM PORT"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j, flag As Integer
Dim k As Long
Dim B As String
Dim a As String
Dim x, y, x1 As Integer
Dim a1, b1, c1, a2, b2, c2, d1, d2, e1, e2 As Integer
Dim ai, ai1, ai2, ai3, ai4, ai5, ai6 As Integer

Private Sub Command1_Click()
MSComm1.Output = "D"
End Sub

Private Sub Command10_Click()
MSComm1.Output = "I"
End Sub

Private Sub Command11_Click()
MSComm1.Output = "A"
End Sub

Private Sub Command12_Click()
MSComm1.Output = "B"
End Sub

Private Sub Command13_Click()
MSComm1.Output = "J"
End Sub

Private Sub Command14_Click()
MSComm1.Output = "K"
End Sub

Private Sub Command15_Click()
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Command16_Click()
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Timer5.Enabled = False

End Sub

Private Sub Command2_Click()
MSComm1.Output = "C"
End Sub

Private Sub Command3_Click()
MSComm1.Output = "I"
End Sub

Private Sub Command4_Click()
MSComm1.Output = "H"
End Sub

Private Sub Command5_Click()
MSComm1.Output = "E"
End Sub

Private Sub Command7_Click()
MSComm1.CommPort = Val(Text9)
MSComm1.PortOpen = True
End Sub

Private Sub Command8_Click()
MSComm1.Output = "G"
End Sub

Private Sub Command9_Click()
MSComm1.Output = "F"
End Sub

Private Sub Form_Load()
'MSComm1.PortOpen = True

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
a = MSComm1.Input
Text1 = a
For i = 1 To 40

If Mid$(a, i, 1) = "A" Then
    If Mid$(a, i + 1, 1) = 1 Then
        Picture1.BackColor = vbRed
        ai1 = 1
    End If
    If Mid$(a, i + 1, 1) = 0 Then
        Picture1.BackColor = vbGreen
    End If
    If Mid$(a, i + 2, 1) = 1 Then
        Picture2.BackColor = vbRed
        ai2 = 1
    End If
    If Mid$(a, i + 2, 1) = 0 Then
        Picture2.BackColor = vbGreen
    End If
    If Mid$(a, i + 3, 1) = 1 Then
        Picture3.BackColor = vbRed
        ai3 = 1
    End If
    If Mid$(a, i + 3, 1) = 0 Then
        Picture3.BackColor = vbGreen
    End If
    If Mid$(a, i + 4, 1) = 1 Then
        Picture4.BackColor = vbRed
        ai4 = 1
    End If
    If Mid$(a, i + 4, 1) = 0 Then
        Picture4.BackColor = vbGreen
    End If
    If Mid$(a, i + 5, 1) = 1 Then
        Picture5.BackColor = vbRed
        ai5 = 1
    End If
    If Mid$(a, i + 5, 1) = 0 Then
        Picture5.BackColor = vbGreen
    End If
    If Mid$(a, i + 6, 1) = 1 Then
        Picture6.BackColor = vbRed
        ai6 = 1
    End If
    If Mid$(a, i + 6, 1) = 0 Then
        Picture6.BackColor = vbGreen
    End If

    If Mid$(a, i + 7, 1) = 1 Then
        Label1.Caption = "Preet Kumar In"
    End If
    If Mid$(a, i + 7, 1) = 0 Then
        Label1.Caption = "Preet Kumar Out"
    End If
    If Mid$(a, i + 8, 1) = 1 Then
        Label1.Caption = "Kashish Kedia In"
    End If
    If Mid$(a, i + 8, 1) = 0 Then
        Label1.Caption = "Kashish Kedia Out"
    End If
    If Mid$(a, i + 9, 1) = 1 Then
        Label1.Caption = "Devansh Maheshwari  In"
    End If
    If Mid$(a, i + 9, 1) = 0 Then
        Label1.Caption = "Devansh Maheshwari  Out"
    End If
    If Mid$(a, i + 10, 1) = 1 Then
        Label1.Caption = "Azam Khan In"
    End If
    If Mid$(a, i + 10, 1) = 0 Then
        Label1.Caption = "Azam Khan Out"
    End If
    If Mid$(a, i + 11, 1) = 1 Then
        Label1.Caption = "Arun Sir In"
    End If
    If Mid$(a, i + 11, 1) = 0 Then
        Label1.Caption = "Arun Sir Out"
    End If
    If Mid$(a, i + 12, 1) = 1 Then
        Label1.Caption = "Padmini Mam In"
    End If
    If Mid$(a, i + 12, 1) = 0 Then
        Label1.Caption = "Padmini Mam Out"
    End If
    
    ai = ai1 + ai2 + ai3 + ai4 + ai5 + ai6
    
If ai = 6 Then
Label8.Caption = "Parking Full"
Else
Label8.Caption = "            "
End If
End If

Next i

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
    'MSComm1.Output = "A"
End Sub

