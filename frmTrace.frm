VERSION 5.00
Begin VB.Form frmTrace 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Trace Program"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   120
      Top             =   6240
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   27
      Left            =   11040
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "frmTrace.frx":0000
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   26
      Left            =   11520
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "frmTrace.frx":003D
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   25
      Left            =   12000
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "frmTrace.frx":007A
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   24
      Left            =   12480
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   27
      Text            =   "frmTrace.frx":00B7
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   23
      Left            =   12960
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "frmTrace.frx":00F4
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   22
      Left            =   13440
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "frmTrace.frx":0131
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   21
      Left            =   13920
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "frmTrace.frx":016E
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   20
      Left            =   7680
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "frmTrace.frx":01AB
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   19
      Left            =   8160
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "frmTrace.frx":01E8
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   18
      Left            =   8640
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "frmTrace.frx":0225
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   17
      Left            =   9120
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "frmTrace.frx":0262
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   16
      Left            =   9600
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "frmTrace.frx":029F
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   15
      Left            =   10080
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frmTrace.frx":02DC
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   14
      Left            =   10560
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frmTrace.frx":0319
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   13
      Left            =   4320
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "frmTrace.frx":0356
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   12
      Left            =   4800
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "frmTrace.frx":0393
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   11
      Left            =   5280
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frmTrace.frx":03D0
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   10
      Left            =   5760
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmTrace.frx":040D
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   9
      Left            =   6240
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmTrace.frx":044A
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   8
      Left            =   6720
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmTrace.frx":0487
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   7
      Left            =   7200
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frmTrace.frx":04C4
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   360
      TabIndex        =   9
      Text            =   "Retreiving Number:"
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   5040
      MaxLength       =   27
      TabIndex        =   8
      Top             =   120
      Width           =   10215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   5760
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   99000
      TabIndex        =   0
      Top             =   7440
      Width           =   165
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   4800
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   6
      Left            =   3840
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmTrace.frx":0501
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   5
      Left            =   3360
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmTrace.frx":053E
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   4
      Left            =   2880
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmTrace.frx":057B
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   3
      Left            =   2400
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmTrace.frx":05B8
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   2
      Left            =   1920
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmTrace.frx":05F5
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   1
      Left            =   1440
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmTrace.frx":0632
      Top             =   600
      Width           =   285
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Index           =   0
      Left            =   960
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmTrace.frx":066F
      Top             =   600
      Width           =   285
   End
End
Attribute VB_Name = "frmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sec As Integer
Private NCount As Integer
Private X As Integer
Private I As Integer

Private Sub Form_Click()
Unload Me ' close form
End Sub
Private Sub Form_Load()
frmTelephone.tmrEnder.Enabled = False ' End timers
Sec = 0 ' Clear variables
NCount = 0

For X = 0 To 27
    Text1(X).Text = "012345678901234567890123456789" ' Create text in each text boxes
Next X
End Sub
Private Sub Form_Unload(Cancel As Integer)
frmTrace.Timer1.Enabled = False ' Disable all form timers
frmTrace.Timer2.Enabled = False
frmTrace.Timer3.Enabled = False
frmTrace.Timer4.Enabled = False
End Sub
Private Sub Timer1_Timer()
Randomize ' Randomize seed

For X = 0 To 27 ' For every text box
    For I = 0 To 5 ' slows the random speed so that the characters can be read, if the program runs slowly comment this line
        Text1(X).Text = Int(Rnd * 9) + 1 & Text1(X).Text ' append a new number to the text
    Next I ' if the program runs slowly comment this line
Next X

If Sec = 10 Then
    Timer1.Enabled = False ' Reached ten secs so time to delete a line
    Timer2.Enabled = False
    Timer3.Enabled = True
    Sec = 0 ' Reset value
Else
    Sec = Sec + 1 ' Up the value
End If
End Sub
Private Sub Timer2_Timer()
If X <= 28 Then X = 0: Exit Sub ' Safe guard for errors - thanks Roger Gilchrist for pointing this out
Text1(X).Text = "" ' Reset the text
End Sub
Private Sub Timer3_Timer()
Timer1.Enabled = False ' Disable timers
Timer2.Enabled = False
Text1(0).Visible = False ' This textbox flickers for some reason so remove it first
Call DeleteLine ' Delete a line
End Sub
Private Sub DeleteLine()
On Error Resume Next
Dim intmain As Integer
intmain = (Int(Rnd * 27) + 1) ' Get random number between 1 and 27

If Text1(intmain).Visible = False Then ' textbox not visible so run sub again
    Call DeleteLine
Else
    Text1(intmain).Visible = False ' Delete the textbox with random number if it is visible
    Text2.Text = Text2.Text & Int(Rnd * 9) + 1 ' Add a new digit to the phone number
End If

Timer3.Enabled = False ' Renable timers
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Timer4_Timer()
For X = 0 To 27 ' Each time a line is deleted update NCount
    If Text1(X).Visible = False Then
        NCount = NCount + 1
    End If
Next

If NCount = 27 Then ' When 27 is reached close form
    frmTrace2.Text2.Text = Text2.Text
    frmTrace2.Show
    Unload Me
    Exit Sub
Else
    NCount = 0 ' Reset value
End If
End Sub
