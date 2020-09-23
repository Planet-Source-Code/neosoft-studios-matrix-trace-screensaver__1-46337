VERSION 5.00
Begin VB.Form frmTrace2 
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
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   120
      Top             =   5760
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   1560
      TabIndex        =   2
      Text            =   "Decoding Number:"
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   5040
      MaxLength       =   27
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   99000
      TabIndex        =   0
      Top             =   7440
      Width           =   165
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   120
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   4800
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   $"frmTrace2.frx":0000
      ForeColor       =   &H0000FF00&
      Height          =   1935
      Left            =   3360
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "441546977311"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmTrace2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private I As Integer
Private Sub Form_Click()
Unload Me ' Close form
End Sub
Private Sub Timer1_Timer()
For I = 0 To Text2.MaxLength ' Creates illusion of decoding via creation of random numbers
Text2.Text = Int(Rnd * 9) + 1 & Text2.Text
Next
End Sub
Private Sub Timer2_Timer()
If Text2.MaxLength = 1 Then ' Reduces the text length until a certain length of 12 then creates the final number
Text2.Visible = False
Label2.Visible = True
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
Exit Sub
End If

If Text2.MaxLength <= 12 Then
Timer2.Interval = 500
Text2.MaxLength = Text2.MaxLength - 1
Text2.Left = Text2.Left + 165
Exit Sub
End If

Text2.MaxLength = Text2.MaxLength - 1
End Sub
Private Sub Timer3_Timer()
Call EndSound ' Unload procedures
frmTelephone.Show
Unload Me
End Sub
