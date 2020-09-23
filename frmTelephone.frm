VERSION 5.00
Begin VB.Form frmTelephone 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dialling..."
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrEnder 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   6000
   End
   Begin VB.Timer tmrMove2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   6000
   End
   Begin VB.Timer tmrNew 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   6000
   End
   Begin VB.PictureBox picRect 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   840
      ScaleHeight     =   375
      ScaleWidth      =   8175
      TabIndex        =   1
      Top             =   555
      Width           =   8175
   End
   Begin VB.Timer tmrMove1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   6000
   End
   Begin VB.Timer tmrDetail1 
      Interval        =   6500
      Left            =   600
      Top             =   6000
   End
   Begin VB.Timer tmrRect 
      Interval        =   350
      Left            =   120
      Top             =   6000
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H0000FF00&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   600
      Top             =   555
      Width           =   255
   End
   Begin VB.Label lblDetail 
      BackColor       =   &H00000000&
      Caption         =   "Call trans opt: received. 2-19-98 13:24:18 REC:Log>"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   8415
   End
End
Attribute VB_Name = "frmTelephone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub tmrDetail1_Timer()
tmrRect.Enabled = False ' Disable these timers when displaying the first line of text
shpRect.Visible = True
tmrMove1.Enabled = True
End Sub
Private Sub tmrEnder_Timer()
frmTrace.Show ' Show forms, unload forms and disable timers when closing main form
tmrRect.Enabled = False
shpRect.Visible = True
Unload Me
End Sub
Private Sub tmrMove1_Timer()
lblDetail.Visible = True
tmrDetail1.Enabled = False

shpRect.Left = shpRect.Left + 60 ' Move pictures along to create the illusion of text being written letter by letter
picRect.Left = picRect.Left + 60

If shpRect.Left > 9080 Then ' Reached the end of the text so reset values
    tmrRect.Enabled = True
    tmrMove1.Enabled = False
    tmrNew.Enabled = True
End If
End Sub
Private Sub tmrMove2_Timer()
tmrNew.Enabled = False
tmrRect.Enabled = False

shpRect.Left = shpRect.Left + 60 ' Move pictures along to create the illusion of text being written letter by letter
picRect.Left = picRect.Left + 60

If shpRect.Left > 4380 Then ' Reached the end of the text so reset values
    tmrRect.Enabled = True
    tmrMove2.Enabled = False
    tmrEnder.Enabled = True
End If
End Sub
Private Sub tmrNew_Timer()
tmrRect.Enabled = True
lblDetail.Caption = "Trace program: running" ' Update text caption to read second line of text
shpRect.Left = 600 ' Adjust picture widths for the text
picRect.Left = 840
tmrMove2.Enabled = True
End Sub
Private Sub tmrRect_Timer()
If shpRect.Visible = True Then ' Creates the flashing green caret when nothing is being "written"
    shpRect.Visible = False
Else
    shpRect.Visible = True
End If
End Sub

