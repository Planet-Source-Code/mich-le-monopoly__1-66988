VERSION 5.00
Begin VB.Form frmCPUThinking 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   30
      Top             =   30
   End
   Begin Project1.XP_ProgressBar pgbCPUThink 
      Height          =   360
      Left            =   30
      TabIndex        =   0
      Top             =   765
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   255
      Scrolling       =   9
   End
   Begin VB.Label Label1 
      Caption         =   "CPU CALCULATING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   525
      TabIndex        =   1
      Top             =   165
      Width           =   3630
   End
End
Attribute VB_Name = "frmCPUThinking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

    pgbCPUThink.Value = pgbCPUThink.Value + 1
    
    If pgbCPUThink.Value = 100 Then
        pgbCPUThink.Value = 0
        Timer1.Enabled = False
        frmCPUThinking.Visible = False
        Call frmBoard.Player2GameCont2(frmBoard.imgPlayer2.Left, frmBoard.imgPlayer2.Top)
    End If
    
End Sub
