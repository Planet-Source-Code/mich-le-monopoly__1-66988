VERSION 5.00
Begin VB.Form frmShuffleChance 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin Project1.XP_ProgressBar pgbShuffle 
      Height          =   360
      Left            =   67
      TabIndex        =   0
      Top             =   735
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
      Caption         =   "SHUFFLING CARDS"
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
      Left            =   562
      TabIndex        =   1
      Top             =   135
      Width           =   3630
   End
End
Attribute VB_Name = "frmShuffleChance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()

    pgbShuffle.Value = pgbShuffle.Value + 1
    
    If pgbShuffle.Value = 100 Then
        pgbShuffle.Value = 0
        Timer1.Enabled = False
        Call frmChance.ShowCard
        frmShuffleChance.Visible = False
        End If
    
End Sub


