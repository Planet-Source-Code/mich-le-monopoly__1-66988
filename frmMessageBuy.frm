VERSION 5.00
Begin VB.Form frmMessageBuy 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Bid"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   Picture         =   "frmMessageBuy.frx":0000
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPlayer1Resign 
      Caption         =   "Player 1 Resign"
      Height          =   450
      Left            =   3945
      TabIndex        =   11
      Top             =   3330
      Width           =   1650
   End
   Begin VB.CommandButton cmdPlayer1Bid 
      Caption         =   "Player 1 Bid"
      Height          =   450
      Left            =   3945
      TabIndex        =   8
      Top             =   2775
      Width           =   1650
   End
   Begin Project1.XP_ProgressBar pgbOne 
      Height          =   2985
      Left            =   1365
      TabIndex        =   0
      Top             =   750
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   5265
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
      Orientation     =   1
   End
   Begin Project1.XP_ProgressBar pgbTwo 
      Height          =   2985
      Left            =   2490
      TabIndex        =   7
      Top             =   750
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   5265
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
      Color           =   16711680
      Orientation     =   1
   End
   Begin VB.Label lblPlayer2Bank 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3945
      TabIndex        =   10
      Top             =   630
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblPlayer1Bank 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3945
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblPlayer2Bid 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3975
      TabIndex        =   6
      Top             =   2325
      Width           =   1590
   End
   Begin VB.Label lblPlayer1Bid 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3975
      TabIndex        =   5
      Top             =   1575
      Width           =   1590
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player2 Bid:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3975
      TabIndex        =   4
      Top             =   1950
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player1 Bid:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3975
      TabIndex        =   3
      Top             =   1200
      Width           =   1590
   End
   Begin VB.Label lblMsgPropertyName 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   2
      Top             =   225
      Width           =   1590
   End
   Begin VB.Label lblMsgPropertyPrice 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2415
      TabIndex        =   1
      Top             =   225
      Width           =   1590
   End
   Begin VB.Image imgPlayer2 
      Height          =   480
      Left            =   2475
      Top             =   3750
      Width           =   480
   End
   Begin VB.Image imgPlayer1 
      Height          =   480
      Left            =   1335
      Top             =   3750
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   250
      X2              =   50
      Y1              =   250
      Y2              =   250
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   50
      X2              =   50
      Y1              =   50
      Y2              =   250
   End
End
Attribute VB_Name = "frmMessageBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim Winner As Integer
    Dim LuckyNumber As Integer
    Dim LuckyGuess As Integer

Private Sub cmdPlayer1Bid_Click()
        
    Call CPUBid
    
End Sub

Private Sub CPUBid()
    Dim Player2Bank As Double, Bid As Double
    
    Player2Bank = Mid$(lblPlayer2Bank.Caption, 2)
    Bid = pgbOne.Value + 500

    If LuckyNumber = LuckyGuess Then
        Bid = MsgBox("Player 1 has won the Lotto. You may have the property.", vbOKOnly, "Player 2 Resigns Bid.")
        Winner = 1
        Call frmBoard.MsgYes(Winner, pgbOne.Value)
        frmMessageBuy.Hide
        Exit Sub
    End If
    
    If Bid > Player2Bank Then
        Bid = MsgBox("Player 1 has out bid me. You may have the property.", vbOKOnly, "Player 2 Resigns Bid.")
        Winner = 1
        Call frmBoard.MsgYes(Winner, pgbOne.Value)
        frmMessageBuy.Hide
    ElseIf Bid >= Mid$(lblMsgPropertyPrice.Caption, 2) Then
        Bid = MsgBox("Player 1 may have the property. It's Not worth the price.", vbOKOnly, "Player 2 Resigns Bid.")
        Winner = 1
        Call frmBoard.MsgYes(Winner, pgbOne.Value)
        frmMessageBuy.Hide
    Else
        pgbTwo.Value = Bid
        lblPlayer2Bid.Caption = FormatCurrency(pgbTwo.Value, 2)
        cmdPlayer1Bid.Visible = True
        cmdPlayer1Resign.Visible = True
    End If

End Sub
Public Function FormLoadNow()
    
    Call Randomize
        
    LuckyNumber = 1 + Int(Rnd() * 10)
    LuckyGuess = InputBox("Guess a number between 1 and 10. Guess right and you can get this property VERY cheap.", "Lucky Lotto")
    lblPlayer1Bid.Caption = FormatCurrency(pgbOne.Value, 2)
    pgbTwo.Value = 0
    lblPlayer2Bid.Caption = FormatCurrency(pgbTwo.Value, 2)
    
End Function

Private Sub cmdPlayer1Resign_Click()
    Dim R As String
    
    R = MsgBox("Are you sure you want to leave the Auction?", vbOKCancel, "Leave Auction?")
    
    If R = vbOK Then
        Call frmBoard.MsgYes(2, pgbTwo.Value)
        frmMessageBuy.Hide
    ElseIf R = vbCancel Then
        ' Do nothing
    End If
    
End Sub

Private Sub imgPlayer1_Click()

    If pgbOne.Value < pgbTwo.Value Then
        pgbOne.Value = pgbTwo.Value + 500
    Else
        pgbOne.Value = pgbOne.Value + 500
    End If
    
    lblPlayer1Bid.Caption = FormatCurrency(pgbOne.Value, 2)
    cmdPlayer1Bid.Visible = True
    cmdPlayer1Resign.Visible = True
    
End Sub
