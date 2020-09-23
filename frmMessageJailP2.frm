VERSION 5.00
Begin VB.Form frmMessageJailP2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Bid"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   Picture         =   "frmMessageJailP2.frx":0000
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.XP_ProgressBar pgbCPU 
      Height          =   330
      Left            =   360
      TabIndex        =   8
      Top             =   1710
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   582
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
      Scrolling       =   9
   End
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   4470
      Top             =   3510
   End
   Begin VB.TextBox txtThrows 
      Height          =   285
      Left            =   4110
      TabIndex        =   7
      Top             =   840
      Width           =   630
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay the Fine"
      Height          =   405
      Left            =   2175
      TabIndex        =   4
      Top             =   2775
      Width           =   1380
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4875
      Top             =   1290
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4140
      Top             =   1320
   End
   Begin VB.Image imgCOOJ 
      Height          =   945
      Left            =   3555
      Picture         =   "frmMessageJailP2.frx":0E24
      Stretch         =   -1  'True
      Top             =   3435
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgCCOOJ 
      Height          =   945
      Left            =   1515
      Picture         =   "frmMessageJailP2.frx":1CA6
      Stretch         =   -1  'True
      Top             =   3435
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CPU Decision:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2295
      TabIndex        =   9
      Top             =   1440
      Width           =   1290
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Throws:"
      Height          =   240
      Left            =   3525
      TabIndex        =   6
      Top             =   915
      Width           =   660
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "You will be moved out of Jail and wait for your turn."
      Height          =   255
      Left            =   1215
      TabIndex        =   5
      Top             =   2355
      Width           =   3675
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   25
      X2              =   25
      Y1              =   146
      Y2              =   225
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   25
      X2              =   368
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   368
      X2              =   368
      Y1              =   146
      Y2              =   225
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay a R 5'000.00 Fine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   675
      TabIndex        =   3
      Top             =   2055
      Width           =   2310
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   24
      X2              =   39
      Y1              =   146
      Y2              =   146
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   202
      X2              =   368
      Y1              =   146
      Y2              =   146
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the dice to throw."
      Height          =   255
      Left            =   1170
      TabIndex        =   2
      Top             =   555
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "You have to throw a double to be able to move on."
      Height          =   255
      Left            =   1170
      TabIndex        =   1
      Top             =   345
      Width           =   3675
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   165
      X2              =   369
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   25
      X2              =   40
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Throw to get out!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   690
      TabIndex        =   0
      Top             =   90
      Width           =   2310
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   369
      X2              =   369
      Y1              =   15
      Y2              =   94
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   26
      X2              =   369
      Y1              =   94
      Y2              =   94
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   26
      X2              =   26
      Y1              =   15
      Y2              =   94
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   1
      Left            =   2895
      Picture         =   "frmMessageJailP2.frx":2BA3
      Top             =   840
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   0
      Left            =   2280
      Picture         =   "frmMessageJailP2.frx":2EAD
      Top             =   840
      Width           =   480
   End
End
Attribute VB_Name = "frmMessageJailP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim Ap As String
    Dim mRollCount As Integer
    Dim mDice1 As Integer
    Dim mDice2 As Integer
    Dim mDiceTries As Integer
    Dim mPlayer As Integer
    Dim mPlayerBank As Long
    Dim mDecision As Integer
    
Private Sub cmdPay_Click()
    Dim Dice1 As Integer
    Dim Dice2 As Integer
    
    With frmBoard.imgPlayer2
        .Picture = LoadPicture(Ap & frmBoard.imgPlayer2.Tag & "\Right1.ico")
        .Left = 164
        .Top = 0
        .ToolTipText = 9
    End With

    Call Randomize
    
    Dice1 = 1 + Int(Rnd() * 6)
    Dice2 = 1 + Int(Rnd() * 6)

    With frmBoard.imgDice
        .Item(0).Picture = LoadPicture(Ap & "dice\" & Dice1 & ".ico")
        .Item(1).Picture = LoadPicture(Ap & "dice\" & Dice2 & ".ico")
    End With

    frmBoard.imgPlayer1Turn.Visible = False
    frmBoard.imgPlayer2Turn.Visible = True
    
    mPlayerBank = mPlayerBank - 5000
    frmBoard.lsbOne.AddItem ("Player 2 paid the R 5'000.00 and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
    Call frmBoard.Player2OutOfJail(Dice1, Dice2, 2, mPlayerBank)
    
    mDiceTries = 0
    
    frmMessageJailP2.Visible = False

End Sub

Private Sub Form_Load()

    Call Randomize

    If Right(App.Path, 1) = "\" Then
        Ap = App.Path
    Else
        Ap = App.Path & "\"
    End If

    txtThrows.Text = mDiceTries

End Sub

Private Sub imgCCOOJ_Click()
    Dim Dice1 As Integer
    Dim Dice2 As Integer
    
    With frmBoard.imgPlayer2
        .Picture = LoadPicture(Ap & frmBoard.imgPlayer2.Tag & "\Right1.ico")
        .Left = 164
        .Top = 0
        .ToolTipText = 9
    End With

    Call Randomize
    
    Dice1 = 1 + Int(Rnd() * 6)
    Dice2 = 1 + Int(Rnd() * 6)

    With frmBoard.imgDice
        .Item(0).Picture = LoadPicture(Ap & "dice\" & Dice1 & ".ico")
        .Item(1).Picture = LoadPicture(Ap & "dice\" & Dice2 & ".ico")
    End With

    frmBoard.imgPlayer1Turn.Visible = False
    frmBoard.imgPlayer2Turn.Visible = True
    
    mPlayerBank = mPlayerBank
    frmBoard.lsbOne.AddItem ("Player 2 used Free Card and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
    Call frmBoard.Player2OutOfJail(Dice1, Dice2, 2, mPlayerBank)
    
    mDiceTries = 0
    
    frmMessageJailP2.Visible = False

End Sub

Private Sub imgCOOJ_Click()
    Dim Dice1 As Integer
    Dim Dice2 As Integer
    
    With frmBoard.imgPlayer2
        .Picture = LoadPicture(Ap & frmBoard.imgPlayer2.Tag & "\Right1.ico")
        .Left = 164
        .Top = 0
        .ToolTipText = 9
    End With

    Call Randomize
    
    Dice1 = 1 + Int(Rnd() * 6)
    Dice2 = 1 + Int(Rnd() * 6)

    With frmBoard.imgDice
        .Item(0).Picture = LoadPicture(Ap & "dice\" & Dice1 & ".ico")
        .Item(1).Picture = LoadPicture(Ap & "dice\" & Dice2 & ".ico")
    End With

    frmBoard.imgPlayer1Turn.Visible = False
    frmBoard.imgPlayer2Turn.Visible = True
    
    mPlayerBank = mPlayerBank
    frmBoard.lsbOne.AddItem ("Player 2 used Free Card and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
    Call frmBoard.Player2OutOfJail(Dice1, Dice2, 2, mPlayerBank)
    
    mDiceTries = 0
    
    frmMessageJailP2.Visible = False

End Sub

Private Sub imgDice_Click(Index As Integer)
    
    mDiceTries = mDiceTries + 1
    txtThrows.Text = mDiceTries
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    Dim Dice1 As Integer
    Dim Dice2 As Integer
    
    Call Randomize
    
    Dice1 = 1 + Int(Rnd() * 6)
    Dice2 = 1 + Int(Rnd() * 6)

    mRollCount = mRollCount + 1

    With imgDice
        .Item(0).Picture = LoadPicture(Ap & "dice\" & Dice1 & ".ico")
        .Item(1).Picture = LoadPicture(Ap & "dice\" & Dice2 & ".ico")
    End With

    If mRollCount = 15 Then
        Timer1.Enabled = False
        mRollCount = 0
        mDice1 = Dice1
        mDice2 = Dice2
        Timer2.Enabled = True
    End If
    
End Sub

Private Sub CheckDice()
    Dim msg As String
    
    If mDice1 = mDice2 Then
        
        mDiceTries = 0
        
        With frmBoard.imgPlayer2
            .Picture = LoadPicture(Ap & frmBoard.imgPlayer2.Tag & "\Right1.ico")
            .Left = 164
            .Top = 0
            .ToolTipText = 9
        End With

        With frmBoard.imgDice
            .Item(0).Picture = LoadPicture(Ap & "dice\" & mDice1 & ".ico")
            .Item(1).Picture = LoadPicture(Ap & "dice\" & mDice2 & ".ico")
        End With

        frmBoard.imgPlayer1Turn.Visible = False
        frmBoard.imgPlayer2Turn.Visible = True
        frmBoard.lsbOne.AddItem ("Player 2 threw a double and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
        Call frmBoard.Player2OutOfJail(mDice1, mDice2, 2, mPlayerBank)
    
        frmMessageJailP2.Visible = False
    Else
    
        If mDiceTries = 3 Then
            
            mDiceTries = 0
            
            With frmBoard.imgPlayer2
                .Picture = LoadPicture(Ap & frmBoard.imgPlayer2.Tag & "\Right1.ico")
                .Left = 164
                .Top = 0
                .ToolTipText = 9
            End With

            With frmBoard.imgDice
                .Item(0).Picture = LoadPicture(Ap & "dice\" & mDice1 & ".ico")
                .Item(1).Picture = LoadPicture(Ap & "dice\" & mDice2 & ".ico")
            End With

            frmBoard.imgPlayer1Turn.Visible = False
            frmBoard.imgPlayer2Turn.Visible = True
    
            mPlayerBank = mPlayerBank - 5000
            frmBoard.lsbOne.AddItem ("Player 2 paid the R 5'000.00 and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
            Call frmBoard.Player2OutOfJail(mDice1, mDice2, 2, mPlayerBank)
    
            frmMessageJailP2.Visible = False
        Else
            frmBoard.imgPlayer1Turn.Visible = True
            frmBoard.imgPlayer2Turn.Visible = False
            frmMessageJailP2.Visible = False
        End If
    End If
    
        
End Sub

Private Sub Timer2_Timer()

    Call CheckDice
    Timer2.Enabled = False
    
End Sub

Public Sub Info(Player As Integer, Bank As Long)

    mPlayer = Player
    mPlayerBank = Bank

End Sub

Private Sub Timer3_Timer()

    pgbCPU.Value = pgbCPU.Value + 1
    
    
    If pgbCPU.Value = 100 Then
    
        mDecision = 1 + Int(Rnd() * 2)
        
        If mDecision = 1 Then
            Call imgDice_Click(1)
            pgbCPU.Value = 0
            Timer3.Enabled = False
        End If
        
        If mDecision = 2 Then
        
            If imgCCOOJ.Visible = True Then
                Call imgCCOOJ_Click
                pgbCPU.Value = 0
                Timer3.Enabled = False
                Exit Sub
            End If
            
            If imgCOOJ.Visible = True Then
                Call imgCOOJ_Click
                pgbCPU.Value = 0
                Timer3.Enabled = False
                Exit Sub
            End If
            
            Call cmdPay_Click
            pgbCPU.Value = 0
            Timer3.Enabled = False
            
        End If
    End If
    
End Sub
