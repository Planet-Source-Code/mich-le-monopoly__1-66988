VERSION 5.00
Begin VB.Form frmMessageJail 
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
   Picture         =   "frmMessageJail.frx":0000
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtThrows 
      Height          =   285
      Left            =   4110
      TabIndex        =   7
      Top             =   870
      Width           =   630
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay the Fine"
      Height          =   405
      Left            =   2205
      TabIndex        =   4
      Top             =   2250
      Width           =   1380
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4890
      Top             =   -180
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4155
      Top             =   -180
   End
   Begin VB.Image imgCOOJ 
      Height          =   1485
      Left            =   3570
      Picture         =   "frmMessageJail.frx":0E24
      Stretch         =   -1  'True
      Top             =   2910
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgCCOOJ 
      Height          =   1485
      Left            =   1215
      Picture         =   "frmMessageJail.frx":1CA6
      Stretch         =   -1  'True
      Top             =   2895
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Throws:"
      Height          =   240
      Left            =   3525
      TabIndex        =   6
      Top             =   945
      Width           =   660
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "You will be moved out of Jail and wait for your turn."
      Height          =   255
      Left            =   1245
      TabIndex        =   5
      Top             =   1830
      Width           =   3675
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   27
      X2              =   27
      Y1              =   111
      Y2              =   190
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   27
      X2              =   370
      Y1              =   190
      Y2              =   190
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   370
      X2              =   370
      Y1              =   111
      Y2              =   190
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
      Left            =   705
      TabIndex        =   3
      Top             =   1530
      Width           =   2310
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   26
      X2              =   41
      Y1              =   111
      Y2              =   111
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   204
      X2              =   370
      Y1              =   111
      Y2              =   111
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the dice to throw."
      Height          =   255
      Left            =   1170
      TabIndex        =   2
      Top             =   585
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "You have to throw a double to be able to move on."
      Height          =   255
      Left            =   1170
      TabIndex        =   1
      Top             =   375
      Width           =   3675
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   165
      X2              =   369
      Y1              =   17
      Y2              =   17
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   25
      X2              =   40
      Y1              =   17
      Y2              =   17
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
      Top             =   120
      Width           =   2310
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   369
      X2              =   369
      Y1              =   17
      Y2              =   96
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   26
      X2              =   369
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   26
      X2              =   26
      Y1              =   17
      Y2              =   96
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   1
      Left            =   2895
      Picture         =   "frmMessageJail.frx":2BA3
      Top             =   870
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   0
      Left            =   2280
      Picture         =   "frmMessageJail.frx":2EAD
      Top             =   870
      Width           =   480
   End
End
Attribute VB_Name = "frmMessageJail"
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
    
Private Sub cmdPay_Click()
    Dim Dice1 As Integer
    Dim Dice2 As Integer
    
    With frmBoard.imgPlayer1
        .Picture = LoadPicture(Ap & frmBoard.imgPlayer1.Tag & "\Right1.ico")
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

    frmBoard.imgPlayer1Turn.Visible = True
    frmBoard.imgPlayer2Turn.Visible = False
    
    mPlayerBank = mPlayerBank - 5000
    frmBoard.lsbOne.AddItem ("Player 1 paid the R 5'000.00 and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
    Call frmBoard.Player1OutOfJail(Dice1, Dice2, 1, mPlayerBank)
    
    mDiceTries = 0
    
    frmMessageJail.Visible = False

End Sub

Private Sub Form_Load()

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
    
    With frmBoard.imgPlayer1
        .Picture = LoadPicture(Ap & frmBoard.imgPlayer1.Tag & "\Right1.ico")
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

    frmBoard.imgPlayer1Turn.Visible = True
    frmBoard.imgPlayer2Turn.Visible = False
    
    mPlayerBank = mPlayerBank
    frmBoard.lsbOne.AddItem ("Player 1 used Free Card and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
    Call frmBoard.Player1OutOfJail(Dice1, Dice2, 1, mPlayerBank)
    
    mDiceTries = 0
    
    frmMessageJail.Visible = False

    imgCCOOJ.Visible = False

End Sub

Private Sub imgCOOJ_Click()
    Dim Dice1 As Integer
    Dim Dice2 As Integer
    
    With frmBoard.imgPlayer1
        .Picture = LoadPicture(Ap & frmBoard.imgPlayer1.Tag & "\Right1.ico")
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

    frmBoard.imgPlayer1Turn.Visible = True
    frmBoard.imgPlayer2Turn.Visible = False
    
    mPlayerBank = mPlayerBank
    frmBoard.lsbOne.AddItem ("Player 1 used Free Card and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
    Call frmBoard.Player1OutOfJail(Dice1, Dice2, 1, mPlayerBank)
    
    mDiceTries = 0
    
    frmMessageJail.Visible = False

    imgCOOJ.Visible = False

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
        
        With frmBoard.imgPlayer1
            .Picture = LoadPicture(Ap & frmBoard.imgPlayer1.Tag & "\Right1.ico")
            .Left = 164
            .Top = 0
            .ToolTipText = 9
        End With

        With frmBoard.imgDice
            .Item(0).Picture = LoadPicture(Ap & "dice\" & mDice1 & ".ico")
            .Item(1).Picture = LoadPicture(Ap & "dice\" & mDice2 & ".ico")
        End With

        frmBoard.imgPlayer1Turn.Visible = True
        frmBoard.imgPlayer2Turn.Visible = False
        frmBoard.lsbOne.AddItem ("Player 1 threw a double and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
        Call frmBoard.Player1OutOfJail(mDice1, mDice2, 1, mPlayerBank)
    
        frmMessageJail.Visible = False
    Else
        If mDiceTries = 3 Then
            
            mDiceTries = 0
            
            With frmBoard.imgPlayer1
                .Picture = LoadPicture(Ap & frmBoard.imgPlayer1.Tag & "\Right1.ico")
                .Left = 164
                .Top = 0
                .ToolTipText = 9
            End With

            With frmBoard.imgDice
                .Item(0).Picture = LoadPicture(Ap & "dice\" & mDice1 & ".ico")
                .Item(1).Picture = LoadPicture(Ap & "dice\" & mDice2 & ".ico")
            End With

            frmBoard.imgPlayer1Turn.Visible = True
            frmBoard.imgPlayer2Turn.Visible = False
    
            mPlayerBank = mPlayerBank - 5000
            frmBoard.lsbOne.AddItem ("Player 1 paid the R 5'000.00 and is Out of Jail"), frmBoard.lsbOne.ListCount = 0
            Call frmBoard.Player1OutOfJail(mDice1, mDice2, 1, mPlayerBank)
    
            frmMessageJail.Visible = False
        Else
            msg = MsgBox("Sorry, you have to stay in Jail for another turn.", vbOKOnly, "No Double")
            frmBoard.imgPlayer1Turn.Visible = False
            frmBoard.imgPlayer2Turn.Visible = True
            frmMessageJail.Visible = False
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
