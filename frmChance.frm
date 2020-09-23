VERSION 5.00
Begin VB.Form frmChance 
   BorderStyle     =   0  'None
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1305
      Top             =   465
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1305
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   1890
      Left            =   0
      Picture         =   "frmChance.frx":0000
      Top             =   45
      Width           =   1260
   End
   Begin VB.Label lblCard 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   1410
      TabIndex        =   0
      Top             =   15
      Width           =   4905
   End
End
Attribute VB_Name = "frmChance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim mChanceCard(1 To 16) As String
    Dim mCard(1 To 16) As Integer
    Dim mCardCount As Integer
    Dim mPlayer As Integer
    Dim mBank As Long
    Dim CardControl As Integer
    
Private Sub cmdPay_Click()

    Image1.Enabled = False
    
    Timer1.Enabled = True

End Sub

Private Sub Form_Load()
    Dim X As Integer

    frmChance.BackColor = RGB(207, 160, 137)
    
    Call Randomize

    For X = 1 To 16
        mCard(X) = X
    Next X
    
    mCardCount = 1
    
    mChanceCard(1) = "YOUR BUILDING LOAN MATURES. RECEIVE R 15'000"
    mChanceCard(2) = "ADVANCE TO WESTVILLE. IF YOU PASS GO COLLECT R 20'000"
    mChanceCard(3) = "BANK PAYS YOU DIVIDEND OF R 5'000"
    mChanceCard(4) = "SPEEDING FINE. R1'500"
    mChanceCard(5) = "ADVANCE TO KNYSNA. IF YOU PASS GO COLLECT R 20'000"
    mChanceCard(6) = "ADVANCE TO UMHLANGA ROCKS. IF YOU PASS GO COLLECT R 20'000"
    mChanceCard(7) = "YOU ARE ASSESSED FOR STREET REPAIRS. R 4'000 PER HOUSE. R 11'500 PER HOTEL"
    mChanceCard(8) = "PAY SCHOOL FEES OF R 15'000"
    mChanceCard(9) = "GO FORWARD THIRTY SIX SPACES"
    mChanceCard(10) = "TAKE A TRIP TO JHB INTL. AND IF YOU PASS GO COLLECT R 20'000"
    mChanceCard(11) = "YOU HAVE WON A CROSSWORD COMPETITION. COLLECT R 10'000"
    mChanceCard(12) = "GO TO JAIL"
    mChanceCard(13) = "ADVANCE TO GO"
    mChanceCard(14) = "DRUNK IN CHARGE. FINE OF R 2'000"
    mChanceCard(15) = "MAKE GENERAL REPAIRS ON ALL YOUR HOUSES. R 2'500 PER HOUSE. R 10'000 PER HOTEL"
    mChanceCard(16) = "GET OUT OF JAIL FREE"
    
    Call Shuffle
    
End Sub

Private Sub Shuffle()
    Dim Card As Integer, row As Integer
    
    Call ZeroDeckArray
    Call Randomize
    
    For Card = 1 To 16
        Do
            row = 1 + Int(Rnd() * UBound(mCard))
        Loop While mCard(row) <> 0
        
        mCard(row) = Card
    Next Card

End Sub

Private Sub ZeroDeckArray()
    Dim row As Integer
    
    For row = LBound(mCard) To UBound(mCard)
        mCard(row) = 0
    Next row
    
End Sub

Private Sub Image1_Click()

    If mCard(mCardCount) = 17 Then
        Call Shuffle
        mCardCount = 1
        frmShuffleChance.Timer1.Enabled = True
        frmShuffleChance.Visible = True
    Else
        Call ShowCard
    End If
    
End Sub

Public Function ShowCard()
    
    lblCard.Caption = mChanceCard(mCard(mCardCount))
    
    CardControl = mCard(mCardCount)
    
    mCardCount = mCardCount + 1
    
    Image1.Enabled = False
    
    Timer1.Enabled = True
        
End Function

Public Function ChanceInfo(Player As Integer)

    mPlayer = Player
    
    If Player = 2 Then
        Timer2.Enabled = True
    End If

End Function

Private Sub Timer1_Timer()
    
    Call frmBoard.ChanceCont(mPlayer, CardControl)
    Timer1.Enabled = False
    lblCard.Caption = ""
    Image1.Enabled = True
    frmChance.Visible = False
    
End Sub

Private Sub Timer2_Timer()

    Call Image1_Click
    Timer2.Enabled = False
    
End Sub

