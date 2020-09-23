VERSION 5.00
Begin VB.Form frmCommunityChest 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1395
      Top             =   510
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1395
      Top             =   45
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
      Left            =   1500
      TabIndex        =   0
      Top             =   60
      Width           =   4905
   End
   Begin VB.Image Image1 
      Height          =   1890
      Left            =   90
      Picture         =   "frmCommunityChest.frx":0000
      Top             =   90
      Width           =   1260
   End
End
Attribute VB_Name = "frmCommunityChest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim mCommunityChestCard(1 To 16) As String
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
    Dim x As Integer

    frmCommunityChest.BackColor = RGB(211, 195, 147)

    For x = 1 To 16
        mCard(x) = x
    Next x
    
    mCardCount = 1
    
    mCommunityChestCard(1) = "YOU INHERIT R10'000"
    mCommunityChestCard(2) = "PAY HOSPITAL FEES OF R10'000"
    mCommunityChestCard(3) = "ANNUITY MATURES. COLLECT R10'000"
    mCommunityChestCard(4) = "RECEIVE INTEREST ON 7% PREFERENCE SHARES OF R2'500"
    mCommunityChestCard(5) = "GET OUT OF JAIL FREE"
    mCommunityChestCard(6) = "IT IS YOUR BIRTHDAY COLLECT R1'000 FROM EACH PLAYER"
    mCommunityChestCard(7) = "DOCTOR'S FEE. PAY R5'000"
    mCommunityChestCard(8) = "GO TO MITCHELLS PLAIN. DO NOT COLLECT R20'000.00 past GO"
    mCommunityChestCard(9) = "YOU HAVE WON SECOND PRIZE IN A BEAUTY CONTEST. COLLECT R1'000"
    mCommunityChestCard(10) = "BANK ERROR IN YOUR FAVOUR. COLLECT R10'000"
    mCommunityChestCard(11) = "ADVANCE TO GO"
    mCommunityChestCard(12) = "PAY YOUR INSURANCE PREMIUM OF R5'000"
    mCommunityChestCard(13) = "INCOME TAX REFUND. COLLECT R2'000"
    mCommunityChestCard(14) = "PAY A R1'000 FINE FOR STUPIDITY"
    mCommunityChestCard(15) = "GO TO JAIL"
    mCommunityChestCard(16) = "FOR SALE OF STOCK YOU GET R5'000"
    
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
        frmShuffleCommunityChest.Timer1.Enabled = True
        frmShuffleCommunityChest.Visible = True
    Else
        Call ShowCard
    End If
    
End Sub

Public Function ShowCard()
            
    lblCard.Caption = mCommunityChestCard(mCard(mCardCount))
    
    CardControl = mCard(mCardCount)
    
    mCardCount = mCardCount + 1
    
    Image1.Enabled = False
    
    Timer1.Enabled = True
        
End Function

Public Function CommunityChestInfo(Player As Integer)

    mPlayer = Player
    
    If Player = 2 Then
        Timer2.Enabled = True
    End If

End Function

Private Sub Timer1_Timer()
    
    Call frmBoard.CommunityChestCont(mPlayer, CardControl)
    Timer1.Enabled = False
    lblCard.Caption = ""
    Image1.Enabled = True
    frmCommunityChest.Visible = False
    
End Sub

Private Sub Timer2_Timer()

    Call Image1_Click
    Timer2.Enabled = False
    
End Sub
