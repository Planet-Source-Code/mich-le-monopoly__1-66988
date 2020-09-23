VERSION 5.00
Begin VB.Form frmBuyHouses 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   Picture         =   "frmBuyHouses.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrExit 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6420
      Top             =   3060
   End
   Begin VB.Timer tmrBlue 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1830
      Top             =   2025
   End
   Begin VB.Timer tmrGreen 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   2025
   End
   Begin VB.Timer tmrYellow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4185
      Top             =   2025
   End
   Begin VB.Timer tmrRed 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4185
      Top             =   990
   End
   Begin VB.Timer tmrLight_Brown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2985
      Top             =   975
   End
   Begin VB.Timer tmrPurple 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1830
      Top             =   960
   End
   Begin VB.Timer tmrLight_Blue 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   960
   End
   Begin VB.Timer tmrBrown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   630
      Top             =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6015
      MaskColor       =   &H00FF0000&
      TabIndex        =   5
      Top             =   5205
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6015
      MaskColor       =   &H00FF0000&
      TabIndex        =   4
      Top             =   4575
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPurchase 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Purchase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6015
      MaskColor       =   &H00FF0000&
      TabIndex        =   3
      Top             =   3975
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgPurchaseHouse 
      Height          =   480
      Index           =   0
      Left            =   825
      Tag             =   "0"
      Top             =   3450
      Width           =   480
   End
   Begin VB.Image imgPurchaseHouse 
      Height          =   480
      Index           =   2
      Left            =   4845
      Tag             =   "0"
      Top             =   3450
      Width           =   480
   End
   Begin VB.Image imgPurchaseHouse 
      Height          =   480
      Index           =   1
      Left            =   2835
      Tag             =   "0"
      Top             =   3450
      Width           =   480
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   6150
      TabIndex        =   6
      Top             =   465
      Width           =   1185
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   270
      X2              =   270
      Y1              =   227
      Y2              =   378
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   135
      X2              =   135
      Y1              =   227
      Y2              =   378
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   11
      Left            =   4485
      Picture         =   "frmBuyHouses.frx":92802
      Tag             =   "\HOMESUP\1.ico"
      Top             =   4335
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   12
      Left            =   5190
      Picture         =   "frmBuyHouses.frx":92B0C
      Tag             =   "\HOMESUP\2.ico"
      Top             =   4335
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   14
      Left            =   5190
      Picture         =   "frmBuyHouses.frx":92E16
      Tag             =   "\HOMESUP\4.ico"
      Top             =   4845
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   13
      Left            =   4485
      Picture         =   "frmBuyHouses.frx":93120
      Tag             =   "\HOMESUP\3.ico"
      Top             =   4860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   15
      Left            =   4875
      Picture         =   "frmBuyHouses.frx":9342A
      Tag             =   "\HOMESUP\5.ico"
      Top             =   5355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   6
      Left            =   2490
      Picture         =   "frmBuyHouses.frx":93734
      Tag             =   "\HOMESUP\1.ico"
      Top             =   4350
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   7
      Left            =   3195
      Picture         =   "frmBuyHouses.frx":93A3E
      Tag             =   "\HOMESUP\2.ico"
      Top             =   4350
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   9
      Left            =   3195
      Picture         =   "frmBuyHouses.frx":93D48
      Tag             =   "\HOMESUP\4.ico"
      Top             =   4860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   8
      Left            =   2490
      Picture         =   "frmBuyHouses.frx":94052
      Tag             =   "\HOMESUP\3.ico"
      Top             =   4875
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   10
      Left            =   2880
      Picture         =   "frmBuyHouses.frx":9435C
      Tag             =   "\HOMESUP\5.ico"
      Top             =   5370
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   5
      Left            =   810
      Picture         =   "frmBuyHouses.frx":94666
      Tag             =   "\HOMESUP\5.ico"
      Top             =   5385
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   3
      Left            =   420
      Picture         =   "frmBuyHouses.frx":94970
      Tag             =   "\HOMESUP\3.ico"
      Top             =   4890
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblProp3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4410
      TabIndex        =   2
      Top             =   3915
      Width           =   1350
   End
   Begin VB.Label lblProp2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   1
      Top             =   3900
      Width           =   1350
   End
   Begin VB.Label lblProp1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   405
      TabIndex        =   0
      Top             =   3915
      Width           =   1350
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   4
      Left            =   1125
      Picture         =   "frmBuyHouses.frx":94C7A
      Tag             =   "\HOMESUP\4.ico"
      Top             =   4875
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   2
      Left            =   1125
      Picture         =   "frmBuyHouses.frx":94F84
      Tag             =   "\HOMESUP\2.ico"
      Top             =   4365
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   1
      Left            =   420
      Picture         =   "frmBuyHouses.frx":9528E
      Tag             =   "\HOMESUP\1.ico"
      Top             =   4365
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBlue 
      Height          =   540
      Left            =   2250
      Picture         =   "frmBuyHouses.frx":95598
      Tag             =   "Blue.gif"
      Top             =   1950
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgGreen 
      Height          =   540
      Left            =   3420
      Picture         =   "frmBuyHouses.frx":9571C
      Tag             =   "Green.gif"
      Top             =   1950
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgYellow 
      Height          =   540
      Left            =   4620
      Picture         =   "frmBuyHouses.frx":958F7
      Tag             =   "Yellow.gif"
      Top             =   1980
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgRed 
      Height          =   540
      Left            =   4620
      Picture         =   "frmBuyHouses.frx":95B32
      Tag             =   "Red.gif"
      Top             =   900
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgLight_Brown 
      Height          =   540
      Left            =   3420
      Picture         =   "frmBuyHouses.frx":95D71
      Tag             =   "Light Brown.gif"
      Top             =   900
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgPurple 
      Height          =   540
      Left            =   2250
      Picture         =   "frmBuyHouses.frx":95FB6
      Tag             =   "Purple.gif"
      Top             =   900
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgLight_Blue 
      Height          =   540
      Left            =   1035
      Picture         =   "frmBuyHouses.frx":961BE
      Tag             =   "Light Blue.gif"
      Top             =   900
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgBrown 
      Height          =   540
      Left            =   1050
      Picture         =   "frmBuyHouses.frx":9640A
      Tag             =   "Brown.gif"
      Top             =   1950
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgProp3 
      Height          =   495
      Left            =   4425
      Picture         =   "frmBuyHouses.frx":966B2
      Stretch         =   -1  'True
      Top             =   3420
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgProp1 
      Height          =   495
      Left            =   405
      Picture         =   "frmBuyHouses.frx":9695A
      Stretch         =   -1  'True
      Top             =   3420
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgProp2 
      Height          =   495
      Left            =   2400
      Picture         =   "frmBuyHouses.frx":96C02
      Stretch         =   -1  'True
      Top             =   3420
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmBuyHouses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    
    Dim Ap As String
    Dim INIfile As String
    Dim RetVal As Long

    Dim mPlayerBank As Long
    Dim mPlayer As Integer
    Dim mDead(1 To 3) As Integer
    
    Dim Timercount As Integer
        
Private Sub cmdCancel_Click()

    frmBuyHouses.Visible = False
    Call ResetForm

End Sub

Private Sub cmdClear_Click()

    Call ResetForm

End Sub

Private Sub Form_Load()
    If Right(App.Path, 1) = "\" Then
        Ap = App.Path
    Else
        Ap = App.Path & "\"
    End If

    imgPurchaseHouse(0).Tag = 0
    imgPurchaseHouse(1).Tag = 0
    imgPurchaseHouse(2).Tag = 0
    
End Sub

Private Sub imgBlue_Click()
    Dim X As Integer

    For X = 1 To 15
        imgHouse(X).Tag = 20000
    Next X
    
    imgProp1.Picture = LoadPicture(Ap & imgBlue.Tag)
    lblProp1.Caption = "Franschoek"
    imgProp2.Picture = LoadPicture(Ap & imgBlue.Tag)
    lblProp2.Caption = "Clifton"
    imgProp3.Picture = LoadPicture(Ap & imgBlue.Tag)
    mDead(1) = 36
    mDead(2) = 38
    
    If frmBoard.imgHomes(36).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(36).Tag)
        imgPurchaseHouse(0).Tag = frmBoard.imgHomes(36).Tag
    End If

    If frmBoard.imgHomes(38).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(38).Tag + 5)
        imgPurchaseHouse(1).Tag = frmBoard.imgHomes(38).Tag
    End If

    imgProp1.Visible = True
    lblProp1.Visible = True
    imgProp2.Visible = True
    lblProp2.Visible = True
    imgProp3.Visible = False
    lblProp3.Visible = False
    Line1.Visible = True
    Line2.Visible = True


End Sub

Private Sub imgBrown_Click()
    Dim X As Integer

    For X = 1 To 15
        imgHouse(X).Tag = 5000
    Next X
    
    imgProp1.Picture = LoadPicture(Ap & imgBrown.Tag)
    lblProp1.Caption = "Westville"
    imgProp2.Picture = LoadPicture(Ap & imgBrown.Tag)
    lblProp2.Caption = "Amanzimtoti"
    mDead(1) = 0
    mDead(2) = 2
    
    If frmBoard.imgHomes(0).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(0).Tag)
        imgPurchaseHouse(0).Tag = frmBoard.imgHomes(0).Tag
    End If

    If frmBoard.imgHomes(2).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(2).Tag + 5)
        imgPurchaseHouse(1).Tag = frmBoard.imgHomes(2).Tag
    End If
    
    imgProp1.Visible = True
    lblProp1.Visible = True
    imgProp2.Visible = True
    lblProp2.Visible = True
    imgProp3.Visible = False
    lblProp3.Visible = False
    Line1.Visible = True
    Line2.Visible = True

End Sub

Private Sub imgBrown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgBrown.BorderStyle = 1

End Sub

Private Sub imgBrown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgBrown.BorderStyle = 0

End Sub

Private Sub imgGreen_Click()
    Dim X As Integer

    For X = 1 To 15
        imgHouse(X).Tag = 20000
    Next X
    
    imgProp1.Picture = LoadPicture(Ap & imgGreen.Tag)
    lblProp1.Caption = "Tyger Valley"
    imgProp2.Picture = LoadPicture(Ap & imgGreen.Tag)
    lblProp2.Caption = "Mitchells Plain"
    imgProp3.Picture = LoadPicture(Ap & imgGreen.Tag)
    lblProp3.Caption = "Blouberg Strand"
    mDead(1) = 30
    mDead(2) = 31
    mDead(3) = 33
    
    If frmBoard.imgHomes(30).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(30).Tag)
        imgPurchaseHouse(0).Tag = frmBoard.imgHomes(30).Tag
    End If

    If frmBoard.imgHomes(31).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(31).Tag + 5)
        imgPurchaseHouse(1).Tag = frmBoard.imgHomes(31).Tag
    End If

    If frmBoard.imgHomes(33).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(33).Tag + 10)
        imgPurchaseHouse(2).Tag = frmBoard.imgHomes(33).Tag
    End If
    
    imgProp1.Visible = True
    lblProp1.Visible = True
    imgProp2.Visible = True
    lblProp2.Visible = True
    imgProp3.Visible = True
    lblProp3.Visible = True
    Line1.Visible = True
    Line2.Visible = True

End Sub

Private Sub imgLight_Blue_Click()
    Dim X As Integer

    For X = 1 To 15
        imgHouse(X).Tag = 5000
    Next X
    
    imgProp1.Picture = LoadPicture(Ap & imgLight_Blue.Tag)
    lblProp1.Caption = "Umhlanga Rocks"
    imgProp2.Picture = LoadPicture(Ap & imgLight_Blue.Tag)
    lblProp2.Caption = "Ballito Bay"
    imgProp3.Picture = LoadPicture(Ap & imgLight_Blue.Tag)
    lblProp3.Caption = "La Lucia"
    mDead(1) = 5
    mDead(2) = 7
    mDead(3) = 8
    
    imgPurchaseHouse(0).Tag = 0
    imgPurchaseHouse(1).Tag = 0
    imgPurchaseHouse(2).Tag = 0
    
    If frmBoard.imgHomes(5).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(5).Tag)
        imgPurchaseHouse(0).Tag = frmBoard.imgHomes(5).Tag
    End If

    If frmBoard.imgHomes(7).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(7).Tag + 5)
        imgPurchaseHouse(1).Tag = frmBoard.imgHomes(7).Tag
    End If

    If frmBoard.imgHomes(8).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(8).Tag + 10)
        imgPurchaseHouse(2).Tag = frmBoard.imgHomes(8).Tag
    End If
    
    imgProp1.Visible = True
    lblProp1.Visible = True
    imgProp2.Visible = True
    lblProp2.Visible = True
    imgProp3.Visible = True
    lblProp3.Visible = True
    Line1.Visible = True
    Line2.Visible = True

End Sub

Private Sub imgLight_Blue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgLight_Blue.BorderStyle = 1

End Sub

Private Sub imgLight_Blue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgLight_Blue.BorderStyle = 0

End Sub

Private Sub imgLight_Brown_Click()
    Dim X As Integer

    For X = 1 To 15
        imgHouse(X).Tag = 10000
    Next X
    
    imgProp1.Picture = LoadPicture(Ap & imgLight_Brown.Tag)
    lblProp1.Caption = "Wilderness"
    imgProp2.Picture = LoadPicture(Ap & imgLight_Brown.Tag)
    lblProp2.Caption = "Knysna"
    imgProp3.Picture = LoadPicture(Ap & imgLight_Brown.Tag)
    lblProp3.Caption = "Plettenberg Bay"
    mDead(1) = 15
    mDead(2) = 17
    mDead(3) = 18
    
    If frmBoard.imgHomes(15).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(15).Tag)
        imgPurchaseHouse(0).Tag = frmBoard.imgHomes(15).Tag
    End If

    If frmBoard.imgHomes(17).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(17).Tag + 5)
        imgPurchaseHouse(1).Tag = frmBoard.imgHomes(17).Tag
    End If

    If frmBoard.imgHomes(18).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(18).Tag + 10)
        imgPurchaseHouse(2).Tag = frmBoard.imgHomes(18).Tag
    End If

    imgProp1.Visible = True
    lblProp1.Visible = True
    imgProp2.Visible = True
    lblProp2.Visible = True
    imgProp3.Visible = True
    lblProp3.Visible = True
    Line1.Visible = True
    Line2.Visible = True

End Sub

Private Sub imgLight_Brown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgLight_Brown.BorderStyle = 1

End Sub

Private Sub imgLight_Brown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgLight_Brown.BorderStyle = 0

End Sub

Private Sub imgPurple_Click()
    Dim X As Integer

    For X = 1 To 15
        imgHouse(X).Tag = 10000
    Next X
    
    imgProp1.Picture = LoadPicture(Ap & imgPurple.Tag)
    lblProp1.Caption = "Menlyn Park"
    imgProp2.Picture = LoadPicture(Ap & imgPurple.Tag)
    lblProp2.Caption = "Port Elizabeth"
    imgProp3.Picture = LoadPicture(Ap & imgPurple.Tag)
    lblProp3.Caption = "Waterkloof"
    mDead(1) = 10
    mDead(2) = 12
    mDead(3) = 13
    
    imgPurchaseHouse(0).Tag = 0
    imgPurchaseHouse(1).Tag = 0
    imgPurchaseHouse(2).Tag = 0
    
    If frmBoard.imgHomes(10).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(10).Tag)
        imgPurchaseHouse(0).Tag = frmBoard.imgHomes(10).Tag
    End If

    If frmBoard.imgHomes(12).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(12).Tag + 5)
        imgPurchaseHouse(1).Tag = frmBoard.imgHomes(12).Tag
    End If

    If frmBoard.imgHomes(13).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(13).Tag + 10)
        imgPurchaseHouse(2).Tag = frmBoard.imgHomes(13).Tag
    End If
        
    imgProp1.Visible = True
    lblProp1.Visible = True
    imgProp2.Visible = True
    lblProp2.Visible = True
    imgProp3.Visible = True
    lblProp3.Visible = True
    Line1.Visible = True
    Line2.Visible = True

End Sub

Private Sub imgPurple_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgPurple.BorderStyle = 1

End Sub

Private Sub imgPurple_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgPurple.BorderStyle = 0

End Sub

Private Sub imgRed_Click()
    Dim X As Integer

    For X = 1 To 15
        imgHouse(X).Tag = 15000
    Next X
    
    imgProp1.Picture = LoadPicture(Ap & imgRed.Tag)
    lblProp1.Caption = "Soweto"
    imgProp2.Picture = LoadPicture(Ap & imgRed.Tag)
    lblProp2.Caption = "Hillbrow"
    imgProp3.Picture = LoadPicture(Ap & imgRed.Tag)
    lblProp3.Caption = "Boksburg"
    mDead(1) = 20
    mDead(2) = 22
    mDead(3) = 23

    If frmBoard.imgHomes(20).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(20).Tag)
        imgPurchaseHouse(0).Tag = frmBoard.imgHomes(20).Tag
    End If

    If frmBoard.imgHomes(22).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(22).Tag + 5)
        imgPurchaseHouse(1).Tag = frmBoard.imgHomes(22).Tag
    End If

    If frmBoard.imgHomes(23).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(23).Tag + 10)
        imgPurchaseHouse(2).Tag = frmBoard.imgHomes(23).Tag
    End If

    imgProp1.Visible = True
    lblProp1.Visible = True
    imgProp2.Visible = True
    lblProp2.Visible = True
    imgProp3.Visible = True
    lblProp3.Visible = True
    Line1.Visible = True
    Line2.Visible = True

End Sub

Private Sub imgRed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgRed.BorderStyle = 1

End Sub

Private Sub imgRed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgRed.BorderStyle = 0

End Sub

Private Sub imgYellow_Click()
    Dim X As Integer

    For X = 1 To 15
        imgHouse(X).Tag = 15000
    Next X
    
    imgProp1.Picture = LoadPicture(Ap & imgYellow.Tag)
    lblProp1.Caption = "Randburg"
    imgProp2.Picture = LoadPicture(Ap & imgYellow.Tag)
    lblProp2.Caption = "Sandton"
    imgProp3.Picture = LoadPicture(Ap & imgYellow.Tag)
    lblProp3.Caption = "Hyde Park"
    mDead(1) = 25
    mDead(2) = 26
    mDead(3) = 28

    If frmBoard.imgHomes(25).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(25).Tag)
        imgPurchaseHouse(0).Tag = frmBoard.imgHomes(25).Tag
    End If

    If frmBoard.imgHomes(26).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(26).Tag + 5)
        imgPurchaseHouse(1).Tag = frmBoard.imgHomes(26).Tag
    End If

    If frmBoard.imgHomes(28).Tag > 0 Then
        Call imgHouse_Click(frmBoard.imgHomes(28).Tag + 10)
        imgPurchaseHouse(2).Tag = frmBoard.imgHomes(28).Tag
    End If
    
    imgProp1.Visible = True
    lblProp1.Visible = True
    imgProp2.Visible = True
    lblProp2.Visible = True
    imgProp3.Visible = True
    lblProp3.Visible = True
    Line1.Visible = True
    Line2.Visible = True

End Sub

Private Sub imgYellow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgYellow.BorderStyle = 1

End Sub

Private Sub imgYellow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgYellow.BorderStyle = 0

End Sub

Private Sub imgGreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgGreen.BorderStyle = 1

End Sub

Private Sub imgGreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgGreen.BorderStyle = 0

End Sub

Private Sub imgBlue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgBlue.BorderStyle = 1

End Sub

Private Sub imgBlue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgBlue.BorderStyle = 0

End Sub

Private Sub imgProp1_Click()
    Dim X As Integer

    If imgProp1.BorderStyle = 0 Then
        imgProp1.BorderStyle = 1
        For X = 1 To 5
            imgHouse(X).Visible = True
        Next X
    Else
        imgProp1.BorderStyle = 0
        For X = 1 To 5
            imgHouse(X).Visible = False
        Next X
    End If
    
End Sub

Private Sub imgProp2_Click()
    Dim X As Integer

    If imgProp2.BorderStyle = 0 Then
        imgProp2.BorderStyle = 1
        For X = 6 To 10
            imgHouse(X).Visible = True
        Next X
    Else
        imgProp2.BorderStyle = 0
        For X = 6 To 10
            imgHouse(X).Visible = False
        Next X
    End If

End Sub

Private Sub imgProp3_Click()
    Dim X As Integer

    If imgProp3.BorderStyle = 0 Then
        imgProp3.BorderStyle = 1
        For X = 11 To 15
            imgHouse(X).Visible = True
        Next X
    Else
        imgProp3.BorderStyle = 0
        For X = 11 To 15
            imgHouse(X).Visible = False
        Next X
    End If

End Sub

Private Sub imgHouse_Click(Index As Integer)
        
    If Index <= 5 Then
        imgPurchaseHouse(0).Picture = LoadPicture(Ap & "HOMESUP\" & Index & ".ico")
        imgPurchaseHouse(0).WhatsThisHelpID = imgHouse(Index).Tag
        imgPurchaseHouse(0).Visible = True
        mPlayerBank = mPlayerBank - imgHouse(Index).Tag
        imgPurchaseHouse(0).Tag = Index
        If imgProp3.Visible = False Then
            imgPurchaseHouse(2).Tag = Index
            imgPurchaseHouse(2).Visible = False
        End If
    End If
    
    If Index > 5 And Index <= 10 Then
        imgPurchaseHouse(1).Picture = LoadPicture(Ap & "HOMESUP\" & Int(Index - 5) & ".ico")
        imgPurchaseHouse(0).WhatsThisHelpID = imgHouse(Index).Tag
        imgPurchaseHouse(1).Visible = True
        mPlayerBank = mPlayerBank - imgHouse(Index).Tag
        imgPurchaseHouse(1).Tag = Int(Index - 5)
    End If
    
    If Index > 10 And Index <= 15 Then
        imgPurchaseHouse(2).Picture = LoadPicture(Ap & "HOMESUP\" & Int(Index - 10) & ".ico")
        imgPurchaseHouse(0).WhatsThisHelpID = imgHouse(Index).Tag
        imgPurchaseHouse(2).Visible = True
        mPlayerBank = mPlayerBank - imgHouse(Index).Tag
        imgPurchaseHouse(2).Tag = Int(Index - 10)
    End If
       
    cmdClear.Visible = True
    cmdPurchase.Visible = True
        
End Sub

Private Sub cmdPurchase_Click()
    Dim Prop1 As Integer, Prop2 As Integer, Prop3 As Integer
    Dim X As Integer
    
    If imgPurchaseHouse(0).Tag = "" Then
        Prop1 = 0
    Else
        Prop1 = Int(imgPurchaseHouse(0).Tag)
    End If
    
    If imgPurchaseHouse(1).Tag = "" Then
        Prop2 = 0
    Else
        Prop2 = Int(imgPurchaseHouse(1).Tag)
    End If
        
    If imgPurchaseHouse(2).Tag = "" Then
        Prop3 = 0
    Else
        Prop3 = Int(imgPurchaseHouse(2).Tag)
    End If
    
    ''''''''''''''''''''''CHECK IF PROP PURCHASE APPLIES TO RULES ''''''''''''''''''
    If Prop1 = 1 Then
        If Prop2 >= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 >= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop1 = 2 Then
        If Prop2 >= 4 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 >= 4 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop1 = 3 Then
        If Prop2 = 5 Or Prop2 = 1 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 = 5 Or Prop3 = 1 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop1 = 4 Then
        If Prop2 <= 2 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 <= 2 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop1 = 5 Then
        If Prop2 <= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 <= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    End If
    
    If Prop2 = 1 Then
        If Prop1 >= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 >= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop2 = 2 Then
        If Prop1 >= 4 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 >= 4 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop2 = 3 Then
        If Prop1 = 5 Or Prop1 = 1 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 = 5 Or Prop3 = 1 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop2 = 4 Then
        If Prop1 <= 2 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 <= 2 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop2 = 5 Then
        If Prop1 <= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop3 <= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    End If
    
    If Prop3 = 1 Then
        If Prop2 >= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop1 >= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop3 = 2 Then
        If Prop2 >= 4 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop1 >= 4 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop3 = 3 Then
        If Prop2 = 5 Or Prop2 = 1 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop1 = 5 Or Prop1 = 1 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop3 = 4 Then
        If Prop2 <= 2 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop1 <= 2 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    ElseIf Prop3 = 5 Then
        If Prop2 <= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
        If Prop1 <= 3 Then
            X = MsgBox("Invalid purchase. Your properties must be equal, one greater or one less house for all properties.", vbOKOnly, "Purchase error.")
            Exit Sub
        End If
    End If
    
    '''''''''''''''''''''''''''''' CHECK DONE '''''''''''''''''''''''''''''''''''''''''
    
    If mPlayerBank < 0 Then
        X = MsgBox("Invalid purchase. You do not have enough money.", vbOKOnly, "Purchase error.")
    Else
        Call frmBoard.SetHouses(Prop1, Prop2, Prop3, mPlayer, mDead(1), mDead(2), mDead(3))
        Call ResetForm
        frmBuyHouses.Visible = False
    End If
       
End Sub

Public Function ResetForm()
    Dim X As Integer

    imgProp1.Visible = False
    imgProp2.Visible = False
    imgProp3.Visible = False
    imgProp1.BorderStyle = 0
    imgProp2.BorderStyle = 0
    imgProp3.BorderStyle = 0
    lblProp1.Visible = False
    lblProp2.Visible = False
    lblProp3.Visible = False
    Line1.Visible = False
    Line2.Visible = False
        
    For X = 1 To 15
        imgHouse(X).Visible = False
    Next X
    
    For X = 0 To 2
        imgPurchaseHouse(X).Visible = False
        imgPurchaseHouse(X).Tag = 0
    Next X

End Function

Public Function InitialSettings(Player As Integer, Bank As Long)
    Dim PropCount As Integer
    
    mPlayer = Player
    mPlayerBank = Bank
    
    lblPlayer.Caption = "Player " & mPlayer
    
    imgBrown.Visible = False
    imgLight_Blue.Visible = False
    imgPurple.Visible = False
    imgLight_Brown.Visible = False
    imgRed.Visible = False
    imgYellow.Visible = False
    imgGreen.Visible = False
    imgBlue.Visible = False
    
    If frmBoard.imgDead(0).Tag = Player And frmBoard.imgDead(2).Tag = Player Then
        imgBrown.Visible = True
    End If
    If frmBoard.imgDead(5).Tag = Player And frmBoard.imgDead(7).Tag = Player And frmBoard.imgDead(8).Tag = Player Then
        imgLight_Blue.Visible = True
    End If
    If frmBoard.imgDead(10).Tag = Player And frmBoard.imgDead(12).Tag = Player And frmBoard.imgDead(13).Tag = Player Then
        imgPurple.Visible = True
    End If
    If frmBoard.imgDead(15).Tag = Player And frmBoard.imgDead(17).Tag = Player And frmBoard.imgDead(18).Tag = Player Then
        imgLight_Brown.Visible = True
    End If
    If frmBoard.imgDead(20).Tag = Player And frmBoard.imgDead(22).Tag = Player And frmBoard.imgDead(23).Tag = Player Then
        imgRed.Visible = True
    End If
    If frmBoard.imgDead(25).Tag = Player And frmBoard.imgDead(26).Tag = Player And frmBoard.imgDead(28).Tag = Player Then
        imgYellow.Visible = True
    End If
    If frmBoard.imgDead(30).Tag = Player And frmBoard.imgDead(31).Tag = Player And frmBoard.imgDead(33).Tag = Player Then
        imgGreen.Visible = True
    End If
    If frmBoard.imgDead(36).Tag = Player And frmBoard.imgDead(38).Tag = Player Then
        imgBlue.Visible = True
    End If
    
    If Player = 2 Then
        Call Player2Purchase
    End If
        
End Function

Public Function Player2Purchase()

    If imgBrown.Visible = True Then
        tmrBrown.Enabled = True
        imgBrown.Visible = False
        Exit Function
    End If
    
    If imgLight_Blue.Visible = True Then
        tmrLight_Blue.Enabled = True
        imgLight_Blue.Visible = False
        Exit Function
    End If

    If imgPurple.Visible = True Then
        tmrPurple.Enabled = True
        imgPurple.Visible = False
        Exit Function
    End If

    If imgLight_Brown.Visible = True Then
        tmrLight_Brown.Enabled = True
        imgLight_Brown.Visible = False
        Exit Function
    End If

    If imgRed.Visible = True Then
        tmrRed.Enabled = True
        imgRed.Visible = False
        Exit Function
    End If

    If imgYellow.Visible = True Then
        tmrYellow.Enabled = True
        imgYellow.Visible = False
        Exit Function
    End If

    If imgGreen.Visible = True Then
        tmrGreen.Enabled = True
        imgGreen.Visible = False
        Exit Function
    End If

    If imgBlue.Visible = True Then
        tmrBlue.Enabled = True
        imgBlue.Visible = False
        Exit Function
    End If
    
    If imgBlue.Visible = False Then
        tmrExit.Enabled = True
    End If
    
End Function

Private Sub tmrExit_Timer()

    Timercount = Timercount + 1
    
    If Timercount = 5 Then
        Call ResetForm
    End If
    
    If Timercount = 10 Then
        tmrExit.Enabled = False
        frmBuyHouses.Visible = False
        Timercount = 0
    End If

End Sub

Private Sub tmrBlue_Timer()
    Timercount = Timercount + 1
    
    If Timercount = 10 Then
        frmBuyHouses.Visible = True
        Call imgBlue_Click
        imgProp3.Visible = False
        Call imgProp1_Click
        Call imgProp2_Click
    End If
    
    If Timercount = 20 Then
        If mPlayerBank <= 50000 Then
            Call cmdCancel_Click
            Timercount = 0
            Call Player2Purchase
            tmrBlue.Enabled = False
            Exit Sub
        End If
        If imgPurchaseHouse(0).Tag = 0 Or imgPurchaseHouse(1).Tag = 0 Then
            If imgPurchaseHouse(0).Tag = 0 Then
                If mPlayerBank - imgHouse(1).Tag > 20000 Then
                    Call imgHouse_Click(1)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 0 Then
                If mPlayerBank - imgHouse(6).Tag > 20000 Then
                    Call imgHouse_Click(6)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 1 Or imgPurchaseHouse(1).Tag = 1 Then
            If imgPurchaseHouse(0).Tag = 1 Then
                If mPlayerBank - imgHouse(2).Tag > 20000 Then
                    Call imgHouse_Click(2)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 1 Then
                If mPlayerBank - imgHouse(7).Tag > 20000 Then
                    Call imgHouse_Click(7)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 2 Or imgPurchaseHouse(1).Tag = 2 Then
            If imgPurchaseHouse(0).Tag = 2 Then
                If mPlayerBank - imgHouse(3).Tag > 20000 Then
                    Call imgHouse_Click(3)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 2 Then
                If mPlayerBank - imgHouse(8).Tag > 20000 Then
                    Call imgHouse_Click(8)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 3 Or imgPurchaseHouse(1).Tag = 3 Then
            If imgPurchaseHouse(0).Tag = 3 Then
                If mPlayerBank - imgHouse(4).Tag > 20000 Then
                    Call imgHouse_Click(4)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 3 Then
                If mPlayerBank - imgHouse(9).Tag > 20000 Then
                    Call imgHouse_Click(9)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 4 Or imgPurchaseHouse(1).Tag = 4 Then
            If imgPurchaseHouse(0).Tag = 4 Then
                If mPlayerBank - imgHouse(5).Tag > 20000 Then
                    Call imgHouse_Click(5)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 4 Then
                If mPlayerBank - imgHouse(10).Tag > 20000 Then
                    Call imgHouse_Click(10)
                End If
            End If
        End If
    End If
    
    If Timercount = 30 Then
        Call cmdPurchase_Click
        Timercount = 0
        Call Player2Purchase
        tmrBlue.Enabled = False
    End If
    
End Sub

Private Sub tmrBrown_Timer()
    
    Timercount = Timercount + 1
    
    If Timercount = 10 Then
        frmBuyHouses.Visible = True
        Call imgBrown_Click
        Call imgProp1_Click
        Call imgProp2_Click
    End If
    
    If Timercount = 20 Then
        If mPlayerBank <= 50000 Then
            Call cmdCancel_Click
            Timercount = 0
            Call Player2Purchase
            tmrBrown.Enabled = False
            Exit Sub
        End If
        If imgPurchaseHouse(0).Tag = 0 Or imgPurchaseHouse(1).Tag = 0 Then
            If imgPurchaseHouse(0).Tag = 0 Then
                If mPlayerBank - imgHouse(1).Tag > 20000 Then
                    Call imgHouse_Click(1)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 0 Then
                If mPlayerBank - imgHouse(6).Tag > 20000 Then
                    Call imgHouse_Click(6)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 1 Or imgPurchaseHouse(1).Tag = 1 Then
            If imgPurchaseHouse(0).Tag = 1 Then
                If mPlayerBank - imgHouse(2).Tag > 20000 Then
                    Call imgHouse_Click(2)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 1 Then
                If mPlayerBank - imgHouse(7).Tag > 20000 Then
                    Call imgHouse_Click(7)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 2 Or imgPurchaseHouse(1).Tag = 2 Then
            If imgPurchaseHouse(0).Tag = 2 Then
                If mPlayerBank - imgHouse(3).Tag > 20000 Then
                    Call imgHouse_Click(3)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 2 Then
                If mPlayerBank - imgHouse(8).Tag > 20000 Then
                    Call imgHouse_Click(8)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 3 Or imgPurchaseHouse(1).Tag = 3 Then
            If imgPurchaseHouse(0).Tag = 3 Then
                If mPlayerBank - imgHouse(4).Tag > 20000 Then
                    Call imgHouse_Click(4)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 3 Then
                If mPlayerBank - imgHouse(9).Tag > 20000 Then
                    Call imgHouse_Click(9)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 4 Or imgPurchaseHouse(1).Tag = 4 Then
            If imgPurchaseHouse(0).Tag = 4 Then
                If mPlayerBank - imgHouse(5).Tag > 20000 Then
                    Call imgHouse_Click(5)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 4 Then
                If mPlayerBank - imgHouse(10).Tag > 20000 Then
                    Call imgHouse_Click(10)
                End If
            End If
        End If
    End If
    
    If Timercount = 30 Then
        Call cmdPurchase_Click
        Timercount = 0
        tmrBrown.Enabled = False
        Call Player2Purchase
    End If
    
End Sub

Private Sub tmrGreen_Timer()
    Dim Bank As Long
    
    Timercount = Timercount + 1
    
    If Timercount = 10 Then
        frmBuyHouses.Visible = True
        Call imgGreen_Click
        Call imgProp1_Click
        Call imgProp2_Click
        Call imgProp3_Click
    End If
    
    If Timercount = 20 Then
        If mPlayerBank <= 50000 Then
            Call cmdCancel_Click
            Timercount = 0
            Call Player2Purchase
            tmrGreen.Enabled = False
            Exit Sub
        End If
        If imgPurchaseHouse(0).Tag = 0 Or imgPurchaseHouse(1).Tag = 0 Or imgPurchaseHouse(2).Tag = 0 Then
            If imgPurchaseHouse(0).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(1).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(1)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(6).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(6)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(11).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(11)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 1 Or imgPurchaseHouse(1).Tag = 1 Or imgPurchaseHouse(2).Tag = 1 Then
            If imgPurchaseHouse(0).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(2).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(2)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(7).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(7)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(12).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(12)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 2 Or imgPurchaseHouse(1).Tag = 2 Or imgPurchaseHouse(2).Tag = 2 Then
            If imgPurchaseHouse(0).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(3).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(3)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(8).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(8)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(13).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(13)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 3 Or imgPurchaseHouse(1).Tag = 3 Or imgPurchaseHouse(2).Tag = 3 Then
            If imgPurchaseHouse(0).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(4).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(4)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(9).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(9)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(14).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(14)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 4 Or imgPurchaseHouse(1).Tag = 4 Or imgPurchaseHouse(2).Tag = 4 Then
            If imgPurchaseHouse(0).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(5).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(5)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(10).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(10)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(15).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(15)
                End If
            End If
        End If
    End If
    
    If Timercount = 30 Then
        Call cmdPurchase_Click
        Timercount = 0
        Call Player2Purchase
        tmrGreen.Enabled = False
    End If
    

End Sub

Private Sub tmrLight_Blue_Timer()
    Dim Bank As Long
    
    Timercount = Timercount + 1
    
    If Timercount = 10 Then
        frmBuyHouses.Visible = True
        Call imgLight_Blue_Click
        Call imgProp1_Click
        Call imgProp2_Click
        Call imgProp3_Click
    End If
    
    If Timercount = 20 Then
        If mPlayerBank <= 50000 Then
            Call cmdCancel_Click
            Timercount = 0
            Call Player2Purchase
            tmrLight_Blue.Enabled = False
            Exit Sub
        End If
        If imgPurchaseHouse(0).Tag = 0 Or imgPurchaseHouse(1).Tag = 0 Or imgPurchaseHouse(2).Tag = 0 Then
            If imgPurchaseHouse(0).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(1).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(1)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(6).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(6)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(11).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(11)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 1 Or imgPurchaseHouse(1).Tag = 1 Or imgPurchaseHouse(2).Tag = 1 Then
            If imgPurchaseHouse(0).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(2).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(2)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(7).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(7)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(12).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(12)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 2 Or imgPurchaseHouse(1).Tag = 2 Or imgPurchaseHouse(2).Tag = 2 Then
            If imgPurchaseHouse(0).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(3).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(3)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(8).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(8)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(13).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(13)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 3 Or imgPurchaseHouse(1).Tag = 3 Or imgPurchaseHouse(2).Tag = 3 Then
            If imgPurchaseHouse(0).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(4).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(4)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(9).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(9)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(14).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(14)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 4 Or imgPurchaseHouse(1).Tag = 4 Or imgPurchaseHouse(2).Tag = 4 Then
            If imgPurchaseHouse(0).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(5).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(5)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(10).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(10)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(15).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(15)
                End If
            End If
        End If
    End If
    
    If Timercount = 30 Then
        Call cmdPurchase_Click
        Timercount = 0
        Call Player2Purchase
        tmrLight_Blue.Enabled = False
    End If
    
End Sub

Private Sub tmrLight_Brown_Timer()
    Dim Bank As Long
    
    Timercount = Timercount + 1
    
    If Timercount = 10 Then
        frmBuyHouses.Visible = True
        Call imgLight_Brown_Click
        Call imgProp1_Click
        Call imgProp2_Click
        Call imgProp3_Click
    End If
    
    If Timercount = 20 Then
        If mPlayerBank <= 50000 Then
            Call cmdCancel_Click
            Timercount = 0
            Call Player2Purchase
            tmrLight_Brown.Enabled = False
            Exit Sub
        End If
        If imgPurchaseHouse(0).Tag = 0 Or imgPurchaseHouse(1).Tag = 0 Or imgPurchaseHouse(2).Tag = 0 Then
            If imgPurchaseHouse(0).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(1).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(1)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(6).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(6)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(11).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(11)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 1 Or imgPurchaseHouse(1).Tag = 1 Or imgPurchaseHouse(2).Tag = 1 Then
            If imgPurchaseHouse(0).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(2).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(2)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(7).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(7)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(12).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(12)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 2 Or imgPurchaseHouse(1).Tag = 2 Or imgPurchaseHouse(2).Tag = 2 Then
            If imgPurchaseHouse(0).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(3).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(3)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(8).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(8)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(13).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(13)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 3 Or imgPurchaseHouse(1).Tag = 3 Or imgPurchaseHouse(2).Tag = 3 Then
            If imgPurchaseHouse(0).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(4).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(4)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(9).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(9)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(14).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(14)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 4 Or imgPurchaseHouse(1).Tag = 4 Or imgPurchaseHouse(2).Tag = 4 Then
            If imgPurchaseHouse(0).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(5).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(5)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(10).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(10)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(15).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(15)
                End If
            End If
        End If
    End If
    
    If Timercount = 30 Then
        Call cmdPurchase_Click
        Timercount = 0
        Call Player2Purchase
        tmrLight_Brown.Enabled = False
    End If
    

End Sub

Private Sub tmrPurple_Timer()
    Dim Bank As Long
    
    Timercount = Timercount + 1
    
    If Timercount = 10 Then
        frmBuyHouses.Visible = True
        Call imgPurple_Click
        Call imgProp1_Click
        Call imgProp2_Click
        Call imgProp3_Click
    End If
    
    If Timercount = 20 Then
        If mPlayerBank <= 50000 Then
            Call cmdCancel_Click
            Timercount = 0
            Call Player2Purchase
            tmrPurple.Enabled = False
            Exit Sub
        End If
        If imgPurchaseHouse(0).Tag = 0 Or imgPurchaseHouse(1).Tag = 0 Or imgPurchaseHouse(2).Tag = 0 Then
            If imgPurchaseHouse(0).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(1).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(1)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(6).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(6)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(11).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(11)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 1 Or imgPurchaseHouse(1).Tag = 1 Or imgPurchaseHouse(2).Tag = 1 Then
            If imgPurchaseHouse(0).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(2).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(2)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(7).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(7)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(12).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(12)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 2 Or imgPurchaseHouse(1).Tag = 2 Or imgPurchaseHouse(2).Tag = 2 Then
            If imgPurchaseHouse(0).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(3).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(3)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(8).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(8)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(13).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(13)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 3 Or imgPurchaseHouse(1).Tag = 3 Or imgPurchaseHouse(2).Tag = 3 Then
            If imgPurchaseHouse(0).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(4).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(4)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(9).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(9)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(14).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(14)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 4 Or imgPurchaseHouse(1).Tag = 4 Or imgPurchaseHouse(2).Tag = 4 Then
            If imgPurchaseHouse(0).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(5).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(5)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(10).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(10)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(15).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(15)
                End If
            End If
        End If
    End If
    
    If Timercount = 30 Then
        Call cmdPurchase_Click
        Timercount = 0
        Call Player2Purchase
        tmrPurple.Enabled = False
    End If
    
End Sub

Private Sub tmrRed_Timer()
    Dim Bank As Long
    
    Timercount = Timercount + 1
    
    If Timercount = 10 Then
        frmBuyHouses.Visible = True
        Call imgRed_Click
        Call imgProp1_Click
        Call imgProp2_Click
        Call imgProp3_Click
    End If
    
    If Timercount = 20 Then
        If mPlayerBank <= 50000 Then
            Call cmdCancel_Click
            Timercount = 0
            Call Player2Purchase
            tmrRed.Enabled = False
            Exit Sub
        End If
        If imgPurchaseHouse(0).Tag = 0 Or imgPurchaseHouse(1).Tag = 0 Or imgPurchaseHouse(2).Tag = 0 Then
            If imgPurchaseHouse(0).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(1).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(1)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(6).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(6)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(11).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(11)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 1 Or imgPurchaseHouse(1).Tag = 1 Or imgPurchaseHouse(2).Tag = 1 Then
            If imgPurchaseHouse(0).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(2).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(2)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(7).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(7)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(12).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(12)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 2 Or imgPurchaseHouse(1).Tag = 2 Or imgPurchaseHouse(2).Tag = 2 Then
            If imgPurchaseHouse(0).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(3).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(3)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(8).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(8)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(13).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(13)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 3 Or imgPurchaseHouse(1).Tag = 3 Or imgPurchaseHouse(2).Tag = 3 Then
            If imgPurchaseHouse(0).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(4).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(4)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(9).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(9)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(14).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(14)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 4 Or imgPurchaseHouse(1).Tag = 4 Or imgPurchaseHouse(2).Tag = 4 Then
            If imgPurchaseHouse(0).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(5).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(5)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(10).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(10)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(15).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(15)
                End If
            End If
        End If
    End If
    
    If Timercount = 30 Then
        Call cmdPurchase_Click
        Timercount = 0
        Call Player2Purchase
        tmrRed.Enabled = False
    End If
    

End Sub

Private Sub tmrYellow_Timer()
    Dim Bank As Long
    
    Timercount = Timercount + 1
    
    If Timercount = 10 Then
        frmBuyHouses.Visible = True
        Call imgYellow_Click
        Call imgProp1_Click
        Call imgProp2_Click
        Call imgProp3_Click
    End If
    
    If Timercount = 20 Then
        If mPlayerBank <= 50000 Then
            Call cmdCancel_Click
            Timercount = 0
            Call Player2Purchase
            tmrYellow.Enabled = False
            Exit Sub
        End If
        If imgPurchaseHouse(0).Tag = 0 Or imgPurchaseHouse(1).Tag = 0 Or imgPurchaseHouse(2).Tag = 0 Then
            If imgPurchaseHouse(0).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(1).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(1)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(6).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(6)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 0 Then
                Bank = mPlayerBank - imgHouse(11).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(11)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 1 Or imgPurchaseHouse(1).Tag = 1 Or imgPurchaseHouse(2).Tag = 1 Then
            If imgPurchaseHouse(0).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(2).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(2)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(7).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(7)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 1 Then
                Bank = mPlayerBank - imgHouse(12).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(12)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 2 Or imgPurchaseHouse(1).Tag = 2 Or imgPurchaseHouse(2).Tag = 2 Then
            If imgPurchaseHouse(0).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(3).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(3)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(8).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(8)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 2 Then
                Bank = mPlayerBank - imgHouse(13).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(13)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 3 Or imgPurchaseHouse(1).Tag = 3 Or imgPurchaseHouse(2).Tag = 3 Then
            If imgPurchaseHouse(0).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(4).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(4)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(9).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(9)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 3 Then
                Bank = mPlayerBank - imgHouse(14).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(14)
                End If
            End If
        End If
        
        If imgPurchaseHouse(0).Tag = 4 Or imgPurchaseHouse(1).Tag = 4 Or imgPurchaseHouse(2).Tag = 4 Then
            If imgPurchaseHouse(0).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(5).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(5)
                End If
            End If
            
            If imgPurchaseHouse(1).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(10).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(10)
                End If
            End If
            
            If imgPurchaseHouse(2).Tag = 4 Then
                Bank = mPlayerBank - imgHouse(15).Tag
                If Bank > 20000 Then
                    Call imgHouse_Click(15)
                End If
            End If
        End If
    End If
    
    If Timercount = 30 Then
        Call cmdPurchase_Click
        Timercount = 0
        Call Player2Purchase
        tmrYellow.Enabled = False
    End If
    

End Sub
