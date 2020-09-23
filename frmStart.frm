VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmStart 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Selection"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   ControlBox      =   0   'False
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   2370
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5625
      Top             =   3525
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Players"
      Height          =   1890
      Left            =   1665
      TabIndex        =   1
      Top             =   135
      Width           =   3030
      Begin VB.TextBox txtPlayer2Name 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   4
         Top             =   1185
         Width           =   1365
      End
      Begin VB.TextBox txtPlayer1Name 
         Height          =   285
         Left            =   1140
         TabIndex        =   3
         Top             =   465
         Width           =   1365
      End
      Begin VB.Image imgPlayer2 
         Height          =   480
         Left            =   390
         ToolTipText     =   "None2"
         Top             =   1035
         Width           =   480
      End
      Begin VB.Image imgPlayer1 
         Height          =   480
         Left            =   405
         ToolTipText     =   "None1"
         Top             =   315
         Width           =   480
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   525
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   375
      Visible         =   0   'False
      Width           =   510
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   -1  'True
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   900
      _cy             =   873
   End
   Begin VB.Label txtHelp 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   30
      TabIndex        =   5
      Top             =   4305
      Width           =   6315
   End
   Begin VB.Image imgPreview 
      Height          =   480
      Index           =   3
      Left            =   4800
      Top             =   3525
      Width           =   480
   End
   Begin VB.Image imgPreview 
      Height          =   480
      Index           =   2
      Left            =   3645
      Top             =   3525
      Width           =   480
   End
   Begin VB.Image imgPreview 
      Height          =   480
      Index           =   1
      Left            =   2415
      Top             =   3525
      Width           =   480
   End
   Begin VB.Image imgPreview 
      Height          =   480
      Index           =   0
      Left            =   1215
      Top             =   3525
      Width           =   480
   End
   Begin VB.Label lblCharSel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player 1 - Please choose your character"
      Height          =   210
      Left            =   360
      TabIndex        =   2
      Top             =   2595
      Width           =   2895
   End
   Begin VB.Image imgDog 
      Height          =   480
      Left            =   2115
      Picture         =   "frmStart.frx":030A
      Tag             =   "Dog"
      Top             =   2948
      Width           =   480
   End
   Begin VB.Image imgWheelbarrow 
      Height          =   480
      Left            =   5640
      Picture         =   "frmStart.frx":0614
      Tag             =   "Wheelbarrow"
      Top             =   2948
      Width           =   480
   End
   Begin VB.Image imgThumble 
      Height          =   480
      Left            =   5055
      Picture         =   "frmStart.frx":091E
      Tag             =   "Thumble"
      Top             =   2948
      Width           =   480
   End
   Begin VB.Image imgShoe 
      Height          =   480
      Left            =   4455
      Picture         =   "frmStart.frx":0C28
      Tag             =   "Shoe"
      Top             =   2948
      Width           =   480
   End
   Begin VB.Image imgIron 
      Height          =   480
      Left            =   3855
      Picture         =   "frmStart.frx":0F32
      Tag             =   "Iron"
      Top             =   2948
      Width           =   480
   End
   Begin VB.Image imgHorse 
      Height          =   480
      Left            =   3270
      Picture         =   "frmStart.frx":123C
      Tag             =   "Horse"
      Top             =   2948
      Width           =   480
   End
   Begin VB.Image imgHat 
      Height          =   480
      Left            =   2670
      Picture         =   "frmStart.frx":1546
      Tag             =   "Hat"
      Top             =   2948
      Width           =   480
   End
   Begin VB.Image imgCar 
      Height          =   480
      Left            =   1485
      Picture         =   "frmStart.frx":1850
      Tag             =   "Car"
      Top             =   2948
      Width           =   480
   End
   Begin VB.Image imgCannon 
      Height          =   480
      Left            =   900
      Picture         =   "frmStart.frx":1B5A
      Tag             =   "Cannon"
      Top             =   2948
      Width           =   480
   End
   Begin VB.Image imgShip 
      Height          =   480
      Left            =   300
      Picture         =   "frmStart.frx":1E64
      Tag             =   "Ship"
      Top             =   2948
      Width           =   480
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim Ap As String
    Dim mChar As String
    Dim mIconCount As Integer
    Dim Start As String

Private Sub cmdSelect_Click()
    Dim R As Integer

        If imgPlayer1.Tag = "" Then
            imgPlayer1.Picture = LoadPicture(Ap & mChar & "\Left1.ico")
            lblCharSel.Caption = "Please choose the Computer character."
            imgPlayer1.ToolTipText = mChar
            frmBoard.imgPlayer1Start.ToolTipText = mChar
            imgPlayer1.Tag = 1
            txtHelp.Caption = "Select a Character for CPU and the click on Select."
        ElseIf imgPlayer1.Tag = 1 Then
            If mChar = imgPlayer1.ToolTipText Then
                R = MsgBox("Plaese select a diffirent Character. You cannot use the same Character for both Players.", vbOKOnly, "Character Selection")
            Else
                imgPlayer2.Visible = True
                imgPlayer2.Picture = LoadPicture(Ap & mChar & "\Left1.ico")
                imgPlayer2.ToolTipText = mChar
                frmBoard.imgPlayer2Start.ToolTipText = mChar
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''''' LOAD THE FROM AND THE CHARACTERS ''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                frmStart.Hide
                frmBoard.Show
                frmBoard.imgPlayer1Start.Picture = LoadPicture(Ap & imgPlayer1.ToolTipText & "\Left1.ico")
                frmMessageBuy.imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.ToolTipText & "\Left1.ico")
                frmBoard.imgPlayer2Start.Picture = LoadPicture(Ap & imgPlayer2.ToolTipText & "\Left1.ico")
                frmMessageBuy.imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.ToolTipText & "\Left1.ico")
                frmBoard.imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.ToolTipText & "\Up1.ico")
                frmBoard.imgPlayer1.Tag = imgPlayer1.ToolTipText
                frmBoard.imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.ToolTipText & "\Up1.ico")
                frmBoard.imgPlayer2.Tag = imgPlayer2.ToolTipText
                frmBoard.lblPlayer1Name = txtPlayer1Name.Text
                frmBoard.lblPlayer2Name = txtPlayer2Name.Text
            End If
        End If
            
    If txtPlayer1Name.Text = "" Then
        txtPlayer1Name.Text = "Human"
    End If
    
    If txtPlayer2Name.Text = "" Then
        txtPlayer2Name.Text = "CPU"
    End If
    
            
End Sub

Private Sub Form_Load()

    If Right(App.Path, 1) = "\" Then
        Ap = App.Path
    Else
        Ap = App.Path & "\"
    End If

    txtPlayer1Name.Text = InputBox("Please enter the name for Player 1", "Player Name")
    
    txtHelp.Caption = "Click on an Icon to select your Character and then on Select."
    
    txtPlayer2Name.Text = "CPU"
End Sub

Private Sub imgCannon_Click()

    mChar = imgCannon.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\CANNON.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 1
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 1
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 1
    End If
    tmrAnimate.Enabled = True

End Sub

Private Sub imgCar_Click()

    mChar = imgCar.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\CAR.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 1
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 1
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 1
    End If
    tmrAnimate.Enabled = True

End Sub

Private Sub imgDog_Click()

    mChar = imgDog.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\DOG.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 2
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 2
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 2
    End If
    tmrAnimate.Enabled = True

End Sub

Private Sub imgHat_Click()

    mChar = imgHat.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\HATLP.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 1
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 1
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 1
    End If
    tmrAnimate.Enabled = True
    
End Sub

Private Sub imgHorse_Click()

    mChar = imgHorse.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\HORSE.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 1
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 1
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 1
    End If
    tmrAnimate.Enabled = True
    
End Sub

Private Sub imgIron_Click()

    mChar = imgIron.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\IRON.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 1
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 1
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 1
    End If
    tmrAnimate.Enabled = True
    
End Sub

Private Sub imgShip_Click()

    mChar = imgShip.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\SHIP.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 1
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 1
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 1
    End If
    tmrAnimate.Enabled = True
    

End Sub

Private Sub imgShoe_Click()

    mChar = imgShoe.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\SHOE.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 1
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 1
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 1
    End If
    tmrAnimate.Enabled = True
    
End Sub

Private Sub imgThumble_Click()

    mChar = imgThumble.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\THUMBLE.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 1
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 1
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 1
    End If
    tmrAnimate.Enabled = True
    
End Sub

Private Sub imgWheelbarrow_Click()

    mChar = imgWheelbarrow.Tag
    cmdSelect.Visible = True
    Start = Ap & "Sound\WHEELBARROW.WAV"
    WindowsMediaPlayer1.URL = Start
    WindowsMediaPlayer1.settings.playCount = 2
    WindowsMediaPlayer1.Controls.Play
    If frmBoard.txtP1Start.Text > "" Then
        frmBoard.txtP2Start.Text = Start
        frmBoard.txtP2Loop.Text = 2
    Else
        frmBoard.txtP1Start.Text = Start
        frmBoard.txtP1Loop.Text = 2
    End If
    tmrAnimate.Enabled = True
    
End Sub

Private Sub tmrAnimate_Timer()
    
    mIconCount = mIconCount + 1
    
    With imgPreview
        .Item(0).Picture = LoadPicture(Ap & mChar & "\Left" & mIconCount & ".ico")
        .Item(1).Picture = LoadPicture(Ap & mChar & "\Up" & mIconCount & ".ico")
        .Item(2).Picture = LoadPicture(Ap & mChar & "\Right" & mIconCount & ".ico")
        .Item(3).Picture = LoadPicture(Ap & mChar & "\Down" & mIconCount & ".ico")
    End With
    
    If mIconCount = 7 Then
        mIconCount = 1
    End If
    
    
End Sub


Private Sub txtPlayer2Name_Change()
    Dim R As Integer
    If txtPlayer2Name.Text <> "CPU" Then
        R = MsgBox("You cannot change the Player 2 Name. Thank You.", vbOKOnly, "Sorry.")
    End If
    
    txtPlayer2Name.Text = "CPU"
    
End Sub
