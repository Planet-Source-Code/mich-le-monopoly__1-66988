VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBoard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBoard.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   789
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   630
      Left            =   8415
      TabIndex        =   18
      Top             =   7740
      Width           =   990
   End
   Begin VB.Timer tmrViewOwner 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1950
      Top             =   8310
   End
   Begin VB.Timer tmrPlayer2ToJail 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   11055
      Top             =   7500
   End
   Begin VB.Timer tmrPlayer1ToJail 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10590
      Top             =   7500
   End
   Begin Project1.MorphListBox lsbOne 
      Height          =   1470
      Left            =   2430
      TabIndex        =   16
      Top             =   7485
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   2593
      BackAngle       =   225
      BackColor2      =   16777215
      BackColor1      =   16744576
      BorderColor     =   4194304
      BackMiddleOut   =   0   'False
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   0   'False
      SelColor1       =   8388608
      SelColor2       =   8388608
      SelTextColor    =   16777152
      TrackBarColor1  =   8388608
      TrackBarColor2  =   16761024
      ButtonColor1    =   4194304
      ButtonColor2    =   16744576
      ThumbColor1     =   4194304
      ThumbColor2     =   16744576
      ThumbBorderColor=   16761024
      ArrowUpColor    =   16761024
      ArrowDownColor  =   4194304
      Theme           =   3
      CheckboxArrowColor=   4194304
      CheckBoxColor   =   4194304
      FocusRectColor  =   16761024
      TrackClickColor1=   4194304
      TrackClickColor2=   16744576
   End
   Begin VB.CommandButton cmdMsgNo 
      Caption         =   "NO"
      Height          =   300
      Left            =   1245
      TabIndex        =   14
      Top             =   6855
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton cmdMsgYes 
      Caption         =   "YES"
      Height          =   300
      Left            =   450
      TabIndex        =   13
      Top             =   6855
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Timer Player2m0to1 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11310
      Top             =   8460
   End
   Begin VB.Timer Player2m1to2 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11340
      Top             =   8460
   End
   Begin VB.Timer Player2m2to3 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11340
      Top             =   8445
   End
   Begin VB.Timer Player2m3to4 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11325
      Top             =   8460
   End
   Begin VB.Timer Player2m4to5 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11340
      Top             =   8460
   End
   Begin VB.Timer Player2m5to6 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11340
      Top             =   8445
   End
   Begin VB.Timer Player2m6to7 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8445
   End
   Begin VB.Timer Player2m7to8 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8430
   End
   Begin VB.Timer Player2m8to9 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8445
   End
   Begin VB.Timer Player2m9to10 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8445
   End
   Begin VB.Timer Player2m10to11 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11340
      Top             =   8535
   End
   Begin VB.Timer Player2m11to12 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8535
   End
   Begin VB.Timer Player2m12to13 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8520
   End
   Begin VB.Timer Player2m13to14 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8535
   End
   Begin VB.Timer Player2m14to15 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8535
   End
   Begin VB.Timer Player2m15to16 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11340
      Top             =   8475
   End
   Begin VB.Timer Player2m16to17 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8475
   End
   Begin VB.Timer Player2m17to18 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8460
   End
   Begin VB.Timer Player2m18to19 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8475
   End
   Begin VB.Timer Player2m19to20 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8475
   End
   Begin VB.Timer Player2m20to21 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11340
      Top             =   8490
   End
   Begin VB.Timer Player2m21to22 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8490
   End
   Begin VB.Timer Player2m22to23 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8475
   End
   Begin VB.Timer Player2m23to24 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8490
   End
   Begin VB.Timer Player2m24to25 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8490
   End
   Begin VB.Timer Player2m25to26 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11340
      Top             =   8475
   End
   Begin VB.Timer Player2m26to27 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8475
   End
   Begin VB.Timer Player2m27to28 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8460
   End
   Begin VB.Timer Player2m28to29 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8475
   End
   Begin VB.Timer Player2m29to30 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8475
   End
   Begin VB.Timer Player2m30to31 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11355
      Top             =   8490
   End
   Begin VB.Timer Player2m31to32 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11385
      Top             =   8490
   End
   Begin VB.Timer Player2m32to33 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11385
      Top             =   8475
   End
   Begin VB.Timer Player2m33to34 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11385
      Top             =   8490
   End
   Begin VB.Timer Player2m34to35 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11385
      Top             =   8490
   End
   Begin VB.Timer Player2m35to36 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11340
      Top             =   8490
   End
   Begin VB.Timer Player2m36to37 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8490
   End
   Begin VB.Timer Player2m37to38 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8475
   End
   Begin VB.Timer Player2m38to39 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8490
   End
   Begin VB.Timer Player2m39to0 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   11370
      Top             =   8490
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player 1"
      Height          =   825
      Left            =   8010
      TabIndex        =   4
      Top             =   1110
      Width           =   2385
      Begin VB.Image imgPlayer1Turn 
         Height          =   480
         Left            =   90
         Picture         =   "frmBoard.frx":ED1B
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblPlayer1Name 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   810
         TabIndex        =   6
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label lblPlayerBank 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   1
         Left            =   810
         TabIndex        =   5
         Top             =   510
         Width           =   1380
      End
      Begin VB.Image imgPlayer1Start 
         Height          =   480
         Left            =   90
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame frmPlayer1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player 2"
      Height          =   825
      Left            =   4140
      TabIndex        =   2
      Top             =   5310
      Width           =   2385
      Begin VB.Label lblPlayerBank 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   810
         TabIndex        =   7
         Top             =   510
         Width           =   1380
      End
      Begin VB.Image imgPlayer2Turn 
         Height          =   480
         Left            =   90
         Picture         =   "frmBoard.frx":F025
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblPlayer2Name 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   810
         TabIndex        =   3
         Top             =   225
         Width           =   1380
      End
      Begin VB.Image imgPlayer2Start 
         Height          =   480
         Left            =   90
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Timer tmrDiceRoll 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1710
      Top             =   7545
   End
   Begin VB.Timer Player1m39to0 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10755
      Top             =   8475
   End
   Begin VB.Timer Player1m38to39 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8475
   End
   Begin VB.Timer Player1m37to38 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8475
   End
   Begin VB.Timer Player1m36to37 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10755
      Top             =   8475
   End
   Begin VB.Timer Player1m35to36 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10740
      Top             =   8475
   End
   Begin VB.Timer Player1m34to35 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8475
   End
   Begin VB.Timer Player1m33to34 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8475
   End
   Begin VB.Timer Player1m32to33 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8475
   End
   Begin VB.Timer Player1m31to32 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8475
   End
   Begin VB.Timer Player1m30to31 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10755
      Top             =   8475
   End
   Begin VB.Timer Player1m29to30 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8475
   End
   Begin VB.Timer Player1m28to29 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8475
   End
   Begin VB.Timer Player1m27to28 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8475
   End
   Begin VB.Timer Player1m26to27 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8475
   End
   Begin VB.Timer Player1m25to26 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10755
      Top             =   8475
   End
   Begin VB.Timer Player1m24to25 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8520
   End
   Begin VB.Timer Player1m23to24 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10800
      Top             =   8520
   End
   Begin VB.Timer Player1m22to23 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10800
      Top             =   8520
   End
   Begin VB.Timer Player1m21to22 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8520
   End
   Begin VB.Timer Player1m20to21 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8520
   End
   Begin VB.Timer Player1m19to20 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10815
      Top             =   8475
   End
   Begin VB.Timer Player1m18to19 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10830
      Top             =   8475
   End
   Begin VB.Timer Player1m17to18 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10830
      Top             =   8475
   End
   Begin VB.Timer Player1m16to17 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10815
      Top             =   8475
   End
   Begin VB.Timer Player1m15to16 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10800
      Top             =   8475
   End
   Begin VB.Timer Player1m14to15 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8490
   End
   Begin VB.Timer Player1m13to14 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10800
      Top             =   8490
   End
   Begin VB.Timer Player1m12to13 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10800
      Top             =   8490
   End
   Begin VB.Timer Player1m11to12 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8490
   End
   Begin VB.Timer Player1m10to11 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8490
   End
   Begin VB.Timer Player1m9to10 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8460
   End
   Begin VB.Timer Player1m8to9 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8460
   End
   Begin VB.Timer Player1m7to8 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10785
      Top             =   8460
   End
   Begin VB.Timer Player1m6to7 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10770
      Top             =   8460
   End
   Begin VB.Timer Player1m5to6 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10755
      Top             =   8460
   End
   Begin VB.Timer Player1m4to5 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10725
      Top             =   8475
   End
   Begin VB.Timer Player1m3to4 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10725
      Top             =   8475
   End
   Begin VB.Timer Player1m2to3 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10740
      Top             =   8475
   End
   Begin VB.Timer Player1m1to2 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10725
      Top             =   8475
   End
   Begin VB.Timer Player1m0to1 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   10710
      Top             =   8475
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   465
      Left            =   9900
      TabIndex        =   17
      Top             =   5850
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   820
      _Version        =   393216
      BackColor       =   16777215
      FullWidth       =   41
      FullHeight      =   31
   End
   Begin VB.Image imgHideOwner 
      Height          =   450
      Left            =   450
      Picture         =   "frmBoard.frx":F32F
      ToolTipText     =   "Hide Owners view"
      Top             =   8535
      Width           =   1500
   End
   Begin VB.Image imgViewOwned 
      Height          =   450
      Left            =   450
      Picture         =   "frmBoard.frx":F83B
      ToolTipText     =   "View all owners"
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   39
      Left            =   3285
      Top             =   6315
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   38
      Left            =   3900
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   37
      Left            =   4620
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   36
      Left            =   5385
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   35
      Left            =   6150
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   34
      Left            =   6900
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   33
      Left            =   7650
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   32
      Left            =   8400
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   31
      Left            =   9165
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   30
      Left            =   9900
      Top             =   6540
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   29
      Left            =   10530
      Top             =   6330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   28
      Left            =   10800
      Top             =   5790
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   27
      Left            =   10800
      Top             =   5190
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   26
      Left            =   10800
      Top             =   4605
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   25
      Left            =   10800
      Top             =   4020
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   24
      Left            =   10800
      Top             =   3435
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   23
      Left            =   10800
      Top             =   2850
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   22
      Left            =   10800
      Top             =   2265
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   21
      Left            =   10800
      Top             =   1695
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   20
      Left            =   10800
      Top             =   1095
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   19
      Left            =   10545
      Top             =   525
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   18
      Left            =   9915
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   17
      Left            =   9150
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   16
      Left            =   8430
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   15
      Left            =   7665
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   14
      Left            =   6915
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   13
      Left            =   6165
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   12
      Left            =   5415
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   11
      Left            =   4665
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   10
      Left            =   3915
      Top             =   330
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   9
      Left            =   2760
      Top             =   225
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   8
      Left            =   3030
      Top             =   1095
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   7
      Left            =   3030
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   6
      Left            =   3030
      Top             =   2265
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   5
      Left            =   3030
      Top             =   2850
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   4
      Left            =   3030
      Top             =   3435
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   3
      Left            =   3030
      Top             =   4005
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   2
      Left            =   3030
      Top             =   4605
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   1
      Left            =   3030
      Top             =   5190
      Width           =   480
   End
   Begin VB.Image imgOwner 
      Height          =   480
      Index           =   0
      Left            =   3030
      Top             =   5775
      Width           =   480
   End
   Begin VB.Label lblMsgPlayer 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   375
      TabIndex        =   15
      Top             =   4695
      Visible         =   0   'False
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
      Left            =   375
      TabIndex        =   12
      Top             =   6075
      Visible         =   0   'False
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
      Left            =   375
      TabIndex        =   11
      Top             =   5775
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblMsgMain 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Would you like to purchase this property?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   375
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblMsgNote 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "*Click on the property to view the Title Dead."
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   330
      TabIndex        =   9
      Top             =   6405
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label lblMsgPurchase 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purchase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      TabIndex        =   8
      Top             =   4245
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image imgMessage 
      Height          =   3465
      Left            =   15
      Picture         =   "frmBoard.frx":FD4F
      Stretch         =   -1  'True
      Top             =   3990
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label lblDeadText 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   3225
      Left            =   -15
      TabIndex        =   1
      Top             =   720
      Width           =   2430
   End
   Begin VB.Image imgPlayer2 
      Height          =   480
      Left            =   2460
      Top             =   6975
      Width           =   480
   End
   Begin VB.Image imgPlayer1 
      Height          =   480
      Left            =   2460
      Top             =   6975
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmBoard.frx":11073
      Top             =   7515
      Width           =   480
   End
   Begin VB.Image imgDice 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "frmBoard.frx":1137D
      Top             =   7515
      Width           =   480
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   4
      Left            =   2460
      Tag             =   "0"
      Top             =   3375
      Width           =   1335
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   39
      Left            =   2460
      Tag             =   "0"
      Top             =   6300
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   38
      Left            =   3765
      Tag             =   "0"
      Top             =   6300
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   37
      Left            =   4515
      Tag             =   "0"
      Top             =   6300
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   36
      Left            =   5265
      Tag             =   "0"
      Top             =   6300
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   35
      Left            =   6015
      Tag             =   "0"
      Top             =   6300
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   34
      Left            =   6765
      Tag             =   "0"
      Top             =   6300
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   33
      Left            =   7515
      Tag             =   "0"
      Top             =   6300
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   32
      Left            =   8265
      Tag             =   "0"
      Top             =   6300
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   31
      Left            =   9015
      Tag             =   "0"
      Top             =   6300
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1155
      Index           =   30
      Left            =   9765
      Tag             =   "0"
      Top             =   6300
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1170
      Index           =   29
      Left            =   10515
      Tag             =   "0"
      Top             =   6300
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   28
      Left            =   10515
      Tag             =   "0"
      Top             =   5715
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   27
      Left            =   10515
      Tag             =   "0"
      Top             =   5130
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   26
      Left            =   10515
      Tag             =   "0"
      Top             =   4545
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   25
      Left            =   10515
      Tag             =   "0"
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   24
      Left            =   10515
      Tag             =   "0"
      Top             =   3375
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   23
      Left            =   10515
      Tag             =   "0"
      Top             =   2790
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   22
      Left            =   10515
      Tag             =   "0"
      Top             =   2205
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   21
      Left            =   10515
      Tag             =   "0"
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   20
      Left            =   10515
      Tag             =   "0"
      Top             =   1035
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   1020
      Index           =   19
      Left            =   10515
      Tag             =   "0"
      Top             =   0
      Width           =   1335
   End
   Begin VB.Image imgDead 
      Height          =   1020
      Index           =   18
      Left            =   9765
      Tag             =   "0"
      Top             =   0
      Width           =   750
   End
   Begin VB.Image imgDead 
      Height          =   1020
      Index           =   17
      Left            =   9015
      Tag             =   "0"
      Top             =   0
      Width           =   750
   End
   Begin VB.Image imgDead 
      Height          =   1020
      Index           =   16
      Left            =   8265
      Tag             =   "0"
      Top             =   0
      Width           =   750
   End
   Begin VB.Image imgDead 
      Height          =   1020
      Index           =   15
      Left            =   7515
      Tag             =   "0"
      Top             =   0
      Width           =   750
   End
   Begin VB.Image imgDead 
      Height          =   1020
      Index           =   14
      Left            =   6765
      Tag             =   "0"
      Top             =   0
      Width           =   750
   End
   Begin VB.Image imgDead 
      Height          =   1020
      Index           =   13
      Left            =   6015
      Tag             =   "0"
      Top             =   0
      Width           =   750
   End
   Begin VB.Image imgDead 
      Height          =   1020
      Index           =   12
      Left            =   5265
      Tag             =   "0"
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgDead 
      Height          =   1020
      Index           =   11
      Left            =   4515
      Tag             =   "0"
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgDead 
      Height          =   1005
      Index           =   10
      Left            =   3765
      Tag             =   "0"
      Top             =   0
      Width           =   765
   End
   Begin VB.Image imgDead 
      Height          =   1005
      Index           =   9
      Left            =   2460
      Tag             =   "0"
      Top             =   0
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   8
      Left            =   2460
      Tag             =   "0"
      Top             =   1035
      Width           =   1335
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   7
      Left            =   2460
      Tag             =   "0"
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   6
      Left            =   2460
      Tag             =   "0"
      Top             =   2205
      Width           =   1335
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   5
      Left            =   2460
      Tag             =   "0"
      Top             =   2790
      Width           =   1335
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   3
      Left            =   2460
      Tag             =   "0"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   2
      Left            =   2460
      Tag             =   "0"
      Top             =   4545
      Width           =   1335
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   1
      Left            =   2460
      Tag             =   "0"
      Top             =   5130
      Width           =   1320
   End
   Begin VB.Image imgDead 
      Height          =   585
      Index           =   0
      Left            =   2460
      Tag             =   "0"
      Top             =   5715
      Width           =   1320
   End
   Begin VB.Label lblDeadTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   2400
   End
End
Attribute VB_Name = "frmBoard"
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
    
    '''''''''''''''''''''''''''''''''''''''''''
    ''''''' VARIABLES FOR JAIL STATUS '''''''''
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mPLayer1InJail As Integer
    Dim mPLayer2InJail As Integer
    
    '''''''''''''''''''''''''''''''''''''''''''
    ''''''' VARIABLES FOR PLAYER BANK '''''''''
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mPLayerBank(1 To 2) As Double
    '''''''''''''''''''''''''''''''''''''''''''
    '''''''   VARIABLES FOR DICE ROLL '''''''''
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mDiceRoll As Integer
    Dim mDice1Total As Integer
    Dim mDice2Total As Integer
    Dim mDiceTotal As Integer
    '''''''''''''''''''''''''''''''''''''''''''
    ''''''' VARIABLES FOR BOARD COORDINATES '''
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mMoveCount As Integer
    Dim mJailCount As Integer
    '''''''''''''''''''''''''''''''''''''''''''
    ''''''' VARIABLES FOR PROPERTY LISTING ''''
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mPropertyListing(1 To 28) As Integer
    Dim mPropertyValue As Long
    '''''''''''''''''''''''''''''''''''''''''''
    ''''''' VARIABLES FOR PROPERTY LANDEDON '''
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mLandedProperty As Integer
    '''''''''''''''''''''''''''''''''''''''''''
    ''''''' VARIABLES FOR POLICE LIGHT TURN '''
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mPoliceLightCount As Integer
    
    '''''''''''''''''''''''''''''''''''''''''''
    ''''''' VARIABLES FOR STATIONS OWNED ''''''
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mStation(1 To 4) As Integer

    '''''''''''''''''''''''''''''''''''''''''''
    ''''''' VARIABLES FOR PLAYER JAIL COUNT '''
    '''''''''''''''''''''''''''''''''''''''''''
    Dim mPlayer1JailStay As Integer
    Dim mPlayer2JailStay As Integer
    
Private Sub Command1_Click()

    frmCommunityChest.Visible = True

End Sub

Private Sub Form_Load()
    Dim x As Integer
    
    For x = 0 To 39
        imgOwner(x).Visible = False
    Next x
    
    lsbOne.BackColor1 = RGB(169, 218, 250)
    lsbOne.BackColor2 = vbWhite
    mPoliceLightCount = 1
    frmStart.Show
    frmBoard.Hide
        
    If Right(App.Path, 1) = "\" Then
        Ap = App.Path
    Else
        Ap = App.Path & "\"
    End If

    With imgPlayer1
        .Left = 164
        .Top = 465
    End With
    
    With imgPlayer2
        .Left = 164
        .Top = 465
    End With
    
    lblDeadText.Visible = False
    
    mPLayerBank(1) = 150000
    mPLayerBank(2) = 150000
    
    lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(1), 2)
    lblPlayerBank(2).Caption = FormatCurrency(mPLayerBank(2), 2)

    mPropertyListing(1) = 0
    mPropertyListing(2) = 2
    mPropertyListing(3) = 4
    mPropertyListing(4) = 5
    mPropertyListing(5) = 7
    mPropertyListing(6) = 8
    mPropertyListing(7) = 10
    mPropertyListing(8) = 11
    mPropertyListing(9) = 12
    mPropertyListing(10) = 13
    mPropertyListing(11) = 14
    mPropertyListing(12) = 15
    mPropertyListing(13) = 17
    mPropertyListing(14) = 18
    mPropertyListing(15) = 20
    mPropertyListing(16) = 22
    mPropertyListing(17) = 23
    mPropertyListing(18) = 24
    mPropertyListing(19) = 25
    mPropertyListing(20) = 26
    mPropertyListing(21) = 27
    mPropertyListing(22) = 28
    mPropertyListing(23) = 30
    mPropertyListing(24) = 31
    mPropertyListing(25) = 33
    mPropertyListing(26) = 34
    mPropertyListing(27) = 36
    mPropertyListing(28) = 38

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    imgViewOwned.Picture = LoadPicture(Ap & "button1.gif")
    
End Sub

Private Sub imgDead_Click(Index As Integer)

    Select Case Index
        Case Is = 0
            With lblDeadTitle
                .Caption = vbNewLine & "WESTVILLE"
                .BackColor = RGB(210, 150, 130)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 1
            With lblDeadTitle
                .Caption = "COMMUNITY" & vbNewLine & "CHEST"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call CommunityChest
        Case Is = 2
            With lblDeadTitle
                .Caption = vbNewLine & "AMANZIMTOTI"
                .BackColor = RGB(210, 150, 130)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 3
            With lblDeadTitle
                .Caption = "INCOME" & vbNewLine & "TAX"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call IncomeTax
        Case Is = 4
            With lblDeadTitle
                .Caption = "DURBAN" & vbNewLine & "INTERNATIONAL"
                .BackColor = RGB(210, 210, 210)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 5
            With lblDeadTitle
                .Caption = "UMHLANGA" & vbNewLine & "ROCKS"
                .BackColor = RGB(0, 255, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 6
            With lblDeadTitle
                .Caption = vbNewLine & "CHANCE"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call Chance
        Case Is = 7
            With lblDeadTitle
                .Caption = "BALLITO" & vbNewLine & "BAY"
                .BackColor = RGB(0, 255, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 8
            With lblDeadTitle
                .Caption = "LA" & vbNewLine & "LUCIA"
                .BackColor = RGB(0, 255, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 9
            With lblDeadTitle
                .Caption = "VISITING or" & vbNewLine & "IN JAIL"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call VisitingJail
        Case Is = 10
            With lblDeadTitle
                .Caption = "MENLYN" & vbNewLine & "PARK"
                .BackColor = RGB(255, 0, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 11
            With lblDeadTitle
                .Caption = "ELECTRICAL" & vbNewLine & "BOARD"
                .BackColor = RGB(210, 210, 210)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 12
            With lblDeadTitle
                .Caption = "PORT" & vbNewLine & "ELIZABETH"
                .BackColor = RGB(255, 0, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 13
            With lblDeadTitle
                .Caption = vbNewLine & "WATERKLOOF"
                .BackColor = RGB(255, 0, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 14
            With lblDeadTitle
                .Caption = "BLOEMFONTEIN" & vbNewLine & "INTERNATIONAL"
                .BackColor = RGB(210, 210, 210)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 15
            With lblDeadTitle
                .Caption = vbNewLine & "WILDERNESS"
                .BackColor = RGB(255, 200, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 16
            With lblDeadTitle
                .Caption = "COMMUNITY" & vbNewLine & "CHEST"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call CommunityChest
        Case Is = 17
            With lblDeadTitle
                .Caption = vbNewLine & "KNYSNA"
                .BackColor = RGB(255, 200, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 18
            With lblDeadTitle
                .Caption = "PLETTENBERG" & vbNewLine & "BAY"
                .BackColor = RGB(255, 200, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 19
            With lblDeadTitle
                .Caption = "FREE" & vbNewLine & "PARKING"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call FreeParking
        Case Is = 20
            With lblDeadTitle
                .Caption = vbNewLine & "SOWETO"
                .BackColor = RGB(255, 0, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 21
            With lblDeadTitle
                .Caption = vbNewLine & "CHANCE"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call Chance
        Case Is = 22
            With lblDeadTitle
                .Caption = vbNewLine & "HILLBROW"
                .BackColor = RGB(255, 0, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 23
            With lblDeadTitle
                .Caption = vbNewLine & "BOKSBURG"
                .BackColor = RGB(255, 0, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 24
            With lblDeadTitle
                .Caption = "JOHANESBURG" & vbNewLine & "INTERNATIONAL"
                .BackColor = RGB(210, 210, 210)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 25
            With lblDeadTitle
                .Caption = vbNewLine & "RANDBURG"
                .BackColor = RGB(255, 255, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 26
            With lblDeadTitle
                .Caption = vbNewLine & "SANDTON"
                .BackColor = RGB(255, 255, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 27
            With lblDeadTitle
                .Caption = "WATER" & vbNewLine & "BOARD"
                .BackColor = RGB(210, 210, 210)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 28
            With lblDeadTitle
                .Caption = vbNewLine & "HYDEPARK"
                .BackColor = RGB(255, 255, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 29
            With lblDeadTitle
                .Caption = "GO TO" & vbNewLine & "JAIL!"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call GoToJail
        Case Is = 30
            With lblDeadTitle
                .Caption = vbNewLine & "TYGERVALLEY"
                .BackColor = RGB(0, 255, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 31
            With lblDeadTitle
                .Caption = "MITCHELLS" & vbNewLine & "PLAIN"
                .BackColor = RGB(0, 255, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 32
            With lblDeadTitle
                .Caption = "COMMUNITY" & vbNewLine & "CHEST"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call CommunityChest
        Case Is = 33
            With lblDeadTitle
                .Caption = "BLOUBERG" & vbNewLine & "STRAND"
                .BackColor = RGB(0, 255, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 34
            With lblDeadTitle
                .Caption = "CAPE TOWN" & vbNewLine & "INTERNATIONAL"
                .BackColor = RGB(210, 210, 210)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 35
            With lblDeadTitle
                .Caption = vbNewLine & "CHANCE"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call Chance
        Case Is = 36
            With lblDeadTitle
                .Caption = vbNewLine & "FRANSCHHOEK"
                .BackColor = RGB(0, 0, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 37
            With lblDeadTitle
                .Caption = "LUXURY" & vbNewLine & "TAX"
                .BackColor = RGB(190, 240, 230)
                .ForeColor = vbBlack
            End With
            Call LuxuryTax
        Case Is = 38
            With lblDeadTitle
                .Caption = vbNewLine & "CLIFTON"
                .BackColor = RGB(0, 0, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 39
            With lblDeadTitle
                .Caption = vbNewLine & "GO"
                .BackColor = RGB(250, 150, 150)
                .ForeColor = vbBlack
            End With
            Call GoPastBegin
    End Select

End Sub

Private Sub TitleDeadText(Index As Integer)
    Dim Rent As Double
    Dim OneHouse As Double, TwoHouses As Double
    Dim ThreeHouses As Double, FourHouses As Double
    Dim Hotel As Double, Mortgage As Double
    Dim HouseCost As Double
    
    INIfile = Ap & "\data\" & Index & ".ini"
    
    Rent = GetIni("Rent", "Stand")
    OneHouse = GetIni("Rent", "One")
    TwoHouses = GetIni("Rent", "Two")
    ThreeHouses = GetIni("Rent", "Three")
    FourHouses = GetIni("Rent", "Four")
    Hotel = GetIni("Rent", "Hotel")
    
    HouseCost = GetIni("Cost", "Houses")
    
    Mortgage = GetIni("Mortgage", "Value")
    
    lblDeadText.Visible = True
    
    lblDeadText.Caption = "" & _
        "RENT - Site only     " & FormatCurrency(Rent, 2) & vbNewLine & _
        " ''    With 1 House    " & FormatCurrency(OneHouse, 2) & vbNewLine & _
        " ''    With 2 House    " & FormatCurrency(TwoHouses, 2) & vbNewLine & _
        " ''    With 3 House    " & FormatCurrency(ThreeHouses, 2) & vbNewLine & _
        " ''    With 4 House    " & FormatCurrency(FourHouses, 2) & vbNewLine & _
        " ''    With HOTEL     " & FormatCurrency(Hotel, 2) & vbNewLine & _
        " " & vbNewLine & _
        "If a player owns ALL the Sites of any Colour-Group," & _
        "the rent is Doubled on Unimproved Sites in that group." & _
        vbNewLine & vbNewLine & _
        "   COST of Houses, " & FormatCurrency(HouseCost, 2) & vbNewLine & _
        "   COST of Hotels,   " & FormatCurrency(HouseCost, 2) & vbNewLine & vbNewLine & _
        "   MORTGAGE =     " & FormatCurrency(Mortgage, 2)
        
End Sub

Private Sub Chance()

    lblDeadText.Visible = True
    lblDeadText.Caption = "Show random chance card"
    
End Sub

Private Sub CommunityChest()

    lblDeadText.Visible = True
    lblDeadText.Caption = "Show random community chest card"
    
End Sub

Private Sub IncomeTax()

    lblDeadText.Visible = True
    lblDeadText.Caption = "You have to pay Income Tax of R20'000.00"
    
End Sub

Private Sub VisitingJail()

    lblDeadText.Visible = True
    lblDeadText.Caption = "You are only visiting the Jail."
    
End Sub

Private Sub FreeParking()

    lblDeadText.Visible = True
    lblDeadText.Caption = "You are in the Free Parking zone."
    
End Sub

Private Sub GoToJail()

    lblDeadText.Visible = True
    lblDeadText.Caption = "You MUST GO TO JAIL!"
    
End Sub

Private Sub LuxuryTax()

    lblDeadText.Visible = True
    lblDeadText.Caption = "You must pay Luxury Tax of R10'000.00"
    
End Sub

Private Sub GoPastBegin()

    lblDeadText.Visible = True
    lblDeadText.Caption = "You may collect your R20'000.00"
    
End Sub

Private Function GetIni(section As String, key As String)
    Dim R As String
    Dim Worked As Long

    R = String(255, 0)
    Worked = GetPrivateProfileString(section, key, "", R, Len(R), INIfile)
    
    If Worked <> 0 Then
        GetIni = Trim(Left(R, InStr(R, Chr(0)) - 1))
    End If

End Function

Private Sub imgDice_Click(Index As Integer)
    
    mDiceRoll = 0
    
    If imgPlayer1Turn.Visible = True Then
        If mPLayer1InJail = 1 Then
            Call FindPlayer1
            Exit Sub
        End If
    End If

    If imgPlayer2Turn.Visible = True Then
        If mPLayer2InJail = 1 Then
            Call FindPlayer2
            Exit Sub
        End If
    End If
    
    tmrDiceRoll.Enabled = True
    
End Sub

Private Sub imgViewOwned_Click()
    Dim x As Integer
    
    tmrViewOwner.Enabled = True
    
    For x = 0 To 39
        If imgDead(x).Tag > 0 Then
            If imgDead(x).Tag = 1 Then
                imgOwner(x).Picture = LoadPicture(Ap & imgPlayer1Start.ToolTipText & "\Right1.ico")
                imgOwner(x).Visible = True
                imgOwner(x).BorderStyle = 1
            End If
            If imgDead(x).Tag = 2 Then
                imgOwner(x).Picture = LoadPicture(Ap & imgPlayer2Start.ToolTipText & "\Right1.ico")
                imgOwner(x).Visible = True
                imgOwner(x).BorderStyle = 1
            End If
        End If
    Next x
        
End Sub

Private Sub imgViewOwned_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    imgViewOwned.Picture = LoadPicture(Ap & "Button2.gif")

End Sub

Private Sub imgViewOwned_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    imgViewOwned.Picture = LoadPicture(Ap & "Button1.gif")

End Sub

Private Sub imgHideOwner_Click()
    Dim x As Integer
    
    tmrViewOwner.Enabled = False
    For x = 0 To 39
        imgOwner(x).Visible = False
    Next x
        
End Sub

Private Sub imgHideOwner_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    imgHideOwner.Picture = LoadPicture(Ap & "Button4.gif")

End Sub

Private Sub imgHideOwner_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    imgHideOwner.Picture = LoadPicture(Ap & "Button3.gif")

End Sub

Private Sub tmrDiceRoll_Timer()
    Dim Dice1 As Integer
    Dim Dice2 As Integer
    
    Call Randomize
    
    Dice1 = 1 + Int(Rnd() * 6)
    Dice2 = 1 + Int(Rnd() * 6)

    With imgDice
        .Item(0).Picture = LoadPicture(Ap & "dice\" & Dice1 & ".ico")
        .Item(1).Picture = LoadPicture(Ap & "dice\" & Dice2 & ".ico")
    End With
    
    mDiceRoll = mDiceRoll + 1
    mDice1Total = Dice1
    mDice2Total = Dice2
    mDiceTotal = mDice1Total + mDice2Total
    

    If mDiceRoll = 10 Then
        
        If imgPlayer1Turn.Visible = True Then
            
            If mDice1Total = mDice2Total Then
                imgPlayer1Turn.Visible = True
                imgPlayer2Turn.Visible = False
                Call FindPlayer1
            Else
                imgPlayer1Turn.Visible = False
                imgPlayer2Turn.Visible = True
                Call FindPlayer1
            End If
            
        ElseIf imgPlayer2Turn.Visible = True Then
            
            If mDice1Total = mDice2Total Then
                imgPlayer1Turn.Visible = False
                imgPlayer2Turn.Visible = True
                Call FindPlayer2
            Else
                imgPlayer1Turn.Visible = True
                imgPlayer2Turn.Visible = False
                Call FindPlayer2
            End If
            
        End If
        
        mDiceRoll = 0
        tmrDiceRoll.Enabled = False
    End If
    
End Sub

Private Sub FindPlayer1()
    Dim coordinates As Double
    
    coordinates = imgPlayer1.Left & imgPlayer1.Top
        
    Select Case coordinates
        Case Is = 22035
            imgPlayer1Turn.Visible = False
            imgPlayer2Turn.Visible = True
            mPLayer1InJail = 1
            frmMessageJail.Visible = True
        Case Is = 164381
            Player1m0to1.Enabled = True
        Case Is = 164342
            Player1m1to2.Enabled = True
        Case Is = 164303
            Player1m2to3.Enabled = True
        Case Is = 164264
            Player1m3to4.Enabled = True
        Case Is = 164225
            Player1m4to5.Enabled = True
        Case Is = 164186
            Player1m5to6.Enabled = True
        Case Is = 164147
            Player1m6to7.Enabled = True
        Case Is = 164108
            Player1m7to8.Enabled = True
        Case Is = 16469
            Player1m8to9.Enabled = True
        Case Is = 1640
            Player1m9to10.Enabled = True
        Case Is = 2510
            Player1m10to11.Enabled = True
        Case Is = 3010
            Player1m11to12.Enabled = True
        Case Is = 3510
            Player1m12to13.Enabled = True
        Case Is = 4010
            Player1m13to14.Enabled = True
        Case Is = 4510
            Player1m14to15.Enabled = True
        Case Is = 5010
            Player1m15to16.Enabled = True
        Case Is = 5510
            Player1m16to17.Enabled = True
        Case Is = 6010
            Player1m17to18.Enabled = True
        Case Is = 6510
            Player1m18to19.Enabled = True
        Case Is = 7510
            Player1m19to20.Enabled = True
        Case Is = 75169
            Player1m20to21.Enabled = True
        Case Is = 751108
            Player1m21to22.Enabled = True
        Case Is = 751147
            Player1m22to23.Enabled = True
        Case Is = 751186
            Player1m23to24.Enabled = True
        Case Is = 751225
            Player1m24to25.Enabled = True
        Case Is = 751264
            Player1m25to26.Enabled = True
        Case Is = 751303
            Player1m26to27.Enabled = True
        Case Is = 751342
            Player1m27to28.Enabled = True
        Case Is = 751381
            Player1m28to29.Enabled = True
        Case Is = 751465
            Player1m29to30.Enabled = True
        Case Is = 651465
            Player1m30to31.Enabled = True
        Case Is = 601465
            Player1m31to32.Enabled = True
        Case Is = 551465
            Player1m32to33.Enabled = True
        Case Is = 501465
            Player1m33to34.Enabled = True
        Case Is = 451465
            Player1m34to35.Enabled = True
        Case Is = 401465
            Player1m35to36.Enabled = True
        Case Is = 351465
            Player1m36to37.Enabled = True
        Case Is = 301465
            Player1m37to38.Enabled = True
        Case Is = 251465
            Player1m38to39.Enabled = True
        Case Is = 164465
            Player1m39to0.Enabled = True
    End Select
    
End Sub

Private Sub Player1m39to0_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 12

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m39to0.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m0to1.Enabled = True
            Else
                With imgPlayer1
                .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
                .Left = 164
                .Top = 381
                .ToolTipText = 0
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m0to1_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m0to1.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m1to2.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
                    .Left = 164
                    .Top = 342
                    .ToolTipText = 1
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m1to2_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m1to2.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m2to3.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
                    .Left = 164
                    .Top = 303
                    .ToolTipText = 2
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m2to3_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m2to3.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m3to4.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
                    .Left = 164
                    .Top = 264
                    .ToolTipText = 3
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m3to4_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m3to4.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m4to5.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
                    .Left = 164
                    .Top = 225
                    .ToolTipText = 4
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m4to5_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m4to5.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m5to6.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
                    .Left = 164
                    .Top = 186
                    .ToolTipText = 5
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m5to6_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m5to6.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m6to7.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
                    .Left = 164
                    .Top = 147
                    .ToolTipText = 6
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m6to7_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m6to7.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m7to8.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
                    .Left = 164
                    .Top = 108
                    .ToolTipText = 7
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m7to8_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m7to8.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m8to9.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
                    .Left = 164
                    .Top = 69
                    .ToolTipText = 8
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m8to9_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 9.8

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m8to9.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m9to10.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
                    .Left = 164
                    .Top = 0
                    .ToolTipText = 9
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m9to10_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 12.4

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m9to10.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m10to11.Enabled = True
            Else
                With imgPlayer1
                   .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
                   .Left = 251
                   .Top = 0
                   .ToolTipText = 10
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m10to11_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 7

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m10to11.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m11to12.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
                    .Left = 301
                    .Top = 0
                    .ToolTipText = 11
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m11to12_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 7

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m11to12.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m12to13.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
                    .Left = 351
                    .Top = 0
                    .ToolTipText = 12
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m12to13_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 7

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m12to13.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m13to14.Enabled = True
            Else
                With imgPlayer1
                    .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
                    .Left = 401
                    .Top = 0
                    .ToolTipText = 13
                End With
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m13to14_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 7

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
            .Left = 451
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m13to14.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m14to15.Enabled = True
            Else
                imgPlayer1.ToolTipText = 14
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m14to15_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 7

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
            .Left = 501
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m14to15.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m15to16.Enabled = True
            Else
                imgPlayer1.ToolTipText = 15
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m15to16_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 7

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
            .Left = 551
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m15to16.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m16to17.Enabled = True
            Else
                imgPlayer1.ToolTipText = 16
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m16to17_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 7

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
            .Left = 601
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m16to17.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m17to18.Enabled = True
            Else
                imgPlayer1.ToolTipText = 17
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m17to18_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 7

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
            .Left = 651
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m17to18.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m18to19.Enabled = True
            Else
                imgPlayer1.ToolTipText = 18
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m18to19_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left + 14.2

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m18to19.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m19to20.Enabled = True
            Else
                imgPlayer1.ToolTipText = 19
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m19to20_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 9.8

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 69
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m19to20.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m20to21.Enabled = True
            Else
                imgPlayer1.ToolTipText = 20
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m20to21_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 108
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m20to21.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m21to22.Enabled = True
            Else
                imgPlayer1.ToolTipText = 21
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m21to22_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 147
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m21to22.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m22to23.Enabled = True
            Else
                imgPlayer1.ToolTipText = 22
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m22to23_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 186
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m22to23.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m23to24.Enabled = True
            Else
                imgPlayer1.ToolTipText = 23
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m23to24_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 225
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m23to24.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m24to25.Enabled = True
            Else
                imgPlayer1.ToolTipText = 24
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m24to25_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 264
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m24to25.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m25to26.Enabled = True
            Else
                imgPlayer1.ToolTipText = 25
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m25to26_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 303
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m25to26.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m26to27.Enabled = True
            Else
                imgPlayer1.ToolTipText = 26
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m26to27_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 342
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m26to27.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m27to28.Enabled = True
            Else
                imgPlayer1.ToolTipText = 27
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m27to28_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 5.5

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down1.ico")
            .Left = 751
            .Top = 381
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m27to28.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m28to29.Enabled = True
            Else
                imgPlayer1.ToolTipText = 28
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m28to29_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top + 12

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 751
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m28to29.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m29to30.Enabled = True
            Else
                imgPlayer1.ToolTipText = 29
                mPLayer1InJail = 1
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m29to30_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 14.2

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 651
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m29to30.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m30to31.Enabled = True
            Else
                imgPlayer1.ToolTipText = 30
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m30to31_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 7.14

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 601
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m30to31.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m31to32.Enabled = True
            Else
                imgPlayer1.ToolTipText = 31
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m31to32_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 7.14

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 551
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m31to32.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m32to33.Enabled = True
            Else
                imgPlayer1.ToolTipText = 32
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m32to33_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 7.14

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 501
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m32to33.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m33to34.Enabled = True
            Else
                imgPlayer1.ToolTipText = 33
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m33to34_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 7.14

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 451
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m33to34.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m34to35.Enabled = True
            Else
                imgPlayer1.ToolTipText = 34
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m34to35_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 7.14

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 401
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m34to35.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m35to36.Enabled = True
            Else
                imgPlayer1.ToolTipText = 35
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m35to36_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 7.14

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 351
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m35to36.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m36to37.Enabled = True
            Else
                imgPlayer1.ToolTipText = 36
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m36to37_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 7.14

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 301
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m36to37.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m37to38.Enabled = True
            Else
                imgPlayer1.ToolTipText = 37
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m37to38_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 7.14

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left1.ico")
            .Left = 251
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m37to38.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m38to39.Enabled = True
            Else
                imgPlayer1.ToolTipText = 38
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub Player1m38to39_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer1.Left = imgPlayer1.Left - 12.4

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mPLayerBank(1) = mPLayerBank(1) + 20000
        lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(1), 2)
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Up1.ico")
            .Left = 164
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player1m38to39.Enabled = False
            If mDiceTotal <> 0 Then
                Player1m39to0.Enabled = True
            Else
                imgPlayer1.ToolTipText = 39
                Call Player1GameCont(imgPlayer1.Left, imgPlayer1.Top)
            End If
    End If
    
End Sub

Private Sub FindPlayer2()
    Dim coordinates As Double
    
    coordinates = imgPlayer2.Left & imgPlayer2.Top
        
    Select Case coordinates
        Case Is = 22035
            imgPlayer1Turn.Visible = True
            imgPlayer2Turn.Visible = False
            mPLayer2InJail = 1
            frmMessageJailP2.Visible = True
            frmMessageJailP2.Timer3.Enabled = True
        Case Is = 164381
            Player2m0to1.Enabled = True
        Case Is = 164342
            Player2m1to2.Enabled = True
        Case Is = 164303
            Player2m2to3.Enabled = True
        Case Is = 164264
            Player2m3to4.Enabled = True
        Case Is = 164225
            Player2m4to5.Enabled = True
        Case Is = 164186
            Player2m5to6.Enabled = True
        Case Is = 164147
            Player2m6to7.Enabled = True
        Case Is = 164108
            Player2m7to8.Enabled = True
        Case Is = 16469
            Player2m8to9.Enabled = True
        Case Is = 1640
            Player2m9to10.Enabled = True
        Case Is = 2510
            Player2m10to11.Enabled = True
        Case Is = 3010
            Player2m11to12.Enabled = True
        Case Is = 3510
            Player2m12to13.Enabled = True
        Case Is = 4010
            Player2m13to14.Enabled = True
        Case Is = 4510
            Player2m14to15.Enabled = True
        Case Is = 5010
            Player2m15to16.Enabled = True
        Case Is = 5510
            Player2m16to17.Enabled = True
        Case Is = 6010
            Player2m17to18.Enabled = True
        Case Is = 6510
            Player2m18to19.Enabled = True
        Case Is = 7510
            Player2m19to20.Enabled = True
        Case Is = 75169
            Player2m20to21.Enabled = True
        Case Is = 751108
            Player2m21to22.Enabled = True
        Case Is = 751147
            Player2m22to23.Enabled = True
        Case Is = 751186
            Player2m23to24.Enabled = True
        Case Is = 751225
            Player2m24to25.Enabled = True
        Case Is = 751264
            Player2m25to26.Enabled = True
        Case Is = 751303
            Player2m26to27.Enabled = True
        Case Is = 751342
            Player2m27to28.Enabled = True
        Case Is = 751381
            Player2m28to29.Enabled = True
        Case Is = 751465
            Player2m29to30.Enabled = True
        Case Is = 651465
            Player2m30to31.Enabled = True
        Case Is = 601465
            Player2m31to32.Enabled = True
        Case Is = 551465
            Player2m32to33.Enabled = True
        Case Is = 501465
            Player2m33to34.Enabled = True
        Case Is = 451465
            Player2m34to35.Enabled = True
        Case Is = 401465
            Player2m35to36.Enabled = True
        Case Is = 351465
            Player2m36to37.Enabled = True
        Case Is = 301465
            Player2m37to38.Enabled = True
        Case Is = 251465
            Player2m38to39.Enabled = True
        Case Is = 164465
            Player2m39to0.Enabled = True
    End Select
    
End Sub

Private Sub Player2m39to0_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 12

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 381
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m39to0.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m0to1.Enabled = True
            Else
                imgPlayer2.ToolTipText = 0
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m0to1_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 342
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m0to1.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m1to2.Enabled = True
            Else
                imgPlayer2.ToolTipText = 1
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m1to2_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 303
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m1to2.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m2to3.Enabled = True
            Else
                imgPlayer2.ToolTipText = 2
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m2to3_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 264
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m2to3.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m3to4.Enabled = True
            Else
                imgPlayer2.ToolTipText = 3
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m3to4_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 225
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m3to4.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m4to5.Enabled = True
            Else
                imgPlayer2.ToolTipText = 4
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m4to5_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 186
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m4to5.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m5to6.Enabled = True
            Else
                imgPlayer2.ToolTipText = 5
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m5to6_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 147
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m5to6.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m6to7.Enabled = True
            Else
                imgPlayer2.ToolTipText = 6
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m6to7_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 108
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m6to7.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m7to8.Enabled = True
            Else
                imgPlayer2.ToolTipText = 7
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m7to8_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 69
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m7to8.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m8to9.Enabled = True
            Else
                imgPlayer2.ToolTipText = 8
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m8to9_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 9.8

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 164
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m8to9.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m9to10.Enabled = True
            Else
                imgPlayer2.ToolTipText = 9
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m9to10_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 12.4

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 251
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m9to10.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m10to11.Enabled = True
            Else
                imgPlayer2.ToolTipText = 10
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m10to11_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 7

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 301
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m10to11.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m11to12.Enabled = True
            Else
                imgPlayer2.ToolTipText = 11
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m11to12_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 7

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 351
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m11to12.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m12to13.Enabled = True
            Else
                imgPlayer2.ToolTipText = 12
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m12to13_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 7

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 401
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m12to13.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m13to14.Enabled = True
            Else
                imgPlayer2.ToolTipText = 13
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m13to14_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 7

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 451
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m13to14.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m14to15.Enabled = True
            Else
                imgPlayer2.ToolTipText = 14
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m14to15_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 7

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 501
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m14to15.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m15to16.Enabled = True
            Else
                imgPlayer2.ToolTipText = 15
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m15to16_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 7

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 551
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m15to16.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m16to17.Enabled = True
            Else
                imgPlayer2.ToolTipText = 16
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m16to17_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 7

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 601
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m16to17.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m17to18.Enabled = True
            Else
                imgPlayer2.ToolTipText = 17
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m17to18_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 7

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 651
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m17to18.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m18to19.Enabled = True
            Else
                imgPlayer2.ToolTipText = 18
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m18to19_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left + 14.2

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 0
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m18to19.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m19to20.Enabled = True
            Else
                imgPlayer2.ToolTipText = 19
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m19to20_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 9.8

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 69
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m19to20.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m20to21.Enabled = True
            Else
                imgPlayer2.ToolTipText = 20
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m20to21_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 108
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m20to21.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m21to22.Enabled = True
            Else
                imgPlayer2.ToolTipText = 21
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m21to22_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 147
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m21to22.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m22to23.Enabled = True
            Else
                imgPlayer2.ToolTipText = 22
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m22to23_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 186
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m22to23.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m23to24.Enabled = True
            Else
                imgPlayer2.ToolTipText = 23
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m23to24_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 225
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m23to24.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m24to25.Enabled = True
            Else
                imgPlayer2.ToolTipText = 24
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m24to25_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 264
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m24to25.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m25to26.Enabled = True
            Else
                imgPlayer2.ToolTipText = 25
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m25to26_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 303
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m25to26.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m26to27.Enabled = True
            Else
                imgPlayer2.ToolTipText = 26
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m26to27_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 342
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m26to27.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m27to28.Enabled = True
            Else
                imgPlayer2.ToolTipText = 27
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m27to28_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 5.5

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down1.ico")
            .Left = 751
            .Top = 381
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m27to28.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m28to29.Enabled = True
            Else
                imgPlayer2.ToolTipText = 28
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m28to29_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top + 12

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Down" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 751
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m28to29.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m29to30.Enabled = True
            Else
                imgPlayer2.ToolTipText = 29
                mPLayer2InJail = 1
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m29to30_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 14.2

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 651
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m29to30.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m30to31.Enabled = True
            Else
                imgPlayer2.ToolTipText = 30
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m30to31_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 7.14

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 601
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m30to31.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m31to32.Enabled = True
            Else
                imgPlayer2.ToolTipText = 31
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m31to32_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 7.14

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 551
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m31to32.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m32to33.Enabled = True
            Else
                imgPlayer2.ToolTipText = 32
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m32to33_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 7.14

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 501
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m32to33.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m33to34.Enabled = True
            Else
                imgPlayer2.ToolTipText = 33
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m33to34_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 7.14

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 451
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m33to34.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m34to35.Enabled = True
            Else
                imgPlayer2.ToolTipText = 34
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m34to35_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 7.14

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 401
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m34to35.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m35to36.Enabled = True
            Else
                imgPlayer2.ToolTipText = 35
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m35to36_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 7.14

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 351
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m35to36.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m36to37.Enabled = True
            Else
                imgPlayer2.ToolTipText = 36
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m36to37_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 7.14

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 301
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m36to37.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m37to38.Enabled = True
            Else
                imgPlayer2.ToolTipText = 37
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m37to38_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 7.14

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left1.ico")
            .Left = 251
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m37to38.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m38to39.Enabled = True
            Else
                imgPlayer2.ToolTipText = 38
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player2m38to39_Timer()
    
    mMoveCount = mMoveCount + 1
    
    imgPlayer2.Left = imgPlayer2.Left - 12.4

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mPLayerBank(2) = mPLayerBank(2) + 20000
        lblPlayerBank(2).Caption = FormatCurrency(mPLayerBank(2), 2)
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Up1.ico")
            .Left = 164
            .Top = 465
        End With
        mMoveCount = 0
        mDiceTotal = mDiceTotal - 1
        Player2m38to39.Enabled = False
            If mDiceTotal <> 0 Then
                Player2m39to0.Enabled = True
            Else
                imgPlayer2.ToolTipText = 39
                Call Player2GameCont(imgPlayer2.Left, imgPlayer2.Top)
            End If
    End If
    
End Sub

Private Sub Player1GameCont(mLeft As Integer, mTop As Integer)
    Dim x As Integer
    
    mLandedProperty = imgPlayer1.ToolTipText
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ''''' LANDED ON PROPERTY '''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    For x = 1 To 28
        If imgPlayer1.ToolTipText = mPropertyListing(x) Then
            Call LandedOnProperty(1, mPropertyListing(x))
        End If
    Next x
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ''''' INCOME TAX         '''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    If imgPlayer1.ToolTipText = 3 Then
        lsbOne.AddItem ("Player 1 pays R20'000 Income Tax"), lsbOne.ListCount = 0
        mPLayerBank(1) = mPLayerBank(1) - 20000
        lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(1), 2)
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ''''' LUXURY TAX         '''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    If imgPlayer1.ToolTipText = 37 Then
        lsbOne.AddItem ("Player 1 pays R10'000 Luxury Tax"), lsbOne.ListCount = 0
        mPLayerBank(1) = mPLayerBank(1) - 10000
        lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(1), 2)
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ''''' GO TO JAIL         '''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    If imgPlayer1.ToolTipText = 29 Then
        lsbOne.AddItem ("Player 1 Has to go to Jail"), lsbOne.ListCount = 0
        Animation1.Visible = True
        Animation1.Open (Ap & "Policelight.avi")
        Animation1.Play
        tmrPlayer1ToJail.Enabled = True
        mPLayer1InJail = 1
        imgPlayer1Turn.Visible = False
        imgPlayer2Turn.Visible = True
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ''''' LAND ON CUMMUNITY CHEST ''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    If imgPlayer1.ToolTipText = 1 Then
        Call frmCommunityChest.CommunityChestInfo(1, mPLayerBank(1))
        frmCommunityChest.Visible = True
    End If
    
End Sub

Private Sub Player2GameCont(mLeft As Integer, mTop As Integer)

    ''''''''''''''''''''''''''''''''''''''''''''
    ''''' INCOME TAX         '''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    If imgPlayer2.ToolTipText = 3 Then
        lsbOne.AddItem ("Player 2 pays R20'000 Income Tax"), lsbOne.ListCount = 0
        mPLayerBank(2) = mPLayerBank(2) - 20000
        lblPlayerBank(2).Caption = FormatCurrency(mPLayerBank(2), 2)
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''
    ''''' LUXURY TAX         '''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    If imgPlayer2.ToolTipText = 37 Then
        lsbOne.AddItem ("Player 2 pays R10'000 Luxury Tax"), lsbOne.ListCount = 0
        mPLayerBank(2) = mPLayerBank(2) - 10000
        lblPlayerBank(2).Caption = FormatCurrency(mPLayerBank(2), 2)
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''
    ''''' GO TO JAIL         '''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    If imgPlayer2.ToolTipText = 29 Then
        lsbOne.AddItem ("Player 2 Has to go to Jail"), lsbOne.ListCount = 0
        Animation1.Visible = True
        Animation1.Open (Ap & "Policelight.avi")
        Animation1.Play
        tmrPlayer2ToJail.Enabled = True
        mPLayer2InJail = 1
        imgPlayer1Turn.Visible = True
        imgPlayer2Turn.Visible = False
        Exit Sub
    End If

    frmCPUThinking.Timer1.Enabled = True
    frmCPUThinking.Visible = True
    
End Sub

Public Function Player2GameCont2(mLeft As Integer, mTop As Integer)
    Dim x As Integer
    
    mLandedProperty = imgPlayer2.ToolTipText
    
    ''''''''''''''''''''''''''''''''''''''''''''
    ''''' LANDED ON PROPERTY '''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''
    For x = 1 To 28
        If imgPlayer2.ToolTipText = mPropertyListing(x) Then
            Call LandedOnProperty(2, mPropertyListing(x))
        End If
    Next x
        
End Function

Private Sub LandedOnProperty(Player As Integer, Property As Integer)
    Dim Owner As Integer
    
    Owner = imgDead(Property).Tag
    
    If Owner = Player Then
        ''''' DO NO ACTION FOR NOW
    ElseIf Owner <> Player Then
        If Owner = 0 Then ''''''COMPLETED
            INIfile = Ap & "\data\" & Property & ".ini"
            lblMsgPlayer.Caption = "Player " & Player
            lblMsgPropertyName.Caption = GetIni("Dead", "Name")
            lblMsgPropertyPrice.Caption = FormatCurrency(GetIni("Buy", "Cost"), 2)
            mPropertyValue = GetIni("Buy", "Cost")
            Call DisplayPurchaseMessage
        Else
            Call PayTheRent(Player, Property, Owner)
        End If
    End If
    
End Sub

Private Sub cmdMsgYes_Click()
    Dim Player As Integer, Price As Double
    
    Price = lblMsgPropertyPrice.Caption
    Player = Mid$(lblMsgPlayer.Caption, 8, 1)
    Call MsgYes(Player, Price)
    
End Sub

Public Function MsgYes(Player As Integer, Price As Double)
    Dim PropertyName As String, x As Integer
            
    INIfile = Ap & "\data\" & mLandedProperty & ".ini"
    PropertyName = GetIni("Dead", "Name")

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''SET RENT PROPERTY FOR UTILITIES'''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If mLandedProperty = 11 Then
        If imgDead(11).Tag = imgDead(24).Tag Then
            mPLayerBank(Player) = mPLayerBank(Player) - Price
            lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
            imgDead(mLandedProperty).Tag = Player
            imgDead(mLandedProperty).WhatsThisHelpID = 8
            lsbOne.AddItem ("Player " & Player & " has purchased " & PropertyName), lsbOne.ListCount = 0
            Call HidePurchaseMessage
        Else
            mPLayerBank(Player) = mPLayerBank(Player) - Price
            lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
            imgDead(mLandedProperty).Tag = Player
            imgDead(mLandedProperty).WhatsThisHelpID = 7
            lsbOne.AddItem ("Player " & Player & " has purchased " & PropertyName), lsbOne.ListCount = 0
            Call HidePurchaseMessage
        End If
        Exit Function
    End If
    
    If mLandedProperty = 24 Then
        If imgDead(11).Tag = imgDead(24).Tag Then
            mPLayerBank(Player) = mPLayerBank(Player) - Price
            lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
            imgDead(mLandedProperty).Tag = Player
            imgDead(mLandedProperty).WhatsThisHelpID = 8
            lsbOne.AddItem ("Player " & Player & " has purchased " & PropertyName), lsbOne.ListCount = 0
            Call HidePurchaseMessage
        Else
            mPLayerBank(Player) = mPLayerBank(Player) - Price
            lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
            imgDead(mLandedProperty).Tag = Player
            imgDead(mLandedProperty).WhatsThisHelpID = 7
            lsbOne.AddItem ("Player " & Player & " has purchased " & PropertyName), lsbOne.ListCount = 0
            Call HidePurchaseMessage
        End If
        Exit Function
    End If

    ''''''''''''SET RENT PROPERTY END
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''SET RENT PROPERTY FOR STATIONS '''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For x = 4 To 34 Step 10
    If mLandedProperty = x Then
    
    If mLandedProperty = 4 Then
        mPLayerBank(Player) = mPLayerBank(Player) - Price
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        imgDead(mLandedProperty).Tag = Player
        mStation(1) = Player
        lsbOne.AddItem ("Player " & Player & " has purchased " & PropertyName), lsbOne.ListCount = 0
    End If
    
    If mLandedProperty = 14 Then
        mPLayerBank(Player) = mPLayerBank(Player) - Price
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        imgDead(mLandedProperty).Tag = Player
        mStation(2) = Player
        lsbOne.AddItem ("Player " & Player & " has purchased " & PropertyName), lsbOne.ListCount = 0
    End If
        
    If mLandedProperty = 24 Then
        mPLayerBank(Player) = mPLayerBank(Player) - Price
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        imgDead(mLandedProperty).Tag = Player
        mStation(3) = Player
        lsbOne.AddItem ("Player " & Player & " has purchased " & PropertyName), lsbOne.ListCount = 0
    End If
    
    If mLandedProperty = 34 Then
        mPLayerBank(Player) = mPLayerBank(Player) - Price
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        imgDead(mLandedProperty).Tag = Player
        mStation(4) = Player
        lsbOne.AddItem ("Player " & Player & " has purchased " & PropertyName), lsbOne.ListCount = 0
    End If

    Call HidePurchaseMessage
    Call SetStationRent
    Exit Function
    End If
    Next x
    
    ''''''''''''SET RENT PROPERTY END
    
    mPLayerBank(Player) = mPLayerBank(Player) - Price
    lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
    imgDead(mLandedProperty).Tag = Player
    imgDead(mLandedProperty).WhatsThisHelpID = 0
    lsbOne.AddItem ("Player " & Player & " has purchased " & PropertyName), lsbOne.ListCount = 0
    Call HidePurchaseMessage
    Call PropertyGrouping
    
End Function

Private Sub SetStationRent()
    Dim x As Integer
    Dim Stat1 As Integer, Stat2 As Integer
    
    Stat1 = 0
    Stat2 = 0
    
    For x = 1 To 4
        If mStation(x) = 1 Then
             Stat1 = Stat1 + 1
        End If
    Next x
    
    For x = 1 To 4
        If mStation(x) = 2 Then
             Stat2 = Stat2 + 1
        End If
    Next x
    
    For x = 4 To 34 Step 10
        If imgDead(x).Tag = 1 Then
            imgDead(x).WhatsThisHelpID = Stat1
            imgDead(x).ToolTipText = Stat1
        End If
    Next x
    
    For x = 4 To 34 Step 10
        If imgDead(x).Tag = 2 Then
            imgDead(x).WhatsThisHelpID = Stat2
            imgDead(x).ToolTipText = Stat1
        End If
    Next x
    
End Sub
Private Sub cmdMsgNo_Click()
    
    Call HidePurchaseMessage
    frmMessageBuy.lblMsgPropertyName = lblMsgPropertyName
    frmMessageBuy.lblMsgPropertyPrice = lblMsgPropertyPrice
    frmMessageBuy.pgbOne.Min = 0
    frmMessageBuy.pgbOne.Max = Mid$(lblMsgPropertyPrice.Caption, 2)
    frmMessageBuy.pgbOne.Value = Mid$(lblMsgPropertyPrice.Caption, 2) * 0.05
    frmMessageBuy.pgbTwo.Min = 0
    frmMessageBuy.pgbTwo.Max = Mid$(lblMsgPropertyPrice.Caption, 2)
    frmMessageBuy.lblPlayer1Bank = mPLayerBank(1)
    frmMessageBuy.lblPlayer2Bank = mPLayerBank(2)
    Call frmMessageBuy.FormLoadNow
    frmMessageBuy.Show
    
End Sub


Private Sub PayTheRent(Player As Integer, Property As Integer, Owner As Integer)
    Dim Rent As Integer, OtherPlayer As Integer
    
    If Player = 1 Then
        OtherPlayer = 2
    ElseIf Player = 2 Then
        OtherPlayer = 1
    End If
    
    If imgDead(Property).WhatsThisHelpID = 0 Then
        INIfile = Ap & "\data\" & Property & ".ini"
        Rent = GetIni("Rent", "Stand")
        lsbOne.AddItem ("Player " & Player & " paid rent of: " & FormatCurrency(Rent, 2)), lsbOne.ListCount = 0
        mPLayerBank(Player) = mPLayerBank(Player) - Rent
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        mPLayerBank(OtherPlayer) = mPLayerBank(OtherPlayer) + Rent
        lblPlayerBank(OtherPlayer).Caption = FormatCurrency(mPLayerBank(OtherPlayer), 2)
    End If

    If imgDead(Property).WhatsThisHelpID = 1 Then
        INIfile = Ap & "\data\" & Property & ".ini"
        Rent = GetIni("Rent", "One")
        lsbOne.AddItem ("Player " & Player & " paid rent of: " & FormatCurrency(Rent, 2)), lsbOne.ListCount = 0
        mPLayerBank(Player) = mPLayerBank(Player) - Rent
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        mPLayerBank(OtherPlayer) = mPLayerBank(OtherPlayer) + Rent
        lblPlayerBank(OtherPlayer).Caption = FormatCurrency(mPLayerBank(OtherPlayer), 2)
    End If

    If imgDead(Property).WhatsThisHelpID = 2 Then
        INIfile = Ap & "\data\" & Property & ".ini"
        Rent = GetIni("Rent", "Two")
        lsbOne.AddItem ("Player " & Player & " paid rent of: " & FormatCurrency(Rent, 2)), lsbOne.ListCount = 0
        mPLayerBank(Player) = mPLayerBank(Player) - Rent
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        mPLayerBank(OtherPlayer) = mPLayerBank(OtherPlayer) + Rent
        lblPlayerBank(OtherPlayer).Caption = FormatCurrency(mPLayerBank(OtherPlayer), 2)
    End If

    If imgDead(Property).WhatsThisHelpID = 3 Then
        INIfile = Ap & "\data\" & Property & ".ini"
        Rent = GetIni("Rent", "Three")
        lsbOne.AddItem ("Player " & Player & " paid rent of: " & FormatCurrency(Rent, 2)), lsbOne.ListCount = 0
        mPLayerBank(Player) = mPLayerBank(Player) - Rent
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        mPLayerBank(OtherPlayer) = mPLayerBank(OtherPlayer) + Rent
        lblPlayerBank(OtherPlayer).Caption = FormatCurrency(mPLayerBank(OtherPlayer), 2)
    End If

    If imgDead(Property).WhatsThisHelpID = 4 Then
        INIfile = Ap & "\data\" & Property & ".ini"
        Rent = GetIni("Rent", "Four")
        lsbOne.AddItem ("Player " & Player & " paid rent of: " & FormatCurrency(Rent, 2)), lsbOne.ListCount = 0
        mPLayerBank(Player) = mPLayerBank(Player) - Rent
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        mPLayerBank(OtherPlayer) = mPLayerBank(OtherPlayer) + Rent
        lblPlayerBank(OtherPlayer).Caption = FormatCurrency(mPLayerBank(OtherPlayer), 2)
    End If

    If imgDead(Property).WhatsThisHelpID = 5 Then
        INIfile = Ap & "\data\" & Property & ".ini"
        Rent = GetIni("Rent", "Hotel")
        lsbOne.AddItem ("Player " & Player & " paid rent of: " & FormatCurrency(Rent, 2)), lsbOne.ListCount = 0
        mPLayerBank(Player) = mPLayerBank(Player) - Rent
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        mPLayerBank(OtherPlayer) = mPLayerBank(OtherPlayer) + Rent
        lblPlayerBank(OtherPlayer).Caption = FormatCurrency(mPLayerBank(OtherPlayer), 2)
    End If

    If imgDead(Property).WhatsThisHelpID = 6 Then
        INIfile = Ap & "\data\" & Property & ".ini"
        Rent = GetIni("Rent", "Stand") * 2
        lsbOne.AddItem ("Player " & Player & " paid rent of: " & FormatCurrency(Rent, 2)), lsbOne.ListCount = 0
        mPLayerBank(Player) = mPLayerBank(Player) - Rent
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        mPLayerBank(OtherPlayer) = mPLayerBank(OtherPlayer) + Rent
        lblPlayerBank(OtherPlayer).Caption = FormatCurrency(mPLayerBank(OtherPlayer), 2)
    End If

    If imgDead(Property).WhatsThisHelpID = 7 Then
        Rent = mDiceTotal * Int(100)
        lsbOne.AddItem ("Player " & Player & " paid rent of: " & FormatCurrency(Rent, 2)), lsbOne.ListCount = 0
        mPLayerBank(Player) = mPLayerBank(Player) - Rent
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        mPLayerBank(OtherPlayer) = mPLayerBank(OtherPlayer) + Rent
        lblPlayerBank(OtherPlayer).Caption = FormatCurrency(mPLayerBank(OtherPlayer), 2)
    End If

    If imgDead(Property).WhatsThisHelpID = 8 Then
        Rent = mDiceTotal * 200
        lsbOne.AddItem ("Player " & Player & " paid rent of: " & FormatCurrency(Rent, 2)), lsbOne.ListCount = 0
        mPLayerBank(Player) = mPLayerBank(Player) - Rent
        lblPlayerBank(Player).Caption = FormatCurrency(mPLayerBank(Player), 2)
        mPLayerBank(OtherPlayer) = mPLayerBank(OtherPlayer) + Rent
        lblPlayerBank(OtherPlayer).Caption = FormatCurrency(mPLayerBank(OtherPlayer), 2)
    End If

End Sub

Private Sub DisplayPurchaseMessage()
    Dim cpuValue As Long
    Dim PropValue As Long
    
    imgMessage.Visible = True
    lblMsgPlayer.Visible = True
    lblMsgPurchase.Visible = True
    lblMsgNote.Visible = True
    lblMsgMain.Visible = True
    lblMsgPropertyName.Visible = True
    lblMsgPropertyPrice.Visible = True
    PropValue = Mid$(lblMsgPropertyPrice.Caption, 2, 15)
    cmdMsgYes.Visible = True
    cmdMsgNo.Visible = True
    
    If lblMsgPlayer = "Player 2" Then
        If mPLayerBank(2) > mPropertyValue Then
            Call cmdMsgYes_Click
        Else
            Call cmdMsgNo_Click
        End If
    End If
    
End Sub

Private Sub HidePurchaseMessage()

    imgMessage.Visible = False
    lblMsgPlayer.Visible = False
    lblMsgPurchase.Visible = False
    lblMsgNote.Visible = False
    lblMsgMain.Visible = False
    lblMsgPropertyName.Visible = False
    lblMsgPropertyPrice.Visible = False
    lblMsgPropertyPrice.Visible = False
    cmdMsgYes.Visible = False
    cmdMsgNo.Visible = False
    imgDice(0).Visible = True
    imgDice(1).Visible = True
    
End Sub

Private Sub PropertyGrouping()
    Dim Prop1 As Integer, Prop2 As Integer
    Dim Prop3 As Integer, Prop4 As Integer

    '''''''''''''''''''''''''''BROWN
    Prop1 = imgDead(0).Tag
    Prop2 = imgDead(2).Tag
    If Prop1 = Prop2 Then
        If imgDead(0).WhatsThisHelpID = 0 Then
            imgDead(0).WhatsThisHelpID = 6
        End If
        If imgDead(2).WhatsThisHelpID = 0 Then
            imgDead(2).WhatsThisHelpID = 6
        End If
    End If
    
    '''''''''''''''''''''''''''LIGHT BLUE
    Prop1 = imgDead(5).Tag
    Prop2 = imgDead(7).Tag
    Prop3 = imgDead(8).Tag
    Prop4 = 0
    If Prop1 = Prop2 Then
        If Prop2 = Prop3 Then
            If imgDead(5).WhatsThisHelpID = 0 Then
                imgDead(5).WhatsThisHelpID = 6
            End If
            If imgDead(7).WhatsThisHelpID = 0 Then
                imgDead(7).WhatsThisHelpID = 6
            End If
            If imgDead(8).WhatsThisHelpID = 0 Then
                imgDead(8).WhatsThisHelpID = 6
            End If
        End If
    End If

    '''''''''''''''''''''''''''PINK
    Prop1 = imgDead(10).Tag
    Prop2 = imgDead(12).Tag
    Prop3 = imgDead(13).Tag
    Prop4 = 0
    If Prop1 = Prop2 Then
        If Prop2 = Prop3 Then
            If imgDead(10).WhatsThisHelpID = 0 Then
                imgDead(10).WhatsThisHelpID = 6
            End If
            If imgDead(12).WhatsThisHelpID = 0 Then
                imgDead(12).WhatsThisHelpID = 6
            End If
            If imgDead(13).WhatsThisHelpID = 0 Then
                imgDead(13).WhatsThisHelpID = 6
            End If
        End If
    End If

    '''''''''''''''''''''''''''LIGHT BROWN
    Prop1 = imgDead(15).Tag
    Prop2 = imgDead(17).Tag
    Prop3 = imgDead(18).Tag
    Prop4 = 0
    If Prop1 = Prop2 Then
        If Prop2 = Prop3 Then
            If imgDead(15).WhatsThisHelpID = 0 Then
                imgDead(15).WhatsThisHelpID = 6
            End If
            If imgDead(17).WhatsThisHelpID = 0 Then
                imgDead(17).WhatsThisHelpID = 6
            End If
            If imgDead(18).WhatsThisHelpID = 0 Then
                imgDead(18).WhatsThisHelpID = 6
            End If
        End If
    End If

    '''''''''''''''''''''''''''RED
    Prop1 = imgDead(20).Tag
    Prop2 = imgDead(22).Tag
    Prop3 = imgDead(23).Tag
    Prop4 = 0
    If Prop1 = Prop2 Then
        If Prop2 = Prop3 Then
            If imgDead(20).WhatsThisHelpID = 0 Then
                imgDead(20).WhatsThisHelpID = 6
            End If
            If imgDead(22).WhatsThisHelpID = 0 Then
                imgDead(22).WhatsThisHelpID = 6
            End If
            If imgDead(23).WhatsThisHelpID = 0 Then
                imgDead(23).WhatsThisHelpID = 6
            End If
        End If
    End If

    '''''''''''''''''''''''''''YELLOW
    Prop1 = imgDead(25).Tag
    Prop2 = imgDead(26).Tag
    Prop3 = imgDead(28).Tag
    Prop4 = 0
    If Prop1 = Prop2 Then
        If Prop2 = Prop3 Then
            If imgDead(25).WhatsThisHelpID = 0 Then
                imgDead(25).WhatsThisHelpID = 6
            End If
            If imgDead(26).WhatsThisHelpID = 0 Then
                imgDead(26).WhatsThisHelpID = 6
            End If
            If imgDead(28).WhatsThisHelpID = 0 Then
                imgDead(28).WhatsThisHelpID = 6
            End If
        End If
    End If

    '''''''''''''''''''''''''''GREEN
    Prop1 = imgDead(30).Tag
    Prop2 = imgDead(31).Tag
    Prop3 = imgDead(33).Tag
    Prop4 = 0
    If Prop1 = Prop2 Then
        If Prop2 = Prop3 Then
            If imgDead(30).WhatsThisHelpID = 0 Then
                imgDead(30).WhatsThisHelpID = 6
            End If
            If imgDead(31).WhatsThisHelpID = 0 Then
                imgDead(31).WhatsThisHelpID = 6
            End If
            If imgDead(33).WhatsThisHelpID = 0 Then
                imgDead(33).WhatsThisHelpID = 6
            End If
        End If
    End If

    ''''''''''''''''''''''''''''DARK BLUE
    Prop1 = imgDead(36).Tag
    Prop2 = imgDead(38).Tag
    Prop3 = 0
    Prop4 = 0
    If Prop1 = Prop2 Then
        If imgDead(36).WhatsThisHelpID = 0 Then
            imgDead(36).WhatsThisHelpID = 6
        End If
        If imgDead(38).WhatsThisHelpID = 0 Then
            imgDead(38).WhatsThisHelpID = 6
        End If
    End If

End Sub

Private Sub Form_Terminate()
    
    Call UnloadProcedures
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call UnloadProcedures

End Sub

Private Sub UnloadProcedures()
    
    Call Unload(frmStart)
    Call Unload(frmBoard)
    End
    
End Sub

Private Sub tmrPlayer1ToJail_Timer()
    Dim Bank As Long
    
    Bank = mPLayerBank(1)
    
    mMoveCount = mMoveCount + 1
    mJailCount = mJailCount + 1
    
    imgPlayer1.Top = imgPlayer1.Top - 3.07
    imgPlayer1.Left = imgPlayer1.Left - 3.79

    imgPlayer1.Picture = LoadPicture(Ap & imgPlayer1.Tag & "\left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
    End If

    If mJailCount = 140 Then
        mJailCount = 0
        tmrPlayer1ToJail.Enabled = False
        Animation1.Stop
        Animation1.Visible = False
        Call frmMessageJail.Info(1, Bank)
        With imgPlayer1
            .Picture = LoadPicture(Ap & imgPlayer1.Tag & "\Right1.ico")
            .Left = 220
            .Top = 35
        End With
    End If

End Sub

Public Sub Player1OutOfJail(Dice1 As Integer, Dice2 As Integer, Player As Integer, Bank As Long)

    mPLayerBank(1) = Bank
    lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(1), 2)
    mDice1Total = Dice1
    mDice2Total = Dice2
    mDiceTotal = mDice1Total + mDice2Total
    mPLayer1InJail = 0
    
    If Player = 1 Then
            
        If mDice1Total = mDice2Total Then
            imgPlayer1Turn.Visible = True
            imgPlayer2Turn.Visible = False
            Call FindPlayer1
        Else
            imgPlayer1Turn.Visible = False
            imgPlayer2Turn.Visible = True
            Call FindPlayer1
        End If
        
        mDiceRoll = 0
        tmrDiceRoll.Enabled = False
        
    End If

End Sub

Private Sub tmrPlayer2ToJail_Timer()
    Dim Bank As Long
    
    Bank = mPLayerBank(2)
    
    mMoveCount = mMoveCount + 1
    mJailCount = mJailCount + 1
    
    imgPlayer2.Top = imgPlayer2.Top - 3.07
    imgPlayer2.Left = imgPlayer2.Left - 3.79

    imgPlayer2.Picture = LoadPicture(Ap & imgPlayer2.Tag & "\left" & mMoveCount & ".ico")
    
    If mMoveCount = 7 Then
        mMoveCount = 0
    End If

    If mJailCount = 140 Then
        mJailCount = 0
        tmrPlayer2ToJail.Enabled = False
        Animation1.Stop
        Animation1.Visible = False
        Call frmMessageJailP2.Info(1, Bank)
        With imgPlayer2
            .Picture = LoadPicture(Ap & imgPlayer2.Tag & "\Right1.ico")
            .Left = 220
            .Top = 35
        End With
    End If


End Sub

Public Sub Player2OutOfJail(Dice1 As Integer, Dice2 As Integer, Player As Integer, Bank As Long)

    mPLayerBank(2) = Bank
    lblPlayerBank(2).Caption = FormatCurrency(mPLayerBank(2), 2)
    mDice1Total = Dice1
    mDice2Total = Dice2
    mDiceTotal = mDice1Total + mDice2Total
    mPLayer2InJail = 0
    
    If Player = 2 Then
            
        If mDice1Total = mDice2Total Then
            imgPlayer1Turn.Visible = False
            imgPlayer2Turn.Visible = True
            Call FindPlayer2
        Else
            imgPlayer1Turn.Visible = True
            imgPlayer2Turn.Visible = False
            Call FindPlayer2
        End If
        
        mDiceRoll = 0
        tmrDiceRoll.Enabled = False
        
    End If

End Sub

Private Sub tmrViewOwner_Timer()
    Dim x As Integer
    
    For x = 0 To 39
        If imgOwner(x).Visible = True Then
            If imgOwner(x).BorderStyle = 1 Then
                imgOwner(x).BorderStyle = 0
            Else
                imgOwner(x).BorderStyle = 1
            End If
        End If
    Next x
    
End Sub

Public Function CommunityChestCont(Player As Integer, Bank As Long, Card As Integer)

    Select Case Card
        Case Is = 1
            lsbOne.AddItem ("Player " & Player & " inherited R 10'000.00."), lsbOne.ListCount = 0
            mPLayerBank(Player) = mPLayerBank(Player) + 10000
            lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(Player), 2)
        Case Is = 2
            lsbOne.AddItem ("Player " & Player & " had to pay the Hospital R 10'000.00."), lsbOne.ListCount = 0
            mPLayerBank(Player) = mPLayerBank(Player) - 10000
            lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(Player), 2)
        Case Is = 3
            lsbOne.AddItem ("Player " & Player & " Annuity Matures. R 10'000.00"), lsbOne.ListCount = 0
            mPLayerBank(Player) = mPLayerBank(Player) + 10000
            lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(Player), 2)
        Case Is = 4
            lsbOne.AddItem ("Player " & Player & " received interest. R 2'500.00"), lsbOne.ListCount = 0
            mPLayerBank(Player) = mPLayerBank(Player) + 2500
            lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(Player), 2)
        Case Is = 5
            lsbOne.AddItem ("Player " & Player & " received Get Out Of Jail card."), lsbOne.ListCount = 0
            If Player = 1 Then
                frmMessageJail.imgCCOOJ.Visible = True
            ElseIf Player = 2 Then
                frmMessageJailP2.imgCCOOJ.Visible = True
            End If
        Case Is = 6
            lsbOne.AddItem ("Player " & Player & " Had a Birthday. R 1'000.00"), lsbOne.ListCount = 0
            mPLayerBank(Player) = mPLayerBank(Player) + 1000
            lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(Player), 2)
            If Player = 1 Then
                mPLayerBank(2) = mPLayerBank(2) - 1000
                lblPlayerBank(2).Caption = FormatCurrency(mPLayerBank(2), 2)
            ElseIf Player = 2 Then
                mPLayerBank(1) = mPLayerBank(1) - 1000
                lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(1), 2)
            End If
        Case Is = 7
            lsbOne.AddItem ("Player " & Player & " paid Doctor fees. R 5'000.00"), lsbOne.ListCount = 0
            mPLayerBank(Player) = mPLayerBank(Player) + 5000
            lblPlayerBank(1).Caption = FormatCurrency(mPLayerBank(Player), 2)
        Case Is = 8
            If Player = 1 Then
                If imgPlayer1.ToolTipText = 1 Then
                    mDiceTotal = 30
                    Call FindPlayer1
                ElseIf imgPlayer1.ToolTipText = 16 Then
                    mDiceTotal = 15
                    Call FindPlayer1
                ElseIf imgPlayer1.ToolTipText = 32 Then
                    mDiceTotal = 38
                    Call FindPlayer1
                End If
            End If
            If Player = 2 Then
                If imgPlayer2.ToolTipText = 1 Then
                    mDiceTotal = 30
                    Call FindPlayer2
                ElseIf imgPlayer2.ToolTipText = 16 Then
                    mDiceTotal = 15
                    Call FindPlayer2
                ElseIf imgPlayer2.ToolTipText = 32 Then
                    mDiceTotal = 38
                    Call FindPlayer2
                End If
            End If
            
    End Select
    
End Function
