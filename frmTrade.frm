VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTrade 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   ClientHeight    =   9315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtPropertiesP2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6960
      TabIndex        =   10
      Top             =   6270
      Width           =   1935
   End
   Begin VB.TextBox txtTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2925
      TabIndex        =   8
      Top             =   7395
      Width           =   1935
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   4845
      TabIndex        =   7
      Top             =   6765
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Increment       =   1000
      Max             =   500000
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtCash 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2910
      TabIndex        =   6
      Top             =   6765
      Width           =   1935
   End
   Begin VB.TextBox txtProperties 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2925
      TabIndex        =   2
      Top             =   6285
      Width           =   1935
   End
   Begin Project1.dcButton cmdCancel 
      Height          =   495
      Left            =   8580
      TabIndex        =   0
      Top             =   8685
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   873
      BackColor       =   1228031
      ButtonStyle     =   11
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblP2Name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   6945
      TabIndex        =   13
      Top             =   5940
      Width           =   1920
   End
   Begin VB.Label lblP1Name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   2955
      TabIndex        =   12
      Top             =   5940
      Width           =   1920
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Properties:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   5610
      TabIndex        =   11
      Top             =   6330
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Trade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   7425
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   2925
      X2              =   4875
      Y1              =   7290
      Y2              =   7290
   End
   Begin VB.Line Line1 
      X1              =   2925
      X2              =   4875
      Y1              =   7245
      Y2              =   7245
   End
   Begin VB.Label lblDeadText 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   3225
      Left            =   3825
      TabIndex        =   5
      Top             =   1305
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Label lblDeadTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3825
      TabIndex        =   4
      Top             =   510
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cash:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   2190
      TabIndex        =   3
      Top             =   6810
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Properties:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   1575
      TabIndex        =   1
      Top             =   6345
      Width           =   1320
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   0
      Left            =   6825
      Picture         =   "frmTrade.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "6000"
      Top             =   180
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   16
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   29
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   1020
      Index           =   4
      Left            =   9240
      Picture         =   "frmTrade.frx":02A8
      Stretch         =   -1  'True
      Top             =   3495
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   240
      Index           =   5
      Left            =   6825
      Picture         =   "frmTrade.frx":0AA6
      Stretch         =   -1  'True
      ToolTipText     =   "10000"
      Top             =   585
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   9
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   39
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   10
      Left            =   6825
      Picture         =   "frmTrade.frx":0CF2
      Stretch         =   -1  'True
      ToolTipText     =   "14000"
      Top             =   975
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   1020
      Index           =   11
      Left            =   6825
      Picture         =   "frmTrade.frx":0EFA
      Stretch         =   -1  'True
      ToolTipText     =   "10000"
      Top             =   4635
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   1020
      Index           =   14
      Left            =   6825
      Picture         =   "frmTrade.frx":16DD
      Stretch         =   -1  'True
      Top             =   3495
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   240
      Index           =   15
      Left            =   6825
      Picture         =   "frmTrade.frx":1E92
      Stretch         =   -1  'True
      ToolTipText     =   "18000"
      Top             =   1380
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   37
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   3
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   20
      Left            =   6825
      Picture         =   "frmTrade.frx":20D7
      Stretch         =   -1  'True
      ToolTipText     =   "22000"
      Top             =   1785
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   35
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   1020
      Index           =   24
      Left            =   8430
      Picture         =   "frmTrade.frx":2316
      Stretch         =   -1  'True
      Top             =   3510
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   240
      Index           =   25
      Left            =   6825
      Picture         =   "frmTrade.frx":2ABE
      Stretch         =   -1  'True
      Top             =   2190
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   1020
      Index           =   27
      Left            =   7635
      Picture         =   "frmTrade.frx":2CF9
      Stretch         =   -1  'True
      ToolTipText     =   "10000"
      Top             =   4635
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   1
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   30
      Left            =   6825
      Picture         =   "frmTrade.frx":3451
      Stretch         =   -1  'True
      ToolTipText     =   "30000"
      Top             =   2580
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   32
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   1020
      Index           =   34
      Left            =   7635
      Picture         =   "frmTrade.frx":362C
      Stretch         =   -1  'True
      Top             =   3495
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   240
      Index           =   6
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   36
      Left            =   6825
      Picture         =   "frmTrade.frx":3DD1
      Stretch         =   -1  'True
      Top             =   2985
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   19
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   21
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   2
      Left            =   7695
      Picture         =   "frmTrade.frx":3F55
      Stretch         =   -1  'True
      ToolTipText     =   "6000"
      Top             =   180
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   7
      Left            =   7695
      Picture         =   "frmTrade.frx":41FD
      Stretch         =   -1  'True
      ToolTipText     =   "10000"
      Top             =   585
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   8
      Left            =   8580
      Picture         =   "frmTrade.frx":4449
      Stretch         =   -1  'True
      ToolTipText     =   "12000"
      Top             =   585
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   12
      Left            =   7695
      Picture         =   "frmTrade.frx":4695
      Stretch         =   -1  'True
      ToolTipText     =   "14000"
      Top             =   975
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   13
      Left            =   8580
      Picture         =   "frmTrade.frx":489D
      Stretch         =   -1  'True
      ToolTipText     =   "16000"
      Top             =   975
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   17
      Left            =   7695
      Picture         =   "frmTrade.frx":4AA5
      Stretch         =   -1  'True
      ToolTipText     =   "18000"
      Top             =   1380
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   18
      Left            =   8580
      Picture         =   "frmTrade.frx":4CEA
      Stretch         =   -1  'True
      ToolTipText     =   "20000"
      Top             =   1380
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   22
      Left            =   7695
      Picture         =   "frmTrade.frx":4F2F
      Stretch         =   -1  'True
      ToolTipText     =   "22000"
      Top             =   1785
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   23
      Left            =   8580
      Picture         =   "frmTrade.frx":516E
      Stretch         =   -1  'True
      ToolTipText     =   "24000"
      Top             =   1785
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   26
      Left            =   7695
      Picture         =   "frmTrade.frx":53AD
      Stretch         =   -1  'True
      Top             =   2190
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   28
      Left            =   8580
      Picture         =   "frmTrade.frx":55E8
      Stretch         =   -1  'True
      ToolTipText     =   "28000"
      Top             =   2190
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   31
      Left            =   7695
      Picture         =   "frmTrade.frx":5823
      Stretch         =   -1  'True
      ToolTipText     =   "30000"
      Top             =   2580
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   33
      Left            =   8580
      Picture         =   "frmTrade.frx":59FE
      Stretch         =   -1  'True
      ToolTipText     =   "32000"
      Top             =   2580
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgPropP2 
      Height          =   225
      Index           =   38
      Left            =   7695
      Picture         =   "frmTrade.frx":5BD9
      Stretch         =   -1  'True
      ToolTipText     =   "40000"
      Top             =   2985
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   38
      Left            =   990
      Picture         =   "frmTrade.frx":5D5D
      Stretch         =   -1  'True
      ToolTipText     =   "40000"
      Top             =   3000
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   33
      Left            =   1875
      Picture         =   "frmTrade.frx":5EE1
      Stretch         =   -1  'True
      ToolTipText     =   "32000"
      Top             =   2595
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   31
      Left            =   990
      Picture         =   "frmTrade.frx":60BC
      Stretch         =   -1  'True
      ToolTipText     =   "30000"
      Top             =   2595
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   28
      Left            =   1875
      Picture         =   "frmTrade.frx":6297
      Stretch         =   -1  'True
      ToolTipText     =   "28000"
      Top             =   2205
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   26
      Left            =   990
      Picture         =   "frmTrade.frx":64D2
      Stretch         =   -1  'True
      Top             =   2205
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   23
      Left            =   1875
      Picture         =   "frmTrade.frx":670D
      Stretch         =   -1  'True
      ToolTipText     =   "24000"
      Top             =   1800
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   22
      Left            =   990
      Picture         =   "frmTrade.frx":694C
      Stretch         =   -1  'True
      ToolTipText     =   "22000"
      Top             =   1800
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   18
      Left            =   1875
      Picture         =   "frmTrade.frx":6B8B
      Stretch         =   -1  'True
      ToolTipText     =   "20000"
      Top             =   1395
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   17
      Left            =   990
      Picture         =   "frmTrade.frx":6DD0
      Stretch         =   -1  'True
      ToolTipText     =   "18000"
      Top             =   1395
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   13
      Left            =   1875
      Picture         =   "frmTrade.frx":7015
      Stretch         =   -1  'True
      ToolTipText     =   "16000"
      Top             =   990
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   12
      Left            =   990
      Picture         =   "frmTrade.frx":721D
      Stretch         =   -1  'True
      ToolTipText     =   "14000"
      Top             =   990
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   8
      Left            =   1875
      Picture         =   "frmTrade.frx":7425
      Stretch         =   -1  'True
      ToolTipText     =   "12000"
      Top             =   600
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   7
      Left            =   990
      Picture         =   "frmTrade.frx":7671
      Stretch         =   -1  'True
      ToolTipText     =   "10000"
      Top             =   600
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   2
      Left            =   990
      Picture         =   "frmTrade.frx":78BD
      Stretch         =   -1  'True
      ToolTipText     =   "6000"
      Top             =   195
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   39
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   37
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   36
      Left            =   120
      Picture         =   "frmTrade.frx":7B65
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   240
      Index           =   35
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   1020
      Index           =   34
      Left            =   930
      Picture         =   "frmTrade.frx":7CE9
      Stretch         =   -1  'True
      Top             =   3510
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   32
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   30
      Left            =   120
      Picture         =   "frmTrade.frx":848E
      Stretch         =   -1  'True
      ToolTipText     =   "30000"
      Top             =   2595
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   29
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   1020
      Index           =   27
      Left            =   930
      Picture         =   "frmTrade.frx":8669
      Stretch         =   -1  'True
      ToolTipText     =   "10000"
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   240
      Index           =   25
      Left            =   120
      Picture         =   "frmTrade.frx":8DC1
      Stretch         =   -1  'True
      Top             =   2205
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   1020
      Index           =   24
      Left            =   1725
      Picture         =   "frmTrade.frx":8FFC
      Stretch         =   -1  'True
      Top             =   3510
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   21
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   20
      Left            =   120
      Picture         =   "frmTrade.frx":97A4
      Stretch         =   -1  'True
      ToolTipText     =   "22000"
      Top             =   1800
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   19
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   16
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   240
      Index           =   15
      Left            =   120
      Picture         =   "frmTrade.frx":99E3
      Stretch         =   -1  'True
      ToolTipText     =   "18000"
      Top             =   1395
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   1020
      Index           =   14
      Left            =   120
      Picture         =   "frmTrade.frx":9C28
      Stretch         =   -1  'True
      Top             =   3510
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   1020
      Index           =   11
      Left            =   120
      Picture         =   "frmTrade.frx":A3DD
      Stretch         =   -1  'True
      ToolTipText     =   "10000"
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   10
      Left            =   120
      Picture         =   "frmTrade.frx":ABC0
      Stretch         =   -1  'True
      ToolTipText     =   "14000"
      Top             =   990
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   9
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   6
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   240
      Index           =   5
      Left            =   120
      Picture         =   "frmTrade.frx":ADC8
      Stretch         =   -1  'True
      ToolTipText     =   "10000"
      Top             =   600
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   1020
      Index           =   4
      Left            =   2535
      Picture         =   "frmTrade.frx":B014
      Stretch         =   -1  'True
      Top             =   3510
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   3
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   1
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   4650
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgProp 
      Height          =   225
      Index           =   0
      Left            =   120
      Picture         =   "frmTrade.frx":B812
      Stretch         =   -1  'True
      ToolTipText     =   "6000"
      Top             =   195
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "frmTrade"
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
    
    Dim P1Value As Long
    Dim P1PropValue As Long
    Dim P1CashValue As Long
    Dim P2Value As Long
    Dim P2PropValue As Long
    
Private Sub cmdCancel_Click()
    
    Call ResetForm
    frmTrade.Visible = False
    
End Sub

Private Sub Form_Load()
    
    If Right(App.Path, 1) = "\" Then
        Ap = App.Path
    Else
        Ap = App.Path & "\"
    End If

    lblP1Name.Caption = frmBoard.lblPlayer1Name.Caption
    lblP2Name.Caption = frmBoard.lblPlayer2Name.Caption

End Sub

Private Sub imgProp_Click(Index As Integer)
    Dim x As Integer

    If imgProp(Index).BorderStyle = 1 Then
        imgProp(Index).BorderStyle = 0
        P1PropValue = P1PropValue - imgProp(Index).ToolTipText
        P1Value = P1PropValue + P1CashValue
        txtProperties.Text = FormatCurrency(P1PropValue, 2)
    Else
        imgProp(Index).BorderStyle = 1
        lblDeadText.Visible = True
        lblDeadTitle.Visible = True
        P1PropValue = P1PropValue + imgProp(Index).ToolTipText
        txtProperties.Text = FormatCurrency(P1PropValue, 2)
        P1Value = P1PropValue + P1CashValue
        txtProperties.Text = FormatCurrency(P1PropValue, 2)
        
        Select Case Index
        Case Is = 0
            With lblDeadTitle
                .Caption = vbNewLine & "WESTVILLE"
                .BackColor = RGB(210, 150, 130)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 2
            With lblDeadTitle
                .Caption = vbNewLine & "AMANZIMTOTI"
                .BackColor = RGB(210, 150, 130)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
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
        Case Is = 20
            With lblDeadTitle
                .Caption = vbNewLine & "SOWETO"
                .BackColor = RGB(255, 0, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
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
        Case Is = 36
            With lblDeadTitle
                .Caption = vbNewLine & "FRANSCHHOEK"
                .BackColor = RGB(0, 0, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 38
            With lblDeadTitle
                .Caption = vbNewLine & "CLIFTON"
                .BackColor = RGB(0, 0, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
    End Select
    End If
    
    
    txtTotal.Text = FormatCurrency(P1Value, 2)
    
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
    lblDeadTitle.Visible = True
    
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
        
    If Index = 11 Or Index = 27 Then
        lblDeadText.Caption = "" & _
        "'' Only 1 Utility = 100 x Dice Total" & vbNewLine & _
        " '' Both Utilities = 200 x Dice Total" & vbNewLine & _
        vbNewLine & vbNewLine & _
        "   MORTGAGE =     " & FormatCurrency(Mortgage, 2)
    End If
        
    If Index = 4 Then
        lblDeadText.Caption = "" & _
        "''   Own 1 Station      " & FormatCurrency(OneHouse, 2) & vbNewLine & _
        " ''   Own 2 Station's    " & FormatCurrency(TwoHouses, 2) & vbNewLine & _
        " ''   Own 3 Station's    " & FormatCurrency(ThreeHouses, 2) & vbNewLine & _
        " ''   Own 4 Station's    " & FormatCurrency(FourHouses, 2) & vbNewLine & _
        " " & vbNewLine & _
        "   MORTGAGE =     " & FormatCurrency(Mortgage, 2)
    End If
    
    If Index = 14 Then
        lblDeadText.Caption = "" & _
        "''   Own 1 Station      " & FormatCurrency(OneHouse, 2) & vbNewLine & _
        " ''   Own 2 Station's    " & FormatCurrency(TwoHouses, 2) & vbNewLine & _
        " ''   Own 3 Station's    " & FormatCurrency(ThreeHouses, 2) & vbNewLine & _
        " ''   Own 4 Station's    " & FormatCurrency(FourHouses, 2) & vbNewLine & _
        " " & vbNewLine & _
        "   MORTGAGE =     " & FormatCurrency(Mortgage, 2)
    End If
    
    If Index = 24 Then
        lblDeadText.Caption = "" & _
        "''   Own 1 Station      " & FormatCurrency(OneHouse, 2) & vbNewLine & _
        " ''   Own 2 Station's    " & FormatCurrency(TwoHouses, 2) & vbNewLine & _
        " ''   Own 3 Station's    " & FormatCurrency(ThreeHouses, 2) & vbNewLine & _
        " ''   Own 4 Station's    " & FormatCurrency(FourHouses, 2) & vbNewLine & _
        " " & vbNewLine & _
        "   MORTGAGE =     " & FormatCurrency(Mortgage, 2)
    End If
    
    If Index = 34 Then
        lblDeadText.Caption = "" & _
        "''   Own 1 Station      " & FormatCurrency(OneHouse, 2) & vbNewLine & _
        " ''   Own 2 Station's    " & FormatCurrency(TwoHouses, 2) & vbNewLine & _
        " ''   Own 3 Station's    " & FormatCurrency(ThreeHouses, 2) & vbNewLine & _
        " ''   Own 4 Station's    " & FormatCurrency(FourHouses, 2) & vbNewLine & _
        " " & vbNewLine & _
        "   MORTGAGE =     " & FormatCurrency(Mortgage, 2)
    End If
    
End Sub

Private Sub ResetForm()
    Dim x As Integer
    
    For x = 0 To 39
        imgProp(x).Visible = False
        imgPropP2(x).Visible = False
    Next x
    
    P1Value = 0
    P1PropValue = 0
    P1CashValue = 0
    P2Value = 0
    lblDeadText.Visible = False
    lblDeadTitle.Visible = False

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

Private Sub imgPropP2_Click(Index As Integer)
    Dim x As Integer

    If imgPropP2(Index).BorderStyle = 1 Then
        imgPropP2(Index).BorderStyle = 0
        P2PropValue = P2PropValue - imgPropP2(Index).ToolTipText
        txtPropertiesP2.Text = FormatCurrency(P2PropValue, 2)
    Else
        imgPropP2(Index).BorderStyle = 1
        lblDeadText.Visible = True
        lblDeadTitle.Visible = True
        
        P2PropValue = P2PropValue + imgPropP2(Index).ToolTipText
        txtPropertiesP2.Text = FormatCurrency(P2PropValue, 2)
        
        Select Case Index
        Case Is = 0
            With lblDeadTitle
                .Caption = vbNewLine & "WESTVILLE"
                .BackColor = RGB(210, 150, 130)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 2
            With lblDeadTitle
                .Caption = vbNewLine & "AMANZIMTOTI"
                .BackColor = RGB(210, 150, 130)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
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
        Case Is = 20
            With lblDeadTitle
                .Caption = vbNewLine & "SOWETO"
                .BackColor = RGB(255, 0, 0)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
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
        Case Is = 36
            With lblDeadTitle
                .Caption = vbNewLine & "FRANSCHHOEK"
                .BackColor = RGB(0, 0, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
        Case Is = 38
            With lblDeadTitle
                .Caption = vbNewLine & "CLIFTON"
                .BackColor = RGB(0, 0, 255)
                .ForeColor = vbBlack
            End With
            Call TitleDeadText(Index)
    End Select
    End If
    
End Sub


Private Sub UpDown1_Change()

    txtCash.Text = FormatCurrency(UpDown1.Value, 2)
    P1CashValue = UpDown1.Value
    P1Value = P1PropValue + P1CashValue
    txtTotal.Text = FormatCurrency(P1Value, 2)

End Sub
