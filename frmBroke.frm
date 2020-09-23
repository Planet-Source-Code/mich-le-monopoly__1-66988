VERSION 5.00
Begin VB.Form frmBroke 
   BorderStyle     =   0  'None
   Caption         =   "frmBroke"
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   Picture         =   "frmBroke.frx":0000
   ScaleHeight     =   501
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   751
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrPlayer2Sell2 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   9960
      Top             =   3615
   End
   Begin VB.Timer tmrPlayer2Buy 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   9420
      Top             =   4140
   End
   Begin VB.Timer tmrPlayer2Sell 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   9420
      Top             =   3630
   End
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   24
      Left            =   345
      TabIndex        =   77
      Tag             =   "0"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   27
      Left            =   2550
      TabIndex        =   80
      Tag             =   "0"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   26
      Left            =   1830
      TabIndex        =   79
      Tag             =   "0"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   25
      Left            =   1080
      TabIndex        =   78
      Tag             =   "0"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   23
      Left            =   4050
      TabIndex        =   76
      Tag             =   "0"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdDone 
      Height          =   375
      Left            =   9495
      TabIndex        =   46
      Top             =   6795
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   661
      BackColor       =   1228031
      Caption         =   "Done"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Tag             =   "0"
      ToolTipText     =   "6000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "5000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   2
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "6000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   2
      Left            =   1830
      TabIndex        =   5
      Tag             =   "5"
      ToolTipText     =   "10000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   3
      Left            =   2580
      TabIndex        =   6
      ToolTipText     =   "5000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   7
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   3
      Left            =   2580
      TabIndex        =   7
      Tag             =   "7"
      ToolTipText     =   "10000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   4
      Left            =   3330
      TabIndex        =   9
      Tag             =   "8"
      ToolTipText     =   "12000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   5
      Left            =   4080
      TabIndex        =   10
      ToolTipText     =   "10000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   10
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   5
      Left            =   4080
      TabIndex        =   11
      Tag             =   "10"
      ToolTipText     =   "14000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   6
      Left            =   4830
      TabIndex        =   13
      Tag             =   "12"
      ToolTipText     =   "14000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   7
      Left            =   5580
      TabIndex        =   14
      ToolTipText     =   "10000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   13
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   7
      Left            =   5580
      TabIndex        =   15
      Tag             =   "13"
      ToolTipText     =   "16000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   8
      Left            =   6330
      TabIndex        =   17
      Tag             =   "15"
      ToolTipText     =   "18000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   9
      Left            =   7080
      TabIndex        =   18
      ToolTipText     =   "10000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   17
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   9
      Left            =   7080
      TabIndex        =   19
      Tag             =   "17"
      ToolTipText     =   "18000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   10
      Left            =   7830
      TabIndex        =   21
      Tag             =   "18"
      ToolTipText     =   "20000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   11
      Left            =   330
      TabIndex        =   23
      Tag             =   "20"
      ToolTipText     =   "22000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   12
      Left            =   1095
      TabIndex        =   24
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   22
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   12
      Left            =   1080
      TabIndex        =   25
      Tag             =   "22"
      ToolTipText     =   "22000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   13
      Left            =   1830
      TabIndex        =   27
      Tag             =   "23"
      ToolTipText     =   "24000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   14
      Left            =   2595
      TabIndex        =   28
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   25
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   14
      Left            =   2580
      TabIndex        =   29
      Tag             =   "25"
      ToolTipText     =   "26000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   15
      Left            =   3330
      TabIndex        =   31
      Tag             =   "26"
      ToolTipText     =   "26000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   16
      Left            =   4095
      TabIndex        =   32
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   28
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   16
      Left            =   4080
      TabIndex        =   33
      Tag             =   "28"
      ToolTipText     =   "28000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   17
      Left            =   4830
      TabIndex        =   35
      Tag             =   "30"
      ToolTipText     =   "30000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   18
      Left            =   5595
      TabIndex        =   36
      ToolTipText     =   "20000"
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   31
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   18
      Left            =   5580
      TabIndex        =   37
      Tag             =   "31"
      ToolTipText     =   "30000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   19
      Left            =   6330
      TabIndex        =   39
      Tag             =   "33"
      ToolTipText     =   "32000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   20
      Left            =   7095
      TabIndex        =   40
      ToolTipText     =   "20000"
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   36
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   20
      Left            =   7080
      TabIndex        =   41
      Tag             =   "36"
      ToolTipText     =   "35000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   21
      Left            =   7830
      TabIndex        =   43
      Tag             =   "38"
      ToolTipText     =   "40000"
      Top             =   3615
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   23
      Left            =   4050
      TabIndex        =   48
      Tag             =   "27"
      ToolTipText     =   "15000"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   24
      Left            =   345
      TabIndex        =   49
      Tag             =   "14"
      ToolTipText     =   "10000"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   25
      Left            =   1080
      TabIndex        =   50
      Tag             =   "34"
      ToolTipText     =   "10000"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   26
      Left            =   1815
      TabIndex        =   51
      Tag             =   "4"
      ToolTipText     =   "10000"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   27
      Left            =   2550
      TabIndex        =   52
      Tag             =   "24"
      ToolTipText     =   "10000"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   0
      Left            =   330
      TabIndex        =   0
      ToolTipText     =   "5000"
      Top             =   1245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   2
      Left            =   1830
      TabIndex        =   4
      ToolTipText     =   "5000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   5
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   10
      Left            =   7830
      TabIndex        =   20
      ToolTipText     =   "10000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   18
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   8
      Left            =   6330
      TabIndex        =   16
      ToolTipText     =   "10000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   15
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   6
      Left            =   4830
      TabIndex        =   12
      ToolTipText     =   "10000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   12
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   4
      Left            =   3330
      TabIndex        =   8
      ToolTipText     =   "5000"
      Top             =   1245
      Visible         =   0   'False
      WhatsThisHelpID =   8
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   21
      Left            =   7845
      TabIndex        =   42
      ToolTipText     =   "20000"
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   38
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   19
      Left            =   6345
      TabIndex        =   38
      ToolTipText     =   "20000"
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   33
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   17
      Left            =   4845
      TabIndex        =   34
      ToolTipText     =   "20000"
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   30
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   15
      Left            =   3345
      TabIndex        =   30
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   26
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   13
      Left            =   1845
      TabIndex        =   26
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   23
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdSellHouse 
      Height          =   540
      Index           =   11
      Left            =   345
      TabIndex        =   22
      Top             =   3615
      Visible         =   0   'False
      WhatsThisHelpID =   20
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   16776960
      ButtonShape     =   3
      Caption         =   "Sell the Houses"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   0
      Left            =   330
      TabIndex        =   53
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   1
      Left            =   1080
      TabIndex        =   54
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   2
      Left            =   1830
      TabIndex        =   55
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   3
      Left            =   2580
      TabIndex        =   56
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   4
      Left            =   3330
      TabIndex        =   57
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   5
      Left            =   4080
      TabIndex        =   58
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   6
      Left            =   4830
      TabIndex        =   59
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   7
      Left            =   5580
      TabIndex        =   60
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   8
      Left            =   6315
      TabIndex        =   61
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   9
      Left            =   7065
      TabIndex        =   62
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   10
      Left            =   7815
      TabIndex        =   63
      Tag             =   "0"
      Top             =   1875
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   11
      Left            =   330
      TabIndex        =   64
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   12
      Left            =   1080
      TabIndex        =   65
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   13
      Left            =   1830
      TabIndex        =   66
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   14
      Left            =   2580
      TabIndex        =   67
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   15
      Left            =   3330
      TabIndex        =   68
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   16
      Left            =   4080
      TabIndex        =   69
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   17
      Left            =   4830
      TabIndex        =   70
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   18
      Left            =   5580
      TabIndex        =   71
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   19
      Left            =   6315
      TabIndex        =   72
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   20
      Left            =   7065
      TabIndex        =   73
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   21
      Left            =   7815
      TabIndex        =   74
      Tag             =   "0"
      Top             =   4245
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdPayMortgage 
      Height          =   540
      Index           =   22
      Left            =   3330
      TabIndex        =   75
      Tag             =   "0"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   65280
      ButtonShape     =   3
      Caption         =   "Pay Back Mortgage"
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
   Begin Project1.dcButton cmdSellDead 
      Height          =   540
      Index           =   22
      Left            =   3330
      TabIndex        =   47
      Tag             =   "11"
      ToolTipText     =   "15000"
      Top             =   6375
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      BackColor       =   255
      ButtonShape     =   3
      Caption         =   "Morgage Property"
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
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      X1              =   22
      X2              =   571
      Y1              =   328
      Y2              =   328
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      X1              =   22
      X2              =   571
      Y1              =   171
      Y2              =   171
   End
   Begin VB.Image imgAirport 
      Height          =   1155
      Index           =   3
      Left            =   2565
      Picture         =   "frmBroke.frx":154E
      Top             =   5070
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgAirport 
      Height          =   1155
      Index           =   2
      Left            =   1830
      Picture         =   "frmBroke.frx":1CF6
      Top             =   5070
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgAirport 
      Height          =   1155
      Index           =   1
      Left            =   1095
      Picture         =   "frmBroke.frx":24F4
      Top             =   5070
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgAirport 
      Height          =   1155
      Index           =   0
      Left            =   360
      Picture         =   "frmBroke.frx":2C99
      Top             =   5070
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgUtillitie 
      Height          =   1155
      Index           =   1
      Left            =   4080
      Picture         =   "frmBroke.frx":344E
      Stretch         =   -1  'True
      Top             =   5070
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgUtillitie 
      Height          =   1185
      Index           =   0
      Left            =   3345
      Picture         =   "frmBroke.frx":3BA6
      Stretch         =   -1  'True
      Top             =   5055
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank:"
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
      Height          =   360
      Left            =   7800
      TabIndex        =   45
      Top             =   5895
      Width           =   795
   End
   Begin VB.Label lblBank 
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
      Height          =   345
      Left            =   7815
      TabIndex        =   44
      Top             =   6300
      Width           =   3060
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   21
      Left            =   7980
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   20
      Left            =   7230
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   19
      Left            =   6480
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   18
      Left            =   5730
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   17
      Left            =   4980
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   16
      Left            =   4230
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   15
      Left            =   3480
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   14
      Left            =   2730
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   13
      Left            =   1980
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   12
      Left            =   1230
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   11
      Left            =   480
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   10
      Left            =   7980
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   9
      Left            =   7230
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   8
      Left            =   6480
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   7
      Left            =   5730
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   6
      Left            =   4980
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   5
      Left            =   4230
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   4
      Left            =   3480
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   3
      Left            =   2730
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   2
      Left            =   1980
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   1
      Left            =   1230
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHouse 
      Height          =   480
      Index           =   0
      Left            =   480
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   21
      Left            =   7830
      Picture         =   "frmBroke.frx":4389
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   20
      Left            =   7080
      Picture         =   "frmBroke.frx":450D
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   19
      Left            =   6330
      Picture         =   "frmBroke.frx":4691
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   18
      Left            =   5580
      Picture         =   "frmBroke.frx":486C
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   17
      Left            =   4830
      Picture         =   "frmBroke.frx":4A47
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   16
      Left            =   4080
      Picture         =   "frmBroke.frx":4C22
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   15
      Left            =   3330
      Picture         =   "frmBroke.frx":4E5D
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   14
      Left            =   2580
      Picture         =   "frmBroke.frx":5098
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   13
      Left            =   1830
      Picture         =   "frmBroke.frx":52D3
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   12
      Left            =   1080
      Picture         =   "frmBroke.frx":5512
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   11
      Left            =   330
      Picture         =   "frmBroke.frx":5751
      Stretch         =   -1  'True
      Top             =   2700
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   10
      Left            =   7830
      Picture         =   "frmBroke.frx":5990
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   9
      Left            =   7080
      Picture         =   "frmBroke.frx":5BD5
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   8
      Left            =   6330
      Picture         =   "frmBroke.frx":5E1A
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   7
      Left            =   5580
      Picture         =   "frmBroke.frx":605F
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   6
      Left            =   4830
      Picture         =   "frmBroke.frx":6FD1
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   5
      Left            =   4080
      Picture         =   "frmBroke.frx":7F43
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   4
      Left            =   3330
      Picture         =   "frmBroke.frx":8EB5
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   3
      Left            =   2580
      Picture         =   "frmBroke.frx":9101
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   2
      Left            =   1830
      Picture         =   "frmBroke.frx":934D
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   1
      Left            =   1080
      Picture         =   "frmBroke.frx":9599
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgProp 
      Height          =   300
      Index           =   0
      Left            =   330
      Picture         =   "frmBroke.frx":9841
      Stretch         =   -1  'True
      Top             =   315
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "frmBroke"
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

    Dim P2Counter1 As Integer
    Dim P2Counter2 As Integer

Public Function InitialSettings(Player As Integer, Bank As Double)
    Dim x As Integer
    Dim TmpFileName As Integer
    
    On Error Resume Next
    
    Call Form_Load
    
    frmBroke.Visible = True
    
    mPlayerBank = Bank
    mPlayer = Player
    
    For x = 0 To 27
        If frmBoard.imgDead((cmdSellDead(x).Tag)).Tag = Player Then
            If x <= 21 Then                 '''''''''''''
                imgProp(x).Visible = True   '''''''''''''
            End If                          '''''''''''''
            
            If frmBoard.imgHomes((cmdSellDead(x).Tag)).Tag = 0 Then
                TmpFileName = 0
                If x <= 21 Then
                    imgHouse(x).Tag = TmpFileName
                    imgHouse(x).Visible = False
                    cmdSellHouse(x).Visible = False
                End If
                If cmdPayMortgage(x).Visible = False Then
                    cmdSellDead(x).Visible = True
                Else
                    cmdSellDead(x).Visible = False
                End If
            Else
                TmpFileName = frmBoard.imgHomes((cmdSellDead(x).Tag)).Tag
                imgHouse(x).Tag = TmpFileName
                imgHouse(x).Picture = LoadPicture(Ap & "HOMESUP\" & TmpFileName & ".ico")
                imgHouse(x).Visible = True
                cmdSellHouse(x).Visible = True
                cmdSellDead(x).Visible = False
            End If
                        
            If frmBoard.imgDead((cmdSellDead(x).Tag)).WhatsThisHelpID = 9 Then
                cmdPayMortgage(x).Visible = True
                cmdSellDead(x).Visible = False
            End If
        End If
    Next x
                
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'AIRPORTS
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If frmBoard.imgDead(4).Tag = Player Then
        imgAirport(2).Visible = True
        cmdSellDead(26).Visible = True
    End If
    
    If frmBoard.imgDead(14).Tag = Player Then
        imgAirport(0).Visible = True
        cmdSellDead(24).Visible = True
    End If
    
    If frmBoard.imgDead(24).Tag = Player Then
        imgAirport(3).Visible = True
        cmdSellDead(27).Visible = True
    End If

    If frmBoard.imgDead(34).Tag = Player Then
        imgAirport(1).Visible = True
        cmdSellDead(25).Visible = True
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'UTILLITIES
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If frmBoard.imgDead(11).Tag = Player Then
        imgUtillitie(0).Visible = True
        cmdSellDead(22).Visible = True
    End If
    
    If frmBoard.imgDead(27).Tag = Player Then
        imgUtillitie(1).Visible = True
        cmdSellDead(23).Visible = True
    End If
    
    For x = 0 To 27
        If cmdPayMortgage(x).Visible = True Then
            cmdSellDead(x).Visible = False
        End If
    Next x

    lblBank = FormatCurrency(mPlayerBank, 2)
    
    If Player = 2 Then
        Call Player2Continue
    End If
    
End Function

Private Sub cmdDone_Click()
    Dim Msg As String
    
    If mPlayerBank < 0 Then
        Msg = MsgBox("You are broke. Either sell more Deads or Houses. If you click Cancel, the other player will win.", vbOKCancel, "You are broke!")
        frmBroke.Visible = False
    Else
        Call frmBoard.ResetAfterMortgage(mPlayer, mPlayerBank)
        Call Form_Load
        Call ResetForm
        frmBroke.Visible = False
    End If
    
    If Msg = "vbCancel" Then
        Call ResetForm
        frmBroke.Visible = False
        Call frmBoard.GameFinish
    Else
        Call frmBoard.ResetAfterMortgage(mPlayer, mPlayerBank)
        Call Form_Load
        Call ResetForm
        frmBroke.Visible = False
    End If
    
    
End Sub

Private Sub ResetForm()
    Dim x As Integer
    
    Call Form_Load
    
    For x = 0 To 27
        
        If x <= 3 Then
            imgAirport(x).Visible = False
        End If
        
        If x <= 1 Then
            imgUtillitie(x).Visible = False
        End If
        
        If x <= 21 Then
            imgHouse(x).Visible = False
            imgHouse(x).Picture = LoadPicture(Ap & "HOMESUP\6.ico")
            cmdSellHouse(x).Visible = False
        End If
        
        cmdSellDead(x).Visible = False
        cmdSellDead(x).Visible = False
        cmdPayMortgage(x).Visible = False
    Next x
    
End Sub

Private Sub cmdPayMortgage_Click(Index As Integer)

    mPlayerBank = mPlayerBank - (Int(cmdPayMortgage(Index).ToolTipText))
    lblBank = FormatCurrency(mPlayerBank, 2)
    frmBoard.imgDead((cmdSellDead(Index).Tag)).WhatsThisHelpID = 0
    cmdSellDead(Index).Visible = True
    cmdPayMortgage(Index).Visible = False
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''SET RENT PROPERTY FOR UTILITIES'''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If cmdSellDead(Index).Tag = 11 Then
        If frmBoard.imgDead(11).Tag = frmBoard.imgDead(27).Tag Then
            frmBoard.imgDead((cmdSellDead(Index).Tag)).WhatsThisHelpID = 8
            Exit Sub
        Else
            frmBoard.imgDead((cmdSellDead(Index).Tag)).WhatsThisHelpID = 7
            Exit Sub
        End If
        Exit Sub
    End If
    
    If cmdSellDead(Index).Tag = 27 Then
        If frmBoard.imgDead(11).Tag = frmBoard.imgDead(27).Tag Then
            frmBoard.imgDead((cmdSellDead(Index).Tag)).WhatsThisHelpID = 8
            Exit Sub
        Else
            frmBoard.imgDead((cmdSellDead(Index).Tag)).WhatsThisHelpID = 7
            Exit Sub
        End If
        Exit Sub
    End If
    
    ''''''''''''SET RENT PROPERTY END

End Sub

Private Sub cmdSellDead_Click(Index As Integer)
    
    mPlayerBank = mPlayerBank + (Int(cmdSellDead(Index).ToolTipText) / Int(2))
    lblBank = FormatCurrency(mPlayerBank, 2)
    frmBoard.imgDead(cmdSellDead(Index).Tag).WhatsThisHelpID = 9
    cmdSellDead(Index).Visible = False
    cmdPayMortgage(Index).Visible = True
    
End Sub

Private Sub cmdSellHouse_Click(Index As Integer)

    mPlayerBank = mPlayerBank + (Int(cmdSellHouse(Index).ToolTipText) * Int(imgHouse(Index).Tag))
    lblBank = FormatCurrency(mPlayerBank, 2)
    frmBoard.imgHomes(cmdSellHouse(Index).WhatsThisHelpID).Picture = LoadPicture(Ap & "HOMESUP\6.ico")
    frmBoard.imgHomes(cmdSellHouse(Index).WhatsThisHelpID).Tag = 0
    frmBoard.imgHomes(cmdSellHouse(Index).WhatsThisHelpID).WhatsThisHelpID = 0
    imgHouse(Index).Picture = LoadPicture(Ap & "HOMESUP\6.ico")
    frmBoard.imgDead(cmdSellHouse(Index).WhatsThisHelpID).WhatsThisHelpID = 0
    cmdSellDead(Index).Visible = True
    cmdSellHouse(Index).Visible = False
    
End Sub

Private Sub Form_Load()
    Dim x As Integer
    
    If Right(App.Path, 1) = "\" Then
        Ap = App.Path
    Else
        Ap = App.Path & "\"
    End If
    
    For x = 0 To 27
        If x <= 21 Then
            imgProp(x).Visible = False
            imgHouse(x).Visible = False
            cmdSellHouse(x).Visible = False
        End If
        cmdPayMortgage(x).ToolTipText = Int((cmdSellDead(x).ToolTipText / 2) + ((cmdSellDead(x).ToolTipText / 2) * 0.1))
        cmdSellDead(x).Visible = False
        cmdPayMortgage(x).Visible = False
    Next x
    
End Sub

Private Sub Player2Continue()
    Dim x As Integer, Counter As Integer
    
    Counter = 0
    
    For x = 0 To 21
        If imgProp(x).Visible = False Then
            Counter = Counter + 1
        End If
    Next x
    
    For x = 0 To 3
        If imgAirport(x).Visible = False Then
            Counter = Counter + 1
        End If
    Next x

    For x = 0 To 1
        If imgUtillitie(x).Visible = False Then
            Counter = Counter + 1
        End If
    Next x
    
    If Counter = 28 Then
        x = MsgBox(frmBoard.lblPlayer2Name.Caption & " is broke. YOU WIN!", vbOKOnly, "YOU WIN!")
        Call frmBoard.GameFinish
        Exit Sub
    End If
    
    If mPlayerBank <= 0 Then
        tmrPlayer2Sell.Enabled = True
    End If
    
    If mPlayerBank > 0 Then
        tmrPlayer2Buy.Enabled = True
    End If
    
End Sub

Private Sub tmrPlayer2Buy_Timer()
    Dim x As Integer
    
    If P2Counter1 < 28 Then
        If cmdPayMortgage(P2Counter1).Visible = True And mPlayerBank > 20000 Then
            Call cmdPayMortgage_Click(P2Counter1)
            P2Counter1 = P2Counter1 + 1
        End If
    Else
        P2Counter1 = 0
        tmrPlayer2Sell.Enabled = False
        tmrPlayer2Sell2.Enabled = True
    End If

End Sub

Private Sub tmrPlayer2Sell_Timer()
    Dim x As Integer
    
    If P2Counter1 < 22 Then
        If cmdSellHouse(P2Counter1).Visible = True Then
            If mPlayerBank < 0 Then
                Call cmdSellHouse_Click(P2Counter1)
            End If
        End If
        P2Counter1 = P2Counter1 + 1
    Else
        P2Counter1 = 0
        tmrPlayer2Sell.Enabled = False
        tmrPlayer2Sell2.Enabled = True
    End If
    
End Sub

Private Sub tmrPlayer2Sell2_Timer()
    Dim x As Integer
    
    If P2Counter1 < 28 Then
        If cmdSellDead(P2Counter1).Visible = True Then
            If mPlayerBank < 0 Then
                Call cmdSellDead_Click(P2Counter1)
            End If
        End If
        P2Counter1 = P2Counter1 + 1
    Else
        P2Counter1 = 0
        If mPlayerBank <= 0 Then
            frmBroke.Visible = False
            x = MsgBox(frmBoard.lblPlayer2Name.Caption & " is broke. YOU WIN!", vbOKOnly, "YOU WIN!")
            tmrPlayer2Sell2.Enabled = False
            Call frmBoard.GameFinish
            Exit Sub
        End If
        Call cmdDone_Click
        tmrPlayer2Sell2.Enabled = False
    End If
    
End Sub

