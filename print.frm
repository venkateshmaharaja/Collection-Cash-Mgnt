VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form5"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   840
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   12120
      Width           =   1200
   End
   Begin VB.TextBox totchand 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15120
      TabIndex        =   111
      Text            =   "1000000"
      Top             =   9960
      Width           =   2040
   End
   Begin VB.TextBox totcbank 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7440
      TabIndex        =   20
      Text            =   "1000000"
      Top             =   9960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15240
      TabIndex        =   19
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txttoto 
      Height          =   405
      Left            =   16215
      TabIndex        =   10
      Top             =   5640
      Width           =   1140
   End
   Begin VB.TextBox txttocoins 
      Height          =   405
      Left            =   15120
      TabIndex        =   9
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtto5 
      Height          =   405
      Left            =   14160
      TabIndex        =   8
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtto10 
      Height          =   405
      Left            =   13200
      TabIndex        =   7
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtto50 
      Height          =   405
      Left            =   11280
      TabIndex        =   6
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtto20 
      Height          =   405
      Left            =   12240
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtto100 
      Height          =   405
      Left            =   10320
      TabIndex        =   4
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtto500 
      Height          =   405
      Left            =   9360
      TabIndex        =   3
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtto1000 
      Height          =   405
      Left            =   8400
      TabIndex        =   2
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox Txtto2000 
      Height          =   405
      Left            =   7320
      TabIndex        =   1
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Labchcoins 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   110
      Top             =   9480
      Width           =   1935
   End
   Begin VB.Label Labch5 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   109
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Label Labch10 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   108
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label Labch20 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   107
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Label Labch50 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   106
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Labch100 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   105
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Labch500 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   104
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Labch1000 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   103
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Labch2000 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   102
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   14760
      TabIndex        =   101
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   14760
      TabIndex        =   100
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label txtchcoins 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   99
      Top             =   9480
      Width           =   735
   End
   Begin VB.Label txtch5 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   98
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label txtch10 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   97
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label txtch20 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   96
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label txtch50 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   95
      Top             =   8040
      Width           =   735
   End
   Begin VB.Label txtch100 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   94
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label txtch500 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   93
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label txtch1000 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   92
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label txtch2000 
      BackColor       =   &H8000000E&
      Caption         =   " 1000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   91
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   13680
      TabIndex        =   90
      Top             =   9480
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   13680
      TabIndex        =   89
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   13680
      TabIndex        =   88
      Top             =   8760
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   13680
      TabIndex        =   87
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   13680
      TabIndex        =   86
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   13680
      TabIndex        =   85
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   13680
      TabIndex        =   84
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   13680
      TabIndex        =   83
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      Caption         =   "2000    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   82
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   13680
      TabIndex        =   81
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "coins   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   80
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "5          "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   79
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "10        "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   78
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "20        "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   77
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "50        "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   76
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "100      "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   75
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "500      "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   74
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "1000    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   73
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   14760
      TabIndex        =   72
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   14760
      TabIndex        =   71
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   14760
      TabIndex        =   70
      Top             =   8040
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   14760
      TabIndex        =   69
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   14760
      TabIndex        =   68
      Top             =   8760
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   14760
      TabIndex        =   67
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   14760
      TabIndex        =   66
      Top             =   9480
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7080
      TabIndex        =   65
      Top             =   9480
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   7080
      TabIndex        =   64
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   7080
      TabIndex        =   63
      Top             =   8760
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   7080
      TabIndex        =   62
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7080
      TabIndex        =   61
      Top             =   8040
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7080
      TabIndex        =   60
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   59
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label cb1000 
      BackColor       =   &H8000000E&
      Caption         =   "1000    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   58
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label cb500 
      BackColor       =   &H8000000E&
      Caption         =   "500      "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   57
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label cb100 
      BackColor       =   &H8000000E&
      Caption         =   "100      "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   56
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label cb50 
      BackColor       =   &H8000000E&
      Caption         =   "50        "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   55
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label cb20 
      BackColor       =   &H8000000E&
      Caption         =   "20        "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   54
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label cb10 
      BackColor       =   &H8000000E&
      Caption         =   "10        "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   53
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Label cb5 
      BackColor       =   &H8000000E&
      Caption         =   "5          "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   52
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Label cbcoin 
      BackColor       =   &H8000000E&
      Caption         =   "coins   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   51
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6000
      TabIndex        =   50
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label cb2000 
      BackColor       =   &H8000000E&
      Caption         =   "2000    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   49
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   48
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   6000
      TabIndex        =   47
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   46
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   6000
      TabIndex        =   45
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6000
      TabIndex        =   44
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   6000
      TabIndex        =   43
      Top             =   8760
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6000
      TabIndex        =   42
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   6000
      TabIndex        =   41
      Top             =   9480
      Width           =   375
   End
   Begin VB.Label txtcb2000 
      BackColor       =   &H8000000E&
      Caption         =   " 1000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   40
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label txtcb1000 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   39
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label txtcb500 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   38
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label txtcb100 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   37
      Top             =   7680
      Width           =   735
   End
   Begin VB.Label txtcb50 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   36
      Top             =   8040
      Width           =   735
   End
   Begin VB.Label txtcb20 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   35
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label txtcb10 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   34
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label txtcb5 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   33
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label txtcbcoins 
      BackColor       =   &H8000000E&
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   32
      Top             =   9480
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   7080
      TabIndex        =   31
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "  ="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   30
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Labcb2000 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   29
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Labcb1000 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   28
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Labcb500 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   27
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Labcb100 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   26
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Labcb50 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   25
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Labcb20 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Label Labcb10 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   23
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label Labcb5 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   22
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Label Labcbcoins 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   21
      Top             =   9480
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "VSV SOFT TECHNOLOGY,TUTY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   18
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "DATE :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   14040
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "REPORT VIEW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   16
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Labcbank 
      Caption         =   "              CASH IN BANK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   6120
      Width           =   4935
   End
   Begin VB.Label labchand 
      Caption         =   "              CASH IN HAND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   14
      Top             =   6120
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   " TOTAL  AMT     :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4920
      TabIndex        =   13
      Top             =   9960
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "  TOTAL  AMT   :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12600
      TabIndex        =   12
      Top             =   9960
      Width           =   2655
   End
   Begin VB.Label Labamt 
      Caption         =   "                       TOTAL "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4440
      TabIndex        =   11
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   9975
      Left            =   4200
      Top             =   840
      Width           =   13335
   End
   Begin VB.Menu filemnu 
      Caption         =   "File"
      Begin VB.Menu MNUPRINT 
         Caption         =   "&PRINT"
      End
      Begin VB.Menu MNUCLOSE 
         Caption         =   "&CLOSE"
      End
      Begin VB.Menu MNUBACK 
         Caption         =   "&BACK"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Variant
Dim ts As Variant
Option Explicit
 
Const VK_MENU = 18
Const VK_SNAPSHOT = 44
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
 
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
 
Private Function SaveFormPic() As Picture
a = InputBox("enter the name")

Dim pic As StdPicture
 Set pic = Clipboard.GetData(vbCFBitmap)
  keybd_event VK_MENU, 0, 0, 0   'Send ALT key
  keybd_event VK_SNAPSHOT, 0, 0, 0   'Send PRINT SCREEN key
  DoEvents
  keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0   'Release PRINT SCREEN key
  keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0   'Release ALT key
  DoEvents
  Set SaveFormPic = Clipboard.GetData(vbCFBitmap)
 Clipboard.SetData pic, vbCFBitmap
End Function
 
Private Sub Command1_Click()
 End Sub
Private Sub Form_Load()

flex3.ColWidth(2) = 1450
flex3.ColWidth(1) = 547
flex3.ColWidth(0) = 1100


Form5.Show

End Sub

Private Sub MNUBACK_Click()
form4.Show
End Sub

Private Sub MNUCLOSE_Click()

Unload Form1
Unload Form2
Unload Form3
Unload form4
Unload Me

End
End Sub

Private Sub MNUPRINT_Click()
ts = Form3.Calendar1.Value
ts = Format(ts, "_dd_mm_yyyy")
 SavePicture SaveFormPic, "E:\MIN PROJECT\mout\" & a + ts & ".jpg" 'picture location
 Clipboard.Clear
'CommonDialog1.ShowPrinter
'Form5.PrintForm

End Sub

