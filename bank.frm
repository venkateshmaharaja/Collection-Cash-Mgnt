VERSION 5.00
Begin VB.Form form4 
   BackColor       =   &H00FFFF80&
   Caption         =   "bank"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton pluscbcoins 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10150
      TabIndex        =   117
      Top             =   7320
      Width           =   345
   End
   Begin VB.CommandButton pluschcoins 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   116
      Top             =   7320
      Width           =   345
   End
   Begin VB.CommandButton pluscb5 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10150
      TabIndex        =   115
      Top             =   6600
      Width           =   345
   End
   Begin VB.CommandButton plusch5 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   114
      Top             =   6600
      Width           =   345
   End
   Begin VB.CommandButton pluscb10 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10150
      TabIndex        =   113
      Top             =   5880
      Width           =   345
   End
   Begin VB.CommandButton plusch10 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   112
      Top             =   5880
      Width           =   345
   End
   Begin VB.CommandButton pluscb20 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10150
      TabIndex        =   111
      Top             =   5160
      Width           =   345
   End
   Begin VB.CommandButton plusch20 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   110
      Top             =   5160
      Width           =   345
   End
   Begin VB.CommandButton pluscb50 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10150
      TabIndex        =   109
      Top             =   4440
      Width           =   345
   End
   Begin VB.CommandButton plusch50 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   108
      Top             =   4440
      Width           =   345
   End
   Begin VB.CommandButton pluscb100 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10150
      TabIndex        =   107
      Top             =   3720
      Width           =   345
   End
   Begin VB.CommandButton plusch100 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   106
      Top             =   3720
      Width           =   345
   End
   Begin VB.CommandButton pluscb500 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10150
      TabIndex        =   105
      Top             =   3000
      Width           =   345
   End
   Begin VB.CommandButton plusch500 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   104
      Top             =   3000
      Width           =   345
   End
   Begin VB.CommandButton pluscb1000 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10150
      TabIndex        =   103
      Top             =   2280
      Width           =   345
   End
   Begin VB.CommandButton plusch1000 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   102
      Top             =   2280
      Width           =   345
   End
   Begin VB.CommandButton plusch2000 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10500
      TabIndex        =   101
      Top             =   1560
      Width           =   345
   End
   Begin VB.CommandButton pluscb2000 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10150
      TabIndex        =   100
      Top             =   1560
      Width           =   345
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "PRINT PAGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   9120
      Width           =   2655
   End
   Begin VB.CommandButton cmdbackh 
      Caption         =   "< BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   7
      Top             =   8640
      Width           =   3615
   End
   Begin VB.CommandButton cmdbackc 
      Caption         =   "< BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   8640
      Width           =   3615
   End
   Begin VB.TextBox totchand 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   14040
      TabIndex        =   3
      Top             =   8040
      Width           =   2295
   End
   Begin VB.TextBox totcbank 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7920
      TabIndex        =   2
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   12120
      TabIndex        =   99
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   12120
      TabIndex        =   98
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   12120
      TabIndex        =   97
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   12120
      TabIndex        =   96
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   12120
      TabIndex        =   95
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   12120
      TabIndex        =   94
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   12120
      TabIndex        =   93
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   12120
      TabIndex        =   92
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   12120
      TabIndex        =   91
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label labchcoins 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   90
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label labch5 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   89
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label labch10 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   88
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label ch1000 
      BackColor       =   &H8000000E&
      Caption         =   "1000    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   87
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label ch500 
      BackColor       =   &H8000000E&
      Caption         =   "500      "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   86
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label ch100 
      BackColor       =   &H8000000E&
      Caption         =   "100      *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   85
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label ch50 
      BackColor       =   &H8000000E&
      Caption         =   "50        *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   84
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label ch20 
      BackColor       =   &H8000000E&
      Caption         =   "20        *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   83
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label ch10 
      BackColor       =   &H8000000E&
      Caption         =   "10        *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   82
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label ch5 
      BackColor       =   &H8000000E&
      Caption         =   "5          *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   81
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label chcoins 
      BackColor       =   &H8000000E&
      Caption         =   "coins   *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   80
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label ch2000 
      BackColor       =   &H8000000E&
      Caption         =   "2000    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   79
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label txtch2000 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   78
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label txtch1000 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   77
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label txtch500 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   76
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label txtch100 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   75
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label txtch50 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   74
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label txtch20 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   73
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label txtch10 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   72
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label txtch5 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   71
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label txtchcoins 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   70
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   13560
      TabIndex        =   69
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   13560
      TabIndex        =   68
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   13560
      TabIndex        =   67
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   13560
      TabIndex        =   66
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   13560
      TabIndex        =   65
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   13560
      TabIndex        =   64
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   13560
      TabIndex        =   63
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   13560
      TabIndex        =   62
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   13560
      TabIndex        =   61
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label labch2000 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   60
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label labch1000 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   59
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label labch500 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   58
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label labch100 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   57
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label labch50 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   56
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label labch20 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14160
      TabIndex        =   55
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Labcbcoins 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   54
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Labcb5 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   53
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Labcb10 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   52
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Labcb20 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   51
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Labcb50 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   50
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Labcb100 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   49
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Labcb500 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   48
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Labcb1000 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   47
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Labcb2000 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   46
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   7560
      TabIndex        =   45
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   7560
      TabIndex        =   44
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   7560
      TabIndex        =   43
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   7560
      TabIndex        =   42
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   7560
      TabIndex        =   41
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   7560
      TabIndex        =   40
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   7560
      TabIndex        =   39
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   7560
      TabIndex        =   38
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   7560
      TabIndex        =   37
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label txtcbcoins 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   36
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label txtcb5 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   35
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label txtcb10 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   34
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label txtcb20 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   33
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label txtcb50 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   32
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label txtcb100 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   31
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label txtcb500 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   30
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label txtcb1000 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   29
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label txtcb2000 
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   28
      Top             =   1560
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
      Height          =   615
      Index           =   8
      Left            =   6120
      TabIndex        =   27
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
      Height          =   615
      Index           =   7
      Left            =   6120
      TabIndex        =   26
      Top             =   6600
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
      Height          =   615
      Index           =   6
      Left            =   6120
      TabIndex        =   25
      Top             =   5880
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
      Height          =   615
      Index           =   5
      Left            =   6120
      TabIndex        =   24
      Top             =   5160
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
      Height          =   615
      Index           =   4
      Left            =   6120
      TabIndex        =   23
      Top             =   4440
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
      Height          =   615
      Index           =   3
      Left            =   6120
      TabIndex        =   22
      Top             =   3720
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
      Height          =   615
      Index           =   2
      Left            =   6120
      TabIndex        =   21
      Top             =   3000
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
      Height          =   615
      Index           =   1
      Left            =   6120
      TabIndex        =   20
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label cb2000 
      BackColor       =   &H8000000E&
      Caption         =   "2000    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
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
      Height          =   615
      Index           =   0
      Left            =   6120
      TabIndex        =   18
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label cbcoin 
      BackColor       =   &H8000000E&
      Caption         =   "coins   *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   17
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label cb5 
      BackColor       =   &H8000000E&
      Caption         =   "5          *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   16
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label cb10 
      BackColor       =   &H8000000E&
      Caption         =   "10        *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   15
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label cb20 
      BackColor       =   &H8000000E&
      Caption         =   "20        *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   14
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label cb50 
      BackColor       =   &H8000000E&
      Caption         =   "50        *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label cb100 
      BackColor       =   &H8000000E&
      Caption         =   "100      *"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label cb500 
      BackColor       =   &H8000000E&
      Caption         =   "500      "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label cb1000 
      BackColor       =   &H8000000E&
      Caption         =   "1000    "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "CASH  FLOW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   9240
      TabIndex        =   9
      Top             =   0
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   9015
      Left            =   4560
      Top             =   720
      Width           =   12015
   End
   Begin VB.Label Label2 
      Caption         =   "  TOTAL  AMOUNT  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   5
      Top             =   8040
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "  TOTAL  AMOUNT  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   8040
      Width           =   3135
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
      Left            =   10800
      TabIndex        =   1
      Top             =   960
      Width           =   5415
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
      TabIndex        =   0
      Top             =   960
      Width           =   5415
   End
End
Attribute VB_Name = "form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset


Dim D1 As Date
Dim d2 As Date
Dim k As Integer
Dim j As Integer




Private Sub cmdbackc_Click()
Form3.Show

End Sub

Private Sub cmdbackh_Click()
Form3.Show

End Sub

Private Sub cmdprint_Click()

Form5.Text1 = Form3.Calendar1.Value

Form5.txtcb2000 = txtcb2000
Form5.txtcb1000 = txtcb1000
Form5.txtcb500 = txtcb500
Form5.txtcb100 = txtcb100
Form5.txtcb50 = txtcb50
Form5.txtcb20 = txtcb20
Form5.txtcb10 = txtcb10
Form5.txtcbcoins = txtcbcoins

Form5.Labcb2000 = Labcb2000
Form5.Labcb1000 = Labcb1000
Form5.Labcb500 = Labcb500
Form5.Labcb100 = Labcb100
Form5.Labcb50 = Labcb50
Form5.Labcb20 = Labcb20
Form5.Labcb10 = Labcb10
Form5.Labcb5 = Labcb5
Form5.Labcbcoins = Labcbcoins

Form5.totcbank = totcbank

Form5.txtch2000 = txtch2000
Form5.txtch1000 = txtch1000
Form5.txtch500 = txtch500
Form5.txtch100 = txtch100
Form5.txtch50 = txtch50
Form5.txtch20 = txtch20
Form5.txtch10 = txtch10
Form5.txtchcoins = txtchcoins

Form5.Labch2000 = Labch2000
Form5.Labch1000 = Labch1000
Form5.Labch500 = Labch500
Form5.Labch100 = Labch100
Form5.Labch50 = Labch50
Form5.Labch20 = Labch20
Form5.Labch10 = Labch10
Form5.Labch5 = Labch5
Form5.Labchcoins = Labchcoins

Form5.totchand = totchand
Form5.txttoto = Form3.txttoto





m = 0
Form5.flex3.Cols = Form3.flex2.Cols
Form5.flex3.Rows = Form3.flex2.Rows
 
 For m = 0 To Form3.flex2.Rows - 1
 
 
 
Form5.flex3.TextMatrix(m, 0) = Form3.flex2.TextMatrix(o, 0)
Form5.flex3.TextMatrix(m, 1) = Form3.flex2.TextMatrix(o, 1)
Form5.flex3.TextMatrix(m, 2) = Form3.flex2.TextMatrix(o, 2)
Form5.flex3.TextMatrix(m, 3) = Form3.flex2.TextMatrix(o, 3)
Form5.flex3.TextMatrix(m, 4) = Form3.flex2.TextMatrix(o, 4)
Form5.flex3.TextMatrix(m, 5) = Form3.flex2.TextMatrix(o, 5)
Form5.flex3.TextMatrix(m, 6) = Form3.flex2.TextMatrix(o, 6)
Form5.flex3.TextMatrix(m, 7) = Form3.flex2.TextMatrix(o, 7)
Form5.flex3.TextMatrix(m, 8) = Form3.flex2.TextMatrix(o, 8)
Form5.flex3.TextMatrix(m, 9) = Form3.flex2.TextMatrix(o, 9)
Form5.flex3.TextMatrix(m, 10) = Form3.flex2.TextMatrix(o, 10)
Form5.flex3.TextMatrix(m, 11) = Form3.flex2.TextMatrix(o, 11)
Form5.flex3.TextMatrix(m, 12) = Form3.flex2.TextMatrix(o, 12)
o = o + 1
Next




Form5.Txtto2000 = Val(Form3.Txtto2000)
Form5.txtto1000 = Form3.txtto1000.Text
Form5.txtto500 = Form3.txtto500.Text
Form5.txtto100 = Form3.txtto100.Text
Form5.txtto50 = Form3.txtto50.Text
Form5.txtto20 = Form3.txtto20.Text
Form5.txtto10 = Form3.txtto10.Text
Form5.txtto5 = Form3.txtto5.Text
Form5.txttocoins = Form3.txttocoins.Text



End Sub


Private Sub display()
k = 0

Form5.flex3.Cols = 13
Form5.flex3.Rows = 1
Form5.flex3.Row = k
Form5.flex3.Col = 0

Form5.flex3.Text = "DATE"
Form5.flex3.Col = 1
Form5.flex3.Text = "ENO"
Form5.flex3.Col = 2
Form5.flex3.Text = "NAME"
Form5.flex3.Col = 3
Form5.flex3.Text = "2000"
Form5.flex3.Col = 4
Form5.flex3.Text = "1000"
Form5.flex3.Col = 5
Form5.flex3.Text = "500"
Form5.flex3.Col = 6
Form5.flex3.Text = "100"
Form5.flex3.Col = 7
Form5.flex3.Text = "50"
Form5.flex3.Col = 8
Form5.flex3.Text = "20"
Form5.flex3.Col = 9
Form5.flex3.Text = "10"
Form5.flex3.Col = 10
Form5.flex3.Text = "5"
Form5.flex3.Col = 11
Form5.flex3.Text = "COINS"
Form5.flex3.Col = 12
Form5.flex3.Text = "TOTAL"

j = 0
k = k + 1

While rs2.EOF <> True

flex3.Rows = flex3.Rows + 1
flex3.Row = flex3.Row + 1
flex3.Col = j

flex3.Text = rs2.Fields(0)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(1)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(2)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(3)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(4)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(5)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(6)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(7)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(8)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(9)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(10)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(11)
flex3.Col = flex3.Col + 1
flex3.Text = rs2.Fields(12)
k = k + 1
rs2.MoveNext

Wend
rs2.Close
End Sub




Private Sub pluscb10_Click()
If ((txtch10 = "") Or (txtch10 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb10 = Val(txtcb10) + 1
txtch10 = Val(txtch10) - 1

form4.Labcb10 = Val(form4.txtcb10 * 10)
form4.Labch10 = Val(form4.txtch10 * 10)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub
Private Sub pluscb100_Click()
If ((txtch100 = "") Or (txtch100 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb100 = Val(txtcb100) + 1
txtch100 = Val(txtch100) - 1

form4.Labcb100 = Val(form4.txtcb100 * 100)
form4.Labch100 = Val(form4.txtch100 * 100)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub pluscb1000_Click()
If ((txtch1000 = "") Or (txtch1000 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb1000 = Val(txtcb1000) + 1
txtch1000 = Val(txtch1000) - 1

form4.Labcb1000 = Val(form4.txtcb1000 * 1000)
form4.Labch1000 = Val(form4.txtch1000 * 1000)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub pluscb20_Click()
If ((txtch20 = "") Or (txtch20 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb20 = Val(txtcb20) + 1
txtch20 = Val(txtch20) - 1

form4.Labcb20 = Val(form4.txtcb20 * 20)
form4.Labch20 = Val(form4.txtch20 * 20)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub pluscb2000_Click()
If ((txtch2000 = "") Or (txtch2000 = 0)) Then
MsgBox ("NO VALUES?")
Else
txtcb2000 = Val(txtcb2000) + 1
txtch2000 = Val(txtch2000) - 1

form4.Labcb2000 = Val(form4.txtcb2000 * 2000)
form4.Labch2000 = Val(form4.txtch2000 * 2000)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub pluscb5_Click()
If ((txtch5 = "") Or (txtch5 = 0)) Then
MsgBox ("NO VALUES?")
Else
txtcb5 = Val(txtcb5) + 1
txtch5 = Val(txtch5) - 1

form4.Labcb5 = Val(form4.txtcb5 * 5)
form4.Labch5 = Val(form4.txtch5 * 5)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub pluscb50_Click()
If ((txtch50 = "") Or (txtch50 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb50 = Val(txtcb50) + 1
txtch50 = Val(txtch50) - 1

form4.Labcb50 = Val(form4.txtcb50 * 50)
form4.Labch50 = Val(form4.txtch50 * 50)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub pluscb500_Click()
If ((txtch500 = "") Or (txtch500 = 0)) Then
MsgBox ("NO VALUES?")
Else
txtcb500 = Val(txtcb500) + 1
txtch500 = Val(txtch500) - 1

form4.Labcb500 = Val(form4.txtcb500 * 500)
form4.Labch500 = Val(form4.txtch500 * 500)


form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub pluscbcoins_Click()
If ((txtchcoins = "") Or (txtchcoins = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcbcoins = Val(txtcbcoins) + 1
txtchcoins = Val(txtchcoins) - 1

form4.Labcbcoins = Val(form4.txtcbcoins * 1)
form4.Labchcoins = Val(form4.txtchcoins * 1)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub plusch10_Click()
If ((txtcb10 = "") Or (txtcb10 = 0)) Then
MsgBox ("NO VALUES?")
Else
txtcb10 = Val(txtcb10) - 1
txtch10 = Val(txtch10) + 1

form4.Labcb10 = Val(form4.txtcb10 * 10)
form4.Labch10 = Val(form4.txtch10 * 10)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub plusch100_Click()
If ((txtcb100 = "") Or (txtcb100 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb100 = Val(txtcb100) - 1
txtch100 = Val(txtch100) + 1

form4.Labcb100 = Val(form4.txtcb100 * 100)
form4.Labch100 = Val(form4.txtch100 * 100)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub plusch1000_Click()
If ((txtcb1000 = "") Or (txtcb1000 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb1000 = Val(txtcb1000) - 1
txtch1000 = Val(txtch1000) + 1

form4.Labcb1000 = Val(form4.txtcb1000 * 1000)
form4.Labch1000 = Val(form4.txtch1000 * 1000)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub plusch20_Click()
If ((txtcb20 = "") Or (txtcb20 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb20 = Val(txtcb20) - 1
txtch20 = Val(txtch20) + 1

form4.Labcb20 = Val(form4.txtcb20 * 20)
form4.Labch20 = Val(form4.txtch20 * 20)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub plusch2000_Click()
If ((txtcb2000 = "") Or (txtcb2000 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb2000 = Val(txtcb2000) - 1
txtch2000 = Val(txtch2000) + 1

form4.Labcb2000 = Val(form4.txtcb2000 * 2000)
form4.Labch2000 = Val(form4.txtch2000 * 2000)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub plusch5_Click()
If ((txtcb5 = "") Or (txtcb5 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb5 = Val(txtcb5) - 1
txtch5 = Val(txtch5) + 1

form4.Labcb5 = Val(form4.txtcb5 * 5)
form4.Labch5 = Val(form4.txtch5 * 5)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub plusch50_Click()
If ((txtcb50 = "") Or (txtcb50 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb50 = Val(txtcb50) - 1
txtch50 = Val(txtch50) + 1

form4.Labcb50 = Val(form4.txtcb50 * 50)
form4.Labch50 = Val(form4.txtch50 * 50)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

Private Sub plusch500_Click()
If ((txtcb500 = "") Or (txtcb500 = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcb500 = Val(txtcb500) - 1
txtch500 = Val(txtch500) + 1

form4.Labcb500 = Val(form4.txtcb500 * 500)
form4.Labch500 = Val(form4.txtch500 * 500)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub
Private Sub pluschcoins_Click()
If ((txtcbcoins = "") Or (txtcbcoins = 0)) Then
MsgBox ("NO VALUES?")
Else

txtcbcoins = Val(txtcbcoins) - 1
txtchcoins = Val(txtchcoins) + 1

form4.Labcbcoins = Val(form4.txtcbcoins * 1)
form4.Labchcoins = Val(form4.txtchcoins * 1)

form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins) + Val(form4.Labch2000) + Val(form4.Labch1000) + Val(form4.Labch500) + Val(form4.Labch100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100) + Val(form4.Labcb50) + Val(form4.Labcb20) + Val(form4.Labcb10) + Val(form4.Labcb5) + Val(form4.Labcbcoins)
End If
End Sub

