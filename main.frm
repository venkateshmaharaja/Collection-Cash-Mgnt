VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlex2 
      Height          =   5535
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   8
      Cols            =   10
      RowHeightMin    =   787
      BorderStyle     =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlex1 
      Height          =   2955
      Left            =   0
      TabIndex        =   6
      Top             =   405
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5212
      _Version        =   393216
      Rows            =   7
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   428
      BorderStyle     =   0
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3375
      Left            =   6150
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      _Version        =   524288
      _ExtentX        =   7435
      _ExtentY        =   5953
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2017
      Month           =   1
      Day             =   9
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000E&
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1250
   End
   Begin VB.CommandButton cmdopen 
      BackColor       =   &H8000000E&
      Caption         =   "OPEN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1250
   End
   Begin VB.CommandButton cmdsearch 
      BackColor       =   &H8000000E&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1250
   End
   Begin VB.TextBox txtdate 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   3660
      TabIndex        =   1
      Text            =   "01/10/1996"
      Top             =   0
      Width           =   1300
   End
   Begin VB.CommandButton Cmdclose 
      BackColor       =   &H8000000E&
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calendar1_Click()
txtdate.Text = Format(Calendar1.Day & "/" & Calendar1.Month & "/" & Calendar1.Year, "DD/MM/YYYY")

End Sub

Private Sub Cmdclose_Click()
End
End Sub

Private Sub Form_Load()
'MSFlex1.Rows = 1
'MSFlex1.AddItem "NAME  " & "DESCRIPTION" & vbTab & "AMOUNT"
'With MSFlex1
 '   .TextMatrix(0, 0) = "Col # 0"
  ''  .ColWidth(0) = 800
    '.TextMatrix(0, 1) = "Col # 1"
    '.ColWidth(1) = 2000
    '.TextMatrix(0, 2) = "Col # 2"
    '.ColWidth(2) = 1800
'End With
MSFlex1.ColWidth(0) = 2032
MSFlex1.ColWidth(1) = 2900
MSFlex1.ColWidth(2) = 950
MSFlex1.TextMatrix(0, 0) = "NAME"
MSFlex1.TextMatrix(0, 1) = "DESCRIPTION"
MSFlex1.TextMatrix(0, 2) = "AMOUNT"
MSFlex2.ColWidth(0) = 2000
MSFlex2.TextMatrix(0, 0) = "NAME"
MSFlex2.TextMatrix(0, 1) = "2000"
MSFlex2.TextMatrix(0, 2) = "1000"
MSFlex2.TextMatrix(0, 3) = "500"
MSFlex2.TextMatrix(0, 4) = "100"
MSFlex2.TextMatrix(0, 5) = "50"
MSFlex2.TextMatrix(0, 6) = "20"
MSFlex2.TextMatrix(0, 7) = "10"
MSFlex2.TextMatrix(0, 8) = "5"
MSFlex2.TextMatrix(0, 9) = "coins"

End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub MSFlex1_Click()
InputBox (MSFlex1.TextMatrix(1, 0))



End Sub
