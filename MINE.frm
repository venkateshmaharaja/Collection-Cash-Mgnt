VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   ClientHeight    =   10950
   ClientLeft      =   5370
   ClientTop       =   2550
   ClientWidth     =   17265
   BeginProperty Font 
      Name            =   "@Meiryo"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   17265
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex1 
      Height          =   6735
      Left            =   6960
      TabIndex        =   41
      Top             =   960
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   12
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Meiryo"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   4575
      Left            =   3480
      TabIndex        =   40
      Top             =   1560
      Width           =   3135
      _Version        =   524288
      _ExtentX        =   5530
      _ExtentY        =   8070
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2017
      Month           =   10
      Day             =   30
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
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   285
      Left            =   6240
      TabIndex        =   39
      Top             =   7680
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   8520
   End
   Begin VB.CommandButton PREV 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3480
      TabIndex        =   33
      Top             =   6840
      Width           =   3135
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "TODAY TOTAL "
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11520
      TabIndex        =   32
      Top             =   7800
      Width           =   3495
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3480
      TabIndex        =   31
      Top             =   7560
      Width           =   3135
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   720
      TabIndex        =   30
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   29
      Top             =   6240
      Width           =   3135
   End
   Begin VB.TextBox txttot 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   28
      Top             =   6960
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtcoins 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   25
      Text            =   "0"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txt5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Text            =   "0"
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txt10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Text            =   "0"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txt20 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Text            =   "0"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txt50 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      Text            =   "0"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txt100 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Text            =   "0"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txt500 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Text            =   "0"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txt1000 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   17
      Text            =   "0"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txt2000 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   16
      Text            =   "0"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txteno 
      DataField       =   "eno"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
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
      Left            =   5190
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1440
   End
   Begin VB.TextBox txtdate 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      DataSource      =   "Adodc1"
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
      Left            =   3540
      TabIndex        =   3
      Text            =   "01/10/1996"
      Top             =   960
      Width           =   1665
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1485
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
      Height          =   405
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1365
   End
   Begin VB.Shape Shape3 
      Height          =   1455
      Left            =   480
      Top             =   8760
      Width           =   19335
   End
   Begin VB.Label Labmonth 
      BackColor       =   &H00FF0000&
      Caption         =   "Label3 "
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   19800
      TabIndex        =   38
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Labdayname 
      BackColor       =   &H00FF0000&
      Caption         =   "wednesday        "
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   18480
      TabIndex        =   37
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Labscroll 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "hghvuhjvuv"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   36
      Top             =   10800
      Width           =   20655
   End
   Begin VB.Label Labshowtd 
      BackColor       =   &H00FF0000&
      Caption         =   "time"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"MINE.frx":0000
      BeginProperty Font 
         Name            =   "@Meiryo"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2295
      Left            =   600
      TabIndex        =   34
      Top             =   8880
      Width           =   21495
   End
   Begin VB.Shape Shape2 
      Height          =   7695
      Left            =   6840
      Top             =   840
      Width           =   12975
   End
   Begin VB.Shape Shape1 
      Height          =   7695
      Left            =   480
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Latot 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   27
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label La5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   24
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lacoins 
      Caption         =   "COINS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label la10 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label la20 
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label la50 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label la100 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label la500 
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label labeno 
      Caption         =   "  E.NO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label la1000 
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label la2000 
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label labname 
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   " ACCOUNTS  ASSISTANT"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim con2 As New ADODB.Connection

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset

Dim k As Integer
Dim j As Integer


Private Sub Calendar1_Click()
txtdate = Calendar1.Value

End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
cmddelete.Enabled = True
    
    Else
cmddelete.Enabled = False
    
    End If

End Sub

Private Sub Cmdclose_Click()
End

Unload Form1
Unload Form2
Unload Form3
Unload form4
Unload Me

End Sub

Private Sub cmddelete_Click()

txtdate.SetFocus
txtdate = flex1.TextMatrix(flex1.RowSel, 0)
txteno = flex1.TextMatrix(flex1.RowSel, 1)
Combo1.Text = flex1.TextMatrix(flex1.RowSel, 2)
txt2000.Text = flex1.TextMatrix(flex1.RowSel, 3)
txt1000.Text = flex1.TextMatrix(flex1.RowSel, 4)
txt500.Text = flex1.TextMatrix(flex1.RowSel, 5)
txt100.Text = flex1.TextMatrix(flex1.RowSel, 6)
txt50.Text = flex1.TextMatrix(flex1.RowSel, 7)
txt20.Text = flex1.TextMatrix(flex1.RowSel, 8)
txt10.Text = flex1.TextMatrix(flex1.RowSel, 9)
txt5.Text = flex1.TextMatrix(flex1.RowSel, 10)
txtcoins.Text = flex1.TextMatrix(flex1.RowSel, 11)
txttot.Text = flex1.TextMatrix(flex1.RowSel, 12)
rs.Requery
rs.Move (flex1.RowSel - 1)
rs.Delete
If Not rs.BOF Then
rs.MovePrevious
End If
If Not rs.EOF Then
rs.MoveNext
End If
Call display
End Sub



Private Sub cmdedit_Click()
PREV.Enabled = True



txtdate.SetFocus
txtdate = flex1.TextMatrix(flex1.RowSel, 0)
txteno = flex1.TextMatrix(flex1.RowSel, 1)
Combo1.Text = flex1.TextMatrix(flex1.RowSel, 2)
txt2000.Text = flex1.TextMatrix(flex1.RowSel, 3)
txt1000.Text = flex1.TextMatrix(flex1.RowSel, 4)
txt500.Text = flex1.TextMatrix(flex1.RowSel, 5)
txt100.Text = flex1.TextMatrix(flex1.RowSel, 6)
txt50.Text = flex1.TextMatrix(flex1.RowSel, 7)
txt20.Text = flex1.TextMatrix(flex1.RowSel, 8)
txt10.Text = flex1.TextMatrix(flex1.RowSel, 9)
txt5.Text = flex1.TextMatrix(flex1.RowSel, 10)
txtcoins.Text = flex1.TextMatrix(flex1.RowSel, 11)
txttot.Text = flex1.TextMatrix(flex1.RowSel, 12)





End Sub


Private Sub cmdnew_Click()
Combo1.Enabled = True
cmdnew.Enabled = False
MsgBox "Enter Data"
'Combo1.Text = " "
txteno.Text = " "
txt2000.Text = " "
txt1000.Text = " "
txt500.Text = " "
txt100.Text = " "
txt50.Text = " "
txt20.Text = " "
txt10.Text = " "
txt5.Text = " "
txtcoins.Text = " "
txtdate.SetFocus
rs.AddNew

cmdupdate.Enabled = True
cmdsearch.Enabled = False
cmdedit.Enabled = False
'cmddelete.Enabled = False
PREV.Enabled = False


Combo1.Clear
Call loadname


End Sub



Private Sub cmdnext_Click()
Form3.Show
Form3.Calendar1.Value = Form2.Calendar1
Form3.Calendar2.Value = Form2.Calendar1
Form3.txtfrom = Form3.Calendar1.Value
Form3.txtto = Form3.Calendar2.Value
Form3.Combo2.Enabled = False
Form3.Command1.Enabled = False
Form3.cmdsplitauto.Enabled = True
Form3.Cmdsplitemanu.Enabled = True




d2 = Format(Form3.Calendar2.Value, "dd/mm/yyyy")
D1 = Format(Form3.Calendar1.Value, "dd/mm/YYYY")   'todays seraching events
MsgBox (D1)
Call displays

End Sub

Private Sub cmdSearch_Click()
Form3.Show
Form3.Combo2.Enabled = True
Form3.Command1.Enabled = True
Form3.cmdsplitauto.Visible = False
Form3.Cmdsplitemanu.Visible = False


End Sub


Private Sub cmdupdate_Click()
cmdupdate.Enabled = False
cmdnew.Enabled = True
cmdsearch.Enabled = True
cmdedit.Enabled = True
cmdedit.Enabled = True




rs.Fields(0) = txtdate.Text
rs.Fields(2) = Combo1.Text
rs.Fields(1) = txteno.Text
rs.Fields(3) = txt2000.Text
rs.Fields(4) = txt1000.Text
rs.Fields(5) = txt500.Text
rs.Fields(6) = txt100.Text
rs.Fields(7) = txt50.Text
rs.Fields(8) = txt20.Text
rs.Fields(9) = txt10.Text
rs.Fields(10) = txt5.Text
rs.Fields(11) = txtcoins.Text
rs.Fields(12) = txttot

rs.Update

MsgBox "Record is Saved"
Call display
End Sub

Private Sub Combo1_Click()
Dim con4 As ADODB.Connection
Set con4 = New ADODB.Connection
Dim rs4 As ADODB.Recordset
Set rs4 = New ADODB.Recordset
con4.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\vendata.MDB;Persist Security Info=False"
rs4.Open "SELECT eno FROM venname WHERE name LIKE " & "'" & Combo1.Text & "'", con, adOpenStatic, adLockOptimistic
txteno.Text = rs4.Fields(0)
rs4.Close
con4.Close

End Sub
  

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then

Form6.Show

End If
End Sub


Private Sub Form_Load()

flex1.ColWidth(11) = 1100
flex1.ColWidth(10) = 900
flex1.ColWidth(9) = 900
flex1.ColWidth(8) = 900
flex1.ColWidth(7) = 900
flex1.ColWidth(6) = 900
flex1.ColWidth(5) = 900
flex1.ColWidth(4) = 900
flex1.ColWidth(3) = 900
flex1.ColWidth(2) = 1400
flex1.ColWidth(1) = 547
flex1.ColWidth(0) = 1100


con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
   & App.Path & "\vendata.mdb;Persist Security Info=False"
rs.Open "Select * from ventab ", con, adOpenDynamic, adLockOptimistic
If rs.BOF = True Then
rs.MoveFirst
MsgBox ("no record found")

Else
Call display
End If

Call calculate



Calendar1 = DateValue(Now)
txtdate = Calendar1.Value
Labshowtd = Now
Labdayname = WeekdayName(Weekday(Now))
Labmonth = MonthName(Month(Now))

cmdupdate.Enabled = False
PREV.Enabled = False
Combo1.Enabled = False
cmddelete.Enabled = False
Check1.Value = 0


 
 Labscroll.Caption = "          'GOD IS LOVE'        ,         'The  Pen  is  mightier  than  the  Sword'        ,      'Practice  makes  Perfect'          ,        'You  can't  judge  a  book  by  its  Cover'           ,          'Honesty  is  the  best  Policy'          ,            Where  there’s  a  will  there’s  a  Way           ,           'Look before you Leap'"
 Timer1.Enabled = True
 Timer1.Interval = 500

Call loadname

End Sub
Private Sub loadname()

con2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
   & App.Path & "\vendata.mdb;Persist Security Info=False"
rs3.Open "Select * from venname ", con, adOpenDynamic, adLockOptimistic

If rs3.RecordCount <> 0 Then
Do Until rs3.EOF
With Combo1
.AddItem rs3.Fields("name")
End With
rs3.MoveNext
Loop
Else
MsgBox "No Item To Load"
End If
rs3.Close
con2.Close


End Sub
Private Sub calculate()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex1.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex1.TextMatrix(i, 3))
i = i + 1
Wend
End Sub
Private Sub display()

k = 0

flex1.Cols = 13
flex1.Rows = 1
flex1.Row = k
flex1.Col = 0

flex1.Text = "DATE"
flex1.Col = 1
flex1.Text = "ENO"
flex1.Col = 2
flex1.Text = "NAME"
flex1.Col = 3
flex1.Text = "2000"
flex1.Col = 4
flex1.Text = "1000"
flex1.Col = 5
flex1.Text = "500"
flex1.Col = 6
flex1.Text = "100"
flex1.Col = 7
flex1.Text = "50"
flex1.Col = 8
flex1.Text = "20"
flex1.Col = 9
flex1.Text = "10"
flex1.Col = 10
flex1.Text = "5"
flex1.Col = 11
flex1.Text = "COINS"
flex1.Col = 12
flex1.Text = "TOTAL"

j = 0
k = k + 1
rs.MoveFirst
While rs.EOF <> True
flex1.Rows = flex1.Rows + 1




flex1.Row = flex1.Row + 1
flex1.Col = j
flex1.Text = rs.Fields(0)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(1)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(2)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(3)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(4)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(5)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(6)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(7)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(8)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(9)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(10)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(11)
flex1.Col = flex1.Col + 1
flex1.Text = rs.Fields(12)
k = k + 1
rs.MoveNext

Wend



End Sub

Private Sub LAST_Click()

rs.Fields(0) = txtdate.Text
rs.Fields(2) = Combo1.Text
rs.Fields(1) = txteno.Text
rs.Fields(3) = txt2000.Text
rs.Fields(4) = txt1000.Text
rs.Fields(5) = txt500.Text
rs.Fields(6) = txt100.Text
rs.Fields(7) = txt50.Text
rs.Fields(8) = txt20.Text
rs.Fields(9) = txt10.Text
rs.Fields(10) = txt5.Text
rs.Fields(11) = txtcoins.Text
rs.Fields(12) = txttot

rs.Save
End Sub







Private Sub PREV_Click()
rs.Requery
rs.Move (flex1.RowSel - 1)


rs.Fields(0) = txtdate.Text
rs.Fields(2) = Combo1.Text
rs.Fields(1) = txteno.Text
rs.Fields(3) = txt2000.Text
rs.Fields(4) = txt1000.Text
rs.Fields(5) = txt500.Text
rs.Fields(6) = txt100.Text
rs.Fields(7) = txt50.Text
rs.Fields(8) = txt20.Text
rs.Fields(9) = txt10.Text
rs.Fields(10) = txt5.Text
rs.Fields(11) = txtcoins.Text
rs.Fields(12) = txttot

rs.Update

MsgBox "Record is Saved"
Call display

End Sub

Private Sub Timer1_Timer()
Dim str As String
str = Form2.Labscroll.Caption
str = Mid$(str, 2, Len(str)) + Left(str, 1)
Form2.Labscroll.Caption = str
End Sub

Private Sub txt10_Change()
txttot.Text = Val(txt2000) * 2000 + Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txtcoins) * 1

End Sub

Private Sub txt100_Change()
txttot.Text = Val(txt2000) * 2000 + Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txtcoins) * 1

End Sub

Private Sub txt1000_Change()
txttot.Text = Val(txt2000) * 2000 + Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txtcoins) * 1

End Sub

Private Sub txt20_Change()
txttot.Text = Val(txt2000) * 2000 + Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txtcoins) * 1

End Sub

Private Sub txt2000_Change()
txttot.Text = Val(txt2000) * 2000 + Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txtcoins) * 1

End Sub

Private Sub txt5_Change()
txttot.Text = Val(txt2000) * 2000 + Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txtcoins) * 1

End Sub

Private Sub txt50_Change()
txttot.Text = Val(txt2000) * 2000 + Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txtcoins) * 1

End Sub

Private Sub txt500_Change()
txttot.Text = Val(txt2000) * 2000 + Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txtcoins) * 1

End Sub

Private Sub txtcoins_Change()
txttot.Text = Val(txt2000) * 2000 + Val(txt1000) * 1000 + Val(txt500) * 500 + Val(txt100) * 100 + Val(txt50) * 50 + Val(txt20) * 20 + Val(txt10) * 10 + Val(txt5) * 5 + Val(txtcoins) * 1

End Sub





Private Sub displays()
k = 0

Form3.flex2.Cols = 13
Form3.flex2.Rows = 1
Form3.flex2.Row = k
Form3.flex2.Col = 0

Form3.flex2.Text = "DATE"
Form3.flex2.Col = 1
Form3.flex2.Text = "ENO"
Form3.flex2.Col = 2
Form3.flex2.Text = "NAME"
Form3.flex2.Col = 3
Form3.flex2.Text = "2000"
Form3.flex2.Col = 4
Form3.flex2.Text = "1000"
Form3.flex2.Col = 5
Form3.flex2.Text = "500"
Form3.flex2.Col = 6
Form3.flex2.Text = "100"
Form3.flex2.Col = 7
Form3.flex2.Text = "50"
Form3.flex2.Col = 8
Form3.flex2.Text = "20"
Form3.flex2.Col = 9
Form3.flex2.Text = "10"
Form3.flex2.Col = 10
Form3.flex2.Text = "5"
Form3.flex2.Col = 11
Form3.flex2.Text = "COINS"
Form3.flex2.Col = 12
Form3.flex2.Text = "TOTAL"

j = 0
k = k + 1
Dim da1, da2, mo1, mo2, ye1, ye2 As Integer
da1 = Day(D1)
mo1 = Month(D1)
ye1 = Year(D1)
da2 = Day(d2)
mo2 = Month(d2)
ye2 = Year(d2)

'rs5.Open "Select * from ventab where day(date) =" & da1 _
 '                                              & " and month(date) =" & mo1 _
  '                                             & " and year(date)=" & ye1 _
   '                                            , con, adOpenDynamic, adLockOptimistic



rs5.Open "Select * from ventab where day(date) >=" & da1 _
                                    & " and day(date) <=" & da2 _
                                    & " and month(date) >=" & mo1 _
                                    & " and month(date) <=" & mo2 _
                                    & " and year(date) >=" & ye1 _
                                   & " and year(date) <=" & ye2 _
                                    , con, adOpenDynamic, adLockOptimistic



While rs5.EOF <> True

Form3.flex2.Rows = Form3.flex2.Rows + 1
Form3.flex2.Row = Form3.flex2.Row + 1
Form3.flex2.Col = j

Form3.flex2.Text = rs2.Fields(0)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(1)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(2)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(3)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(4)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(5)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(6)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(7)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(8)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(9)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(10)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(11)
Form3.flex2.Col = Form3.flex2.Col + 1
Form3.flex2.Text = rs2.Fields(12)
k = k + 1
rs5.MoveNext

Wend

Call calculate2000
Call calculate1000
Call calculate500
Call calculate100
Call calculate50
Call calculate20
Call calculate10
Call calculate5
Call calculatecoins
Call calculateftotal



rs5.Close
End Sub





Private Sub calculate2000()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 3))
i = i + 1
Wend

Form3.Txtto2000.Text = str(no1)
End Sub
Private Sub calculate1000()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 4))
i = i + 1
Wend

Form3.txtto1000.Text = str(no1)
End Sub

Private Sub calculate500()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 5))
i = i + 1
Wend

Form3.txtto500.Text = str(no1)
End Sub
Private Sub calculate100()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 6))
i = i + 1
Wend

Form3.txtto100.Text = str(no1)
End Sub
Private Sub calculate50()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 7))
i = i + 1
Wend

Form3.txtto50.Text = str(no1)
End Sub
Private Sub calculate20()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 8))
i = i + 1
Wend

Form3.txtto20.Text = str(no1)
End Sub
Private Sub calculate10()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 9))
i = i + 1
Wend

Form3.txtto10.Text = str(no1)
End Sub
Private Sub calculate5()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 10))
i = i + 1
Wend

Form3.txtto5.Text = str(no1)
End Sub
Private Sub calculatecoins()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 11))
i = i + 1
Wend

Form3.txttocoins.Text = str(no1)
End Sub
Private Sub calculateftotal()
Dim i As Integer
Dim no1 As Long
Dim no As Long
i = 1
no = Form3.flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(Form3.flex2.TextMatrix(i, 12))
i = i + 1
Wend

Form3.txttoto.Text = str(no1)
End Sub
