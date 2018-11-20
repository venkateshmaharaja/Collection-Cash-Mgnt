VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00C0C000&
   Caption         =   "Form4"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex2 
      Height          =   7815
      Left            =   6000
      TabIndex        =   24
      Top             =   1320
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   13785
      _Version        =   393216
   End
   Begin MSACAL.Calendar Calendar2 
      Height          =   3135
      Left            =   1440
      TabIndex        =   23
      Top             =   6000
      Width           =   4455
      _Version        =   524288
      _ExtentX        =   7858
      _ExtentY        =   5530
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
   Begin MSACAL.Calendar Calendar1 
      Height          =   3255
      Left            =   1440
      TabIndex        =   22
      Top             =   1680
      Width           =   4455
      _Version        =   524288
      _ExtentX        =   7858
      _ExtentY        =   5741
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
   Begin VB.CommandButton Cmdsplitemanu 
      Caption         =   "SPLITE(manual)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   0
      Top             =   9960
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N.SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3720
      TabIndex        =   2
      Top             =   9240
      Width           =   2175
   End
   Begin VB.CommandButton cmdsplitauto 
      Caption         =   "SPLITE(auto)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   21
      Top             =   9960
      Width           =   6495
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "  <<BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1440
      TabIndex        =   20
      Top             =   9840
      Width           =   4455
   End
   Begin VB.TextBox Txtto2000 
      Height          =   405
      Left            =   9120
      TabIndex        =   18
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txtto1000 
      Height          =   405
      Left            =   10080
      TabIndex        =   17
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txtto500 
      Height          =   405
      Left            =   11040
      TabIndex        =   16
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txtto100 
      Height          =   405
      Left            =   12000
      TabIndex        =   15
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txtto20 
      Height          =   405
      Left            =   13920
      TabIndex        =   14
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txtto50 
      Height          =   405
      Left            =   12960
      TabIndex        =   13
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txtto10 
      Height          =   405
      Left            =   14880
      TabIndex        =   12
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txtto5 
      Height          =   405
      Left            =   15840
      TabIndex        =   11
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txttocoins 
      Height          =   405
      Left            =   16800
      TabIndex        =   10
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txttoto 
      Height          =   405
      Left            =   17775
      TabIndex        =   9
      Top             =   9360
      Width           =   1020
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1320
      Width           =   4455
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "FIND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1440
      TabIndex        =   7
      Top             =   9240
      Width           =   2295
   End
   Begin VB.TextBox txtto 
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox txtfrom 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "TOTAL COLLECTION"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      Height          =   9615
      Left            =   1080
      Top             =   960
      Width           =   17895
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
      Left            =   6000
      TabIndex        =   19
      Top             =   9360
      Width           =   3015
   End
   Begin VB.Label labtodat 
      Caption         =   " TO   DATE         :             :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Labfromdt 
      Caption         =   "FROM DATE      :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   5040
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset

Dim D1 As Date
Dim d2 As Date
Dim k As Integer
Dim j As Integer


Private Sub Calendar1_Click()
txtfrom = Calendar1.Value
End Sub

Private Sub Calendar2_Click()
txtto = Calendar2.Value
End Sub

Private Sub cmdback_Click()
Form2.Show
End Sub


Private Sub cmdfind_Click()


D1 = Format(Calendar1.Value, "dd/mm/YYYY")
d2 = Format(Calendar2.Value, "dd/mm/YYYY")


MsgBox (D1)
MsgBox (d2)
Call display
End Sub

Private Sub cmdsplitauto_Click()

Call plushide      'plus button hidden

'Splite to the cash on bank is 2000 to 100
'And cash in hand only denoted to 50 to coins

form4.Show
form4.txtcb2000 = Txtto2000   'form3 value store on form4
form4.txtcb1000 = txtto1000
form4.txtcb500 = txtto500
form4.txtcb100 = txtto100

form4.Labcb2000 = Val(form4.txtcb2000 * 2000) 'calculation on cash in bank
form4.Labcb1000 = Val(form4.txtcb1000 * 1000)
form4.Labcb500 = Val(form4.txtcb500 * 500)
form4.Labcb100 = Val(form4.txtcb100 * 100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100)



'------------------------------------------------------
form4.txtch50 = txtto50 'form3 value store on form4
form4.txtch20 = txtto20
form4.txtch10 = txtto10
form4.txtch5 = txtto5
form4.txtchcoins = txttocoins

form4.Labch50 = Val(form4.txtch50 * 50)   'cal on cash in hand
form4.Labch20 = Val(form4.txtch20 * 20)
form4.Labch10 = Val(form4.txtch10 * 10)
form4.Labch5 = Val(form4.txtch5 * 5)
form4.Labchcoins = Val(form4.txtchcoins * 1)
form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins)

End Sub

Private Sub Cmdsplitemanu_Click()

Call plusshow  'plus button shown

'Splite to the cash on bank is 2000 to 100
'And cash in hand only denoted to 50 to coins

form4.Show
form4.txtcb2000 = Txtto2000   'form3 value store on form4
form4.txtcb1000 = txtto1000
form4.txtcb500 = txtto500
form4.txtcb100 = txtto100

form4.Labcb2000 = Val(form4.txtcb2000 * 2000) 'calculation on cash in bank
form4.Labcb1000 = Val(form4.txtcb1000 * 1000)
form4.Labcb500 = Val(form4.txtcb500 * 500)
form4.Labcb100 = Val(form4.txtcb100 * 100)
form4.totcbank = Val(form4.Labcb2000) + Val(form4.Labcb1000) + Val(form4.Labcb500) + Val(form4.Labcb100)



'------------------------------------------------------
form4.txtch50 = txtto50 'form3 value store on form4
form4.txtch20 = txtto20
form4.txtch10 = txtto10
form4.txtch5 = txtto5
form4.txtchcoins = txttocoins

form4.Labch50 = Val(form4.txtch50 * 50)   'cal on cash in hand
form4.Labch20 = Val(form4.txtch20 * 20)
form4.Labch10 = Val(form4.txtch10 * 10)
form4.Labch5 = Val(form4.txtch5 * 5)
form4.Labchcoins = Val(form4.txtchcoins * 1)
form4.totchand = Val(form4.Labch50) + Val(form4.Labch20) + Val(form4.Labch10) + Val(form4.Labch5) + Val(form4.Labchcoins)

End Sub


Private Sub loadname()

'con2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
   & App.Path & "\vendata.mdb;Persist Security Info=False"
rs3.Open "Select * from venname ", con, adOpenDynamic, adLockPessimistic


If rs3.RecordCount <> 0 Then
Do Until rs3.EOF
With Combo2
.AddItem rs3.Fields("name")
End With
rs3.MoveNext
Loop
Else
MsgBox "No Item To Load"
End If
rs3.Close
'con2.Close


End Sub


Private Sub Command1_Click()
k = 0

flex2.Cols = 13
flex2.Rows = 1
flex2.Row = k
flex2.Col = 0

flex2.Text = "DATE"
flex2.Col = 1
flex2.Text = "ENO"
flex2.Col = 2
flex2.Text = "NAME"
flex2.Col = 3
flex2.Text = "2000"
flex2.Col = 4
flex2.Text = "1000"
flex2.Col = 5
flex2.Text = "500"
flex2.Col = 6
flex2.Text = "100"
flex2.Col = 7
flex2.Text = "50"
flex2.Col = 8
flex2.Text = "20"
flex2.Col = 9
flex2.Text = "10"
flex2.Col = 10
flex2.Text = "5"
flex2.Col = 11
flex2.Text = "COINS"
flex2.Col = 12
flex2.Text = "TOTAL"

j = 0
k = k + 1
a = Combo2.Text
MsgBox (a)
rs2.Open "Select * from ventab where Name  ='" & a & "'", con, adOpenDynamic, adLockPessimistic



While rs2.EOF <> True

flex2.Rows = flex2.Rows + 1
flex2.Row = flex2.Row + 1
flex2.Col = j

flex2.Text = rs2.Fields(0)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(1)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(2)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(3)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(4)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(5)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(6)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(7)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(8)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(9)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(10)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(11)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(12)
k = k + 1
rs2.MoveNext

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



rs2.Close
End Sub

Private Sub Form_Load()

flex2.ColWidth(2) = 1400
flex2.ColWidth(1) = 547
flex2.ColWidth(0) = 1100



con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
   & App.Path & "\vendata.mdb;Persist Security Info=False"
rs1.Open "Select * from ventab ", con, adOpenDynamic, adLockPessimistic
rs1.MoveFirst
If rs1.BOF = True Then
MsgBox ("no record found")
Else

End If

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


Call loadname
End Sub

Private Sub plusshow()
form4.pluscbcoins.Visible = True
form4.pluscb5.Visible = True
form4.pluscb10.Visible = True
form4.pluscb20.Visible = True
form4.pluscb50.Visible = True
form4.pluscb100.Visible = True
form4.pluscb500.Visible = True
form4.pluscb1000.Visible = True
form4.pluscb2000.Visible = True

form4.pluschcoins.Visible = True
form4.plusch5.Visible = True
form4.plusch10.Visible = True
form4.plusch20.Visible = True
form4.plusch50.Visible = True
form4.plusch100.Visible = True
form4.plusch500.Visible = True
form4.plusch1000.Visible = True
form4.plusch2000.Visible = True



End Sub

Private Sub plushide()
form4.pluscbcoins.Visible = False
form4.pluscb5.Visible = False
form4.pluscb10.Visible = False
form4.pluscb20.Visible = False
form4.pluscb50.Visible = False
form4.pluscb100.Visible = False
form4.pluscb500.Visible = False
form4.pluscb1000.Visible = False
form4.pluscb2000.Visible = False

form4.pluschcoins.Visible = False
form4.plusch5.Visible = False
form4.plusch10.Visible = False
form4.plusch20.Visible = False
form4.plusch50.Visible = False
form4.plusch100.Visible = False
form4.plusch500.Visible = False
form4.plusch1000.Visible = False
form4.plusch2000.Visible = False


End Sub


Private Sub calculate2000()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 3))
i = i + 1
Wend

Txtto2000.Text = str(no1)
End Sub
Private Sub calculate1000()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 4))
i = i + 1
Wend

txtto1000.Text = str(no1)
End Sub

Private Sub calculate500()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 5))
i = i + 1
Wend

txtto500.Text = str(no1)
End Sub
Private Sub calculate100()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 6))
i = i + 1
Wend

txtto100.Text = str(no1)
End Sub
Private Sub calculate50()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 7))
i = i + 1
Wend

txtto50.Text = str(no1)
End Sub
Private Sub calculate20()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 8))
i = i + 1
Wend

txtto20.Text = str(no1)
End Sub
Private Sub calculate10()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 9))
i = i + 1
Wend

txtto10.Text = str(no1)
End Sub
Private Sub calculate5()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 10))
i = i + 1
Wend

txtto5.Text = str(no1)
End Sub
Private Sub calculatecoins()
Dim i As Integer
Dim no1 As Integer
Dim no As Integer
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 11))
i = i + 1
Wend

txttocoins.Text = str(no1)
End Sub
Private Sub calculateftotal()
Dim i As Integer
Dim no1 As Long
Dim no As Long
i = 1
no = flex2.Rows
no1 = 0
While (i < no)
no1 = no1 + Val(flex2.TextMatrix(i, 12))
i = i + 1
Wend

txttoto.Text = str(no1)
End Sub

Private Sub display()
k = 0

flex2.Cols = 13
flex2.Rows = 1
flex2.Row = k
flex2.Col = 0

flex2.Text = "DATE"
flex2.Col = 1
flex2.Text = "ENO"
flex2.Col = 2
flex2.Text = "NAME"
flex2.Col = 3
flex2.Text = "2000"
flex2.Col = 4
flex2.Text = "1000"
flex2.Col = 5
flex2.Text = "500"
flex2.Col = 6
flex2.Text = "100"
flex2.Col = 7
flex2.Text = "50"
flex2.Col = 8
flex2.Text = "20"
flex2.Col = 9
flex2.Text = "10"
flex2.Col = 10
flex2.Text = "5"
flex2.Col = 11
flex2.Text = "COINS"
flex2.Col = 12
flex2.Text = "TOTAL"

j = 0
k = k + 1
Dim da1, da2, mo1, mo2, ye1, ye2 As Integer
da1 = Day(D1)
mo1 = Month(D1)
ye1 = Year(D1)
da2 = Day(d2)
mo2 = Month(d2)
ye2 = Year(d2)

rs2.Open "Select * from ventab where day(date) >=" & da1 _
                                    & " and day(date) <=" & da2 _
                                    & " and month(date) >=" & mo1 _
                                    & " and month(date) <=" & mo2 _
                                    & " and year(date) >=" & ye1 _
                                    & " and year(date) <=" & ye2 _
                                    , con, adOpenDynamic, adLockPessimistic
                                    


While rs2.EOF <> True

flex2.Rows = flex2.Rows + 1
flex2.Row = flex2.Row + 1
flex2.Col = j

flex2.Text = rs2.Fields(0)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(1)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(2)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(3)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(4)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(5)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(6)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(7)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(8)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(9)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(10)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(11)
flex2.Col = flex2.Col + 1
flex2.Text = rs2.Fields(12)
k = k + 1
rs2.MoveNext

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



rs2.Close
End Sub

