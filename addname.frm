VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00404000&
   Caption         =   "Form6"
   ClientHeight    =   5430
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8700
   LinkTopic       =   "Form6"
   ScaleHeight     =   5430
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdnsave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox enoptxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox nameptxt 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Labeno 
      Caption         =   "ENO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Labpname 
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Menu mnuadd 
      Caption         =   "&ADD"
   End
   Begin VB.Menu mnudelete 
      Caption         =   "&DELETE"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs3 As New ADODB.Recordset

Private Sub cmdnsave_Click()
rs3.Fields(0) = nameptxt.Text
rs3.Fields(1) = enoptxt.Text
rs3.Update

MsgBox ("new particular added")



End Sub

Private Sub Command1_Click()
Form2.Refresh
Form2.Combo1.Refresh

Unload Me
rs3.Close
con.Close


End Sub

Private Sub Form_Load()

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
   & App.Path & "\vendata.mdb;Persist Security Info=False"
rs3.Open "Select * from venname ", con, adOpenDynamic, adLockPessimistic


rs3.MoveLast
nameptxt.Text = rs3.Fields(0)
enoptxt.Text = rs3.Fields(1)



'If rs.BOF = True Then
'rs.MoveFirst
'MsgBox ("no record found")

'Else


End Sub

Private Sub mnuadd_Click()
a = Val(enoptxt)
B = a + 1
rs3.AddNew
nameptxt.Text = ""
enoptxt.Text = B
MsgBox ("enter the participant")
End Sub

