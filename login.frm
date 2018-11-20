VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "LOGIN"
   ClientHeight    =   8460
   ClientLeft      =   2745
   ClientTop       =   1155
   ClientWidth     =   14520
   BeginProperty Font 
      Name            =   "Adobe Naskh Medium"
      Size            =   8.25
      Charset         =   0
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   30731.89
   ScaleMode       =   0  'User
   ScaleWidth      =   7386.182
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   12960
      Picture         =   "login.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   7695
      TabIndex        =   17
      Top             =   8160
      Width           =   7695
   End
   Begin VB.PictureBox Picture8 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   6960
      Picture         =   "login.frx":136336
      ScaleHeight     =   2535
      ScaleWidth      =   6015
      TabIndex        =   16
      Top             =   8160
      Width           =   6015
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   -120
      Picture         =   "login.frx":13DE13
      ScaleHeight     =   2535
      ScaleWidth      =   7095
      TabIndex        =   15
      Top             =   8160
      Width           =   7095
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      Picture         =   "login.frx":274149
      ScaleHeight     =   1935
      ScaleWidth      =   7695
      TabIndex        =   13
      Top             =   0
      Width           =   7695
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   12600
      Picture         =   "login.frx":3AA47F
      ScaleHeight     =   2055
      ScaleWidth      =   9375
      TabIndex        =   12
      Top             =   -120
      Width           =   9375
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      Picture         =   "login.frx":4E07B5
      ScaleHeight     =   6375
      ScaleWidth      =   5535
      TabIndex        =   11
      Top             =   1920
      Width           =   5535
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   14640
      Picture         =   "login.frx":56FA8F
      ScaleHeight     =   6735
      ScaleWidth      =   5895
      TabIndex        =   10
      Top             =   1920
      Width           =   5895
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   3120
      Picture         =   "login.frx":5E9521
      ScaleHeight     =   4335
      ScaleWidth      =   11295
      TabIndex        =   9
      Top             =   -2400
      Width           =   11295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   245
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Prestige Elite Std"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6375
      Left            =   5520
      TabIndex        =   0
      Top             =   1920
      Width           =   9135
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   4680
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   0
         Max             =   105
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   2160
         Picture         =   "login.frx":5EE371
         ScaleHeight     =   2535
         ScaleWidth      =   4455
         TabIndex        =   7
         Top             =   240
         Width           =   4455
      End
      Begin VB.CommandButton CMDCANCEL 
         BackColor       =   &H00808080&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5400
         Width           =   2000
      End
      Begin VB.CommandButton CMDOK 
         BackColor       =   &H00808080&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1800
         MaskColor       =   &H00004080&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         Width           =   2000
      End
      Begin VB.TextBox tbpassword 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3960
         Width           =   3855
      End
      Begin VB.TextBox tbusername 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   960
         TabIndex        =   1
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Lapwd 
         BackColor       =   &H00FFFF00&
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5760
         TabIndex        =   4
         Top             =   3960
         Width           =   1995
      End
      Begin VB.Label labuname 
         BackColor       =   &H00FFFF00&
         Caption         =   "USER NAME"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5760
         TabIndex        =   3
         Top             =   2880
         Width           =   1995
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Developed by : A.Venkatesh (T.L)  ,    S.Vairammanikandan    ,    B.Surya kumar "
      BeginProperty Font 
         Name            =   "Adobe Naskh Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   10680
      Width           =   21015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim user As String
Dim pass As String



Private Sub CMDCANCEL_Click()
End

End Sub

Private Sub CMDOK_Click()

user = "venkat"
pass = "vsv"
If tbusername.Text = user Then
 If tbpassword.Text = pass Then
Unload Me
 
 Form2.Enabled = True
 Form2.Show
 ElseIf tbpassword.Text = "" Then
 MsgBox "password field empty", vbExclamation, "login"
 
Else
MsgBox "username and password not matched", vbExclamation, "login"
End If
ElseIf tbusername.Text = "" Then
MsgBox "username field empty", vbExclamation, "login"
Else
MsgBox "invalied username,try again", , "login"
tbpassword.SetFocus
End If
End Sub


Private Sub Form_Load()

Form1.BackColor = RGB(250, 240, 1)
Frame1.BackColor = RGB(1, 0, 112)
labuname.BackColor = RGB(58, 7, 248)
Lapwd.BackColor = RGB(58, 71, 248)
tbpassword.BackColor = RGB(248, 248, 248)
tbusername.BackColor = RGB(248, 248, 248)
labuname.Visible = False
Lapwd.Visible = False
tbpassword.Visible = False
tbusername.Visible = False

Timer1.Enabled = True

End Sub


Private Sub Picture10_Click()

End Sub

Private Sub Timer1_Timer()
 ProgressBar1.Value = ProgressBar1.Value + 5

If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
labuname.Visible = True
Lapwd.Visible = True
tbpassword.Visible = True
tbusername.Visible = True
ProgressBar1.Visible = False

End If
End Sub
