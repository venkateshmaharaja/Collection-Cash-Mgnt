VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "LOGIN"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
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
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5040
      TabIndex        =   0
      Top             =   1920
      Width           =   9135
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
         Picture         =   "colg.frx":0000
         ScaleHeight     =   2535
         ScaleWidth      =   4455
         TabIndex        =   7
         Top             =   120
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
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
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
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   450
         Left            =   960
         TabIndex        =   1
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Lapwd 
         BackColor       =   &H00FFFFC0&
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
         BackColor       =   &H00FFFFC0&
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
 
 'MsgBox "username and password correct", vbInformation, "login"
 
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
Frame1.BackColor = RGB(58, 71, 248)
labuname.BackColor = RGB(58, 71, 248)
Lapwd.BackColor = RGB(58, 71, 248)
tbpassword.BackColor = RGB(0, 0, 0)
tbusername.BackColor = RGB(0, 0, 0)

End Sub

