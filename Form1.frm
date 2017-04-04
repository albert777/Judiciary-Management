VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   7230
   ClientTop       =   4200
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Segoe UI Symbol"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4815
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000015&
         Caption         =   "Login"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Forgot Password"
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Admin ID:"
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.Text = "admin" And Text1.Text = "admin" Then
Form1.Hide
frmSplash.Show
Else
MsgBox ("Invalid Username/password")
End If
End Sub

Private Sub Label3_Click()
MsgBox "Default username:admin Default password: admin", vbOKOnly
End Sub
