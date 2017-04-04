VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "Rockwell Condensed"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   8655
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9060
      Left            =   0
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   0
      Width           =   12060
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Cancel          =   -1  'True
         Caption         =   "Back to start"
         Height          =   615
         Left            =   360
         MaskColor       =   &H000000C0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7560
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "Form6.frx":4A745
         Top             =   1680
         Width           =   10335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "About Indian Judiciary"
         BeginProperty Font 
            Name            =   "Rockwell Condensed"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   7095
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form6.Hide
End Sub
