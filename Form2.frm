VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8370
   ClientLeft      =   4005
   ClientTop       =   990
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Maiandra GD"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8370
   ScaleWidth      =   12015
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
      Height          =   8175
      Left            =   120
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   8115
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.CommandButton Command5 
         Caption         =   "About Judiciary Of India"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4200
         TabIndex        =   8
         Top             =   5400
         Width           =   3015
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Exit Portal"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         TabIndex        =   7
         Top             =   4200
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Search Record"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7200
         TabIndex        =   6
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Update record"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   5
         Top             =   4200
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete record"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   4
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Browse Records"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         TabIndex        =   3
         Top             =   2280
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "New Complaint"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         Picture         =   "Form2.frx":4A745
         TabIndex        =   2
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Judiciary Management System"
         BeginProperty Font 
            Name            =   "Rockwell Condensed"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Left            =   1560
         TabIndex        =   1
         Top             =   120
         Width           =   8535
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Show
Form2.Hide
End Sub

Private Sub Command2_Click()
Form4.Show
Form2.Hide
End Sub

Private Sub Command3_Click()
Form4.Show
Form2.Hide
End Sub

Private Sub Command4_Click()
Form4.Show
Form2.Hide
End Sub

Private Sub Command5_Click()
Form6.Show
Form2.Hide
End Sub

Private Sub Command6_Click()
Form5.Show
Form2.Hide
End Sub

Private Sub Command8_Click()
End
End Sub
