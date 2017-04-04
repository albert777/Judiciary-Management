VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13005
   BeginProperty Font 
      Name            =   "Maiandra GD"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   9165
   ScaleWidth      =   13005
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
      Height          =   16260
      Left            =   -1440
      Picture         =   "Form5.frx":0000
      ScaleHeight     =   16200
      ScaleWidth      =   28800
      TabIndex        =   0
      Top             =   0
      Width           =   28860
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "Delete Record"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   7200
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Update Record"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   7200
         Width           =   2415
      End
      Begin VB.TextBox Text15 
         Height          =   405
         Left            =   10920
         TabIndex        =   35
         Text            =   "dd-mm-yyyy"
         Top             =   5640
         Width           =   2895
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   2295
         Left            =   11040
         ScaleHeight     =   2235
         ScaleWidth      =   2115
         TabIndex        =   33
         Top             =   6240
         Width           =   2175
      End
      Begin VB.TextBox Text14 
         Height          =   405
         Left            =   10920
         TabIndex        =   31
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox Text13 
         Height          =   885
         Left            =   10920
         TabIndex        =   30
         Top             =   3960
         Width           =   2895
      End
      Begin VB.TextBox Text12 
         Height          =   315
         Left            =   10920
         TabIndex        =   29
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox Text11 
         Height          =   315
         Left            =   10920
         TabIndex        =   28
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox Text10 
         Height          =   315
         Left            =   10920
         TabIndex        =   27
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text9 
         Height          =   315
         Left            =   10920
         TabIndex        =   26
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   4440
         TabIndex        =   25
         Top             =   5520
         Width           =   3015
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   4440
         TabIndex        =   24
         Top             =   5040
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         Height          =   885
         Left            =   4440
         TabIndex        =   23
         Top             =   3960
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   4440
         TabIndex        =   22
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   4440
         TabIndex        =   21
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   4440
         TabIndex        =   20
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   4440
         TabIndex        =   19
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Back to start"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         MaskColor       =   &H80000012&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8280
         UseMaskColor    =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   495
         Left            =   9840
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   5880
         TabIndex        =   3
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   9240
         TabIndex        =   34
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Picture"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   9240
         TabIndex        =   32
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Case status"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   9120
         TabIndex        =   18
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   9000
         TabIndex        =   17
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   9120
         TabIndex        =   16
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   9120
         TabIndex        =   15
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   9120
         TabIndex        =   14
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Accussed"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   9120
         TabIndex        =   13
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         BorderWidth     =   5
         X1              =   8880
         X2              =   8880
         Y1              =   1920
         Y2              =   7080
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Crime Type"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Applicant"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Enter FIR ID to be searched for:"
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Search Record"
         BeginProperty Font 
            Name            =   "Rockwell Condensed"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   975
         Left            =   5160
         TabIndex        =   1
         Top             =   120
         Width           =   6375
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlstr2, sqlsrch As String
Dim conn2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset

Private Sub Command1_Click()
searchvar = Text1.Text
sqlsrch = "SELECT * FROM Table1 WHERE FIR_ID =" & "'" & searchvar & "'"

rs2.Close
rs2.Open (sqlsrch), conn2, adOpenDynamic, adLockOptimistic
If rs2.Fields(0) <> "" Then
Text2.Text = rs2.Fields("Aname")
Text3.Text = rs2.Fields("Agender")
Text4.Text = rs2.Fields("Aage")
Text5.Text = rs2.Fields("Anumber")
Text6.Text = rs2.Fields("Aaddress")
Text7.Text = rs2.Fields("CrimeType")
Text8.Text = rs2.Fields("Description")
Text9.Text = rs2.Fields("Bname")
Text10.Text = rs2.Fields("Bgender")
Text11.Text = rs2.Fields("Bage")
Text12.Text = rs2.Fields("Bnumber")
Text13.Text = rs2.Fields("Baddress")
Text14.Text = rs2.Fields("status")
Text15.Text = rs2.Fields("Date")

Picture2.Picture = LoadPicture(rs2.Fields("Picture"))
Else
MsgBox ("No records found!")
End If
End Sub


Private Sub Command2_Click()
Form2.Show
Form5.Hide
End Sub

Private Sub Command3_Click()
rs2.Fields("Aname") = Text2.Text
rs2.Fields("Agender") = Text3.Text
rs2.Fields("Aage") = Text4.Text
rs2.Fields("Anumber") = Text5.Text
rs2.Fields("Aaddress") = Text6.Text
rs2.Fields("CrimeType") = Text7.Text
rs2.Fields("Description") = Text8.Text
rs2.Fields("Bname") = Text9.Text
rs2.Fields("Bgender") = Text10.Text
rs2.Fields("Bage") = Text11.Text
rs2.Fields("Bnumber") = Text12.Text
rs2.Fields("Baddress") = Text13.Text
rs2.Fields("status") = Text14.Text
rs2.Fields("Date") = Text15.Text
rs2.Update
MsgBox "Records Updated", vbInformation

End Sub

Private Sub Command4_Click()
Confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deletion Confirmation")
If Confirm = vbYes Then
rs2.Delete
MsgBox "Record Deleted!", , "Message"
Else
MsgBox "Record Not Deleted!", , "Message"
End If
End Sub

Private Sub Form_Load()
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\win8.1\Desktop\judiciary\Database2.mdb;Persist Security Info=False"
conn2.Open
sqlstr2 = "SELECT * from Table1"
rs2.Open (sqlstr2), conn2, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs2.Close
conn2.Close
Set conn2 = Nothing
End Sub

