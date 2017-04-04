VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   9720
   ClientLeft      =   3570
   ClientTop       =   255
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Maiandra GD"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   9720
   ScaleWidth      =   13530
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
      Left            =   0
      Picture         =   "Form4.frx":0000
      ScaleHeight     =   16200
      ScaleWidth      =   28800
      TabIndex        =   0
      Top             =   0
      Width           =   28860
      Begin VB.CommandButton Command5 
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   7560
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "Update  Record"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   7560
         Width           =   2535
      End
      Begin VB.TextBox Text15 
         Height          =   405
         Left            =   2400
         TabIndex        =   19
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Next>"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6840
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FFFF&
         Caption         =   "<Previous"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6840
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF00&
         Caption         =   "<Back to start"
         Height          =   495
         Left            =   240
         MaskColor       =   &H80000007&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   8760
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   2415
         Left            =   8280
         ScaleHeight     =   2355
         ScaleWidth      =   2955
         TabIndex        =   15
         Top             =   5520
         Width           =   3015
      End
      Begin VB.TextBox Text14 
         Height          =   405
         Left            =   7920
         TabIndex        =   14
         Text            =   "dd-mm-yyyy"
         Top             =   4920
         Width           =   3495
      End
      Begin VB.TextBox Text13 
         Height          =   405
         Left            =   7920
         TabIndex        =   13
         Top             =   4440
         Width           =   3495
      End
      Begin VB.TextBox Text12 
         Height          =   1365
         Left            =   7920
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox Text11 
         Height          =   405
         Left            =   7920
         TabIndex        =   11
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox Text10 
         Height          =   405
         Left            =   7920
         TabIndex        =   10
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox Text9 
         Height          =   405
         Left            =   7920
         TabIndex        =   9
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   7920
         TabIndex        =   8
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   2400
         TabIndex        =   7
         Top             =   5760
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   2400
         TabIndex        =   6
         Top             =   5280
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         Height          =   1485
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   2400
         TabIndex        =   4
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2400
         TabIndex        =   3
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2400
         TabIndex        =   1
         Top             =   1560
         Width           =   3255
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   11880
         Top             =   6720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Left            =   6120
         TabIndex        =   36
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FIR ID"
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
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label15 
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
         Left            =   6000
         TabIndex        =   34
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         BorderWidth     =   5
         X1              =   5880
         X2              =   5880
         Y1              =   1440
         Y2              =   6840
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Case Status"
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
         Left            =   6000
         TabIndex        =   33
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label13 
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
         Left            =   6000
         TabIndex        =   32
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone no."
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
         Left            =   6000
         TabIndex        =   31
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label11 
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
         Left            =   6000
         TabIndex        =   30
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label10 
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
         Left            =   6000
         TabIndex        =   29
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Left            =   6000
         TabIndex        =   28
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label8 
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
         Left            =   240
         TabIndex        =   27
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Crime type "
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
         Left            =   240
         TabIndex        =   26
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   25
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone no."
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
         Left            =   240
         TabIndex        =   24
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   23
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Complaint Records"
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
         Left            =   2520
         TabIndex        =   20
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlstr As String
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset



Private Sub Command1_Click()
Form4.Hide
Form2.Show

End Sub

Private Sub Command2_Click()
If Not rs.BOF Then
rs.MovePrevious

End If
GetText
End Sub

Private Sub Command3_Click()
If Not rs.EOF Then
rs.MoveNext
End If
GetText

End Sub

Private Sub Command4_Click()
rs.Fields("Aname") = Text1.Text
 rs.Fields("Agender") = Text2.Text
 rs.Fields("Aage") = Text3.Text
 rs.Fields("Anumber") = Text4.Text
 rs.Fields("Aaddress") = Text5.Text
 rs.Fields("CrimeType") = Text6.Text
rs.Fields("Description") = Text7.Text
 rs.Fields("Bname") = Text8.Text
 rs.Fields("Bgender") = Text9.Text
 rs.Fields("Bage") = Text10.Text
 rs.Fields("Bnumber") = Text11.Text
 rs.Fields("Baddress") = Text12.Text
 rs.Fields("status") = Text13.Text
 rs.Fields("Date") = Text14.Text
 rs.Fields("FIR_ID") = Text15.Text
rs.Update
MsgBox "Record successfully updated", vbInformation, "Update Record"


End Sub

Private Sub Command5_Click()
Confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deletion Confirmation")
If Confirm = vbYes Then
rs.Delete
MsgBox "Record Deleted!", , "Message"
Else
MsgBox "Record Not Deleted!", , "Message"
End If

End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\win8.1\Desktop\judiciary\Database2.mdb;Persist Security Info=False"
conn.Open
sqlstr = "SELECT * from Table1"
rs.Open (sqlstr), conn, adOpenDynamic, adLockOptimistic
GetText
End Sub


Private Sub GetText()
If rs.EOF = True Or rs.BOF = True Then Exit Sub
Text1.Text = rs.Fields("Aname")
Text2.Text = rs.Fields("Agender")
Text3.Text = rs.Fields("Aage")
Text4.Text = rs.Fields("Anumber")
Text5.Text = rs.Fields("Aaddress")
Text6.Text = rs.Fields("CrimeType")
Text7.Text = rs.Fields("Description")
Text8.Text = rs.Fields("Bname")
Text9.Text = rs.Fields("Bgender")
Text10.Text = rs.Fields("Bage")
Text11.Text = rs.Fields("Bnumber")
Text12.Text = rs.Fields("Baddress")
Text13.Text = rs.Fields("status")
Text14.Text = rs.Fields("Date")
Text15.Text = rs.Fields("FIR_ID")
Picture2.Picture = LoadPicture(rs.Fields("Picture"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close
conn.Close
Set conn = Nothing
End Sub
