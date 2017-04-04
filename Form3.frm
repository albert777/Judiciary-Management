VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9000
   ClientLeft      =   4155
   ClientTop       =   555
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Maiandra GD"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   9000
   ScaleWidth      =   12030
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
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   16200
      ScaleWidth      =   28800
      TabIndex        =   0
      Top             =   120
      Width           =   28860
      Begin VB.TextBox Text15 
         Height          =   405
         Left            =   2640
         TabIndex        =   35
         Top             =   1080
         Width           =   2775
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7440
         Top             =   7680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Browse picture"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6720
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   2055
         Left            =   9000
         ScaleHeight     =   1995
         ScaleWidth      =   2355
         TabIndex        =   32
         Top             =   6240
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000D&
         Caption         =   "<Back"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6720
         UseMaskColor    =   -1  'True
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   7560
         Width           =   2415
      End
      Begin VB.TextBox Text14 
         Height          =   405
         Left            =   8880
         TabIndex        =   29
         Text            =   "dd-mm-yyyy"
         Top             =   5520
         Width           =   2775
      End
      Begin VB.TextBox Text13 
         Height          =   405
         Left            =   8880
         TabIndex        =   27
         Top             =   5040
         Width           =   2775
      End
      Begin VB.TextBox Text12 
         Height          =   1245
         Left            =   8880
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox Text11 
         Height          =   405
         Left            =   8880
         TabIndex        =   25
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox Text10 
         Height          =   405
         Left            =   8880
         TabIndex        =   24
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox Text9 
         Height          =   405
         Left            =   8880
         TabIndex        =   23
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   8880
         TabIndex        =   22
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   2640
         TabIndex        =   21
         Top             =   5520
         Width           =   2775
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   2640
         TabIndex        =   20
         Top             =   5040
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   1365
         Left            =   2640
         TabIndex        =   19
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2640
         TabIndex        =   17
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2640
         TabIndex        =   16
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2640
         TabIndex        =   15
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "FIR ID"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   6240
         TabIndex        =   28
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000B&
         BorderWidth     =   5
         X1              =   5760
         X2              =   5760
         Y1              =   1440
         Y2              =   6000
      End
      Begin VB.Label Label14 
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
         Left            =   6240
         TabIndex        =   14
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Addresss"
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
         Left            =   6240
         TabIndex        =   13
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label12 
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
         Left            =   6240
         TabIndex        =   12
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label11 
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
         Left            =   6240
         TabIndex        =   11
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label10 
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
         Left            =   6240
         TabIndex        =   10
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of accussed"
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
         TabIndex        =   9
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   8
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Crime type"
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
         Left            =   120
         TabIndex        =   7
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone number"
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
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of applicant"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Register new complaint"
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
         Height          =   855
         Left            =   2760
         TabIndex        =   1
         Top             =   120
         Width           =   5895
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlstr1 As String
Dim conn1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset


Private Sub Command1_Click()
rs1.AddNew
rs1.Fields("Aname") = Text1.Text
rs1.Fields("Agender") = Text2.Text
rs1.Fields("Aage") = Text3.Text
rs1.Fields("Anumber") = Text4.Text
rs1.Fields("Aaddress") = Text5.Text
rs1.Fields("CrimeType") = Text6.Text
rs1.Fields("Description") = Text7.Text
rs1.Fields("Bname") = Text8.Text
rs1.Fields("Bgender") = Text9.Text
rs1.Fields("Bage") = Text10.Text
rs1.Fields("Bnumber") = Text11.Text
rs1.Fields("Baddress") = Text12.Text
rs1.Fields("status") = Text13.Text
rs1.Fields("Date") = Text14.Text
rs1.Fields("FIR_ID") = Text15.Text
rs1.Fields("Picture") = CommonDialog1.FileName
rs1.Update
y = MsgBox("Records Successfully Registered")
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = "dd-mm-yyyy"
Text15.Text = ""
Picture2.Picture = LoadPicture(Empty)



End Sub

Private Sub Command2_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Command3_Click()
CommonDialog1.ShowOpen
Picture2.Picture = LoadPicture(CommonDialog1.FileName)

End Sub

Private Sub Form_Load()
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\win8.1\Desktop\judiciary\Database2.mdb;Persist Security Info=False"
conn1.Open
sqlstr1 = "SELECT * from Table1"
rs1.Open (sqlstr1), conn1, adOpenDynamic, adLockOptimistic
End Sub



Private Sub Form_Unload(Cancel As Integer)
rs1.Close
conn1.Close
Set conn1 = Nothing
End Sub
