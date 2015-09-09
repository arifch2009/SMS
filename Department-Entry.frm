VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department-Entry"
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19005
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   ScaleHeight     =   697.949
   ScaleMode       =   0  'User
   ScaleWidth      =   19005
   Begin VB.CommandButton Command14 
      Caption         =   "Show Department Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      TabIndex        =   35
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   34
      Top             =   9240
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   9480
      TabIndex        =   33
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Show Department IDs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   32
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   31
      Top             =   1440
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3720
      TabIndex        =   29
      Top             =   3000
      Width           =   4215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      TabIndex        =   28
      Top             =   9120
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Change Setting"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   27
      Top             =   8400
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Go To Login Screen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   26
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Course Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   25
      Top             =   7080
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Subject Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   24
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Student Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   23
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exam Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   22
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "School Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14040
      TabIndex        =   21
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   18
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   17
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   9240
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   8160
      Width           =   4335
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   13
      Top             =   7320
      Width           =   4335
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3600
      TabIndex        =   7
      Top             =   4560
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3840
      TabIndex        =   2
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "Department ID :"
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
      Left            =   840
      TabIndex        =   30
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000A&
      Caption         =   "Please Click :"
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
      Left            =   14160
      TabIndex        =   20
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000B&
      Caption         =   "Please Click :"
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
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Height          =   6615
      Index           =   1
      Left            =   13440
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label Label10 
      Caption         =   "Website (If Any) :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   14
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Email-ID(If Any) :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Number of courses :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   10
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Established(Year) :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Address(Department) :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Phone(Department) :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Under the School Of :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Department Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department Information :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
 Dim cmd As String
  Dim sql As String
  Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
  cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
  cn.Execute "insert into dpt_master(dpt_id,dpt_name,school_name,dpt_phone,dpt_add,no_of_course,dpt_stab,dpt_email,dpt_web) values('" + Text2.Text + "','" + Text1.Text + "','" + Combo1.Text + "','" + Text4.Text + "','" + Text5.Text + "','" + Text6.Text + "','" + Text7.Text + "','" + Text8.Text + "','" + Text9.Text + "');"
  MsgBox "Saved..."
  Call clr
End Sub

Private Sub Command11_Click()
  End
End Sub

Private Sub Command12_Click()
  Dim cmd As String
  Dim sql As String
 Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
   cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
  sql = "select dpt_id from dpt_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   List1.AddItem rs("dpt_id")
     rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub

Private Sub Command13_Click()
  Call clr
End Sub

Private Sub Command14_Click()
  Form9.Show
  Unload Me
End Sub

Private Sub Command2_Click()
   Dim cmd As String
  Dim sql As String
  Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
  cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
   cn.Execute "update dpt_master set dpt_id='" + Text2.Text + "',dpt_name='" + Text1.Text + "'" + ",school_name='" + Combo1.Text + "'" + ",dpt_phone='" + Text4.Text + "'" + ",dpt_add='" + Text5.Text + "'" + ",no_of_course='" + Text6.Text + "'" + ",dpt_stab='" + Text7.Text + "'" + ",dpt_email='" + Text8.Text + "'" + ",dpt_web='" + Text9.Text + "'" + " where dpt_id='" + List1.Text + "'"
  MsgBox "Saved............."
  Call clr
End Sub

Private Sub Command3_Click()
   Dim cmd As String
  Dim cn As ADODB.Connection
  Dim dcom As ADODB.Command
  cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
  Set dcom = New ADODB.Command
  dcom.CommandText = "delete from dpt_master where dpt_id='" + List1.Text + "'"
  dcom.CommandType = adCmdText
  Set dcom.ActiveConnection = cn
  dcom.Execute
  Call clr
  List1.RemoveItem (List1.ListIndex)
  MsgBox "Record of username " + List1.Text + " is deleted"
  cn.Close
  Set cn = Nothing
  Call clr
End Sub

Private Sub Command4_Click()
  Form4.Show
  Unload Me
End Sub

Private Sub Command5_Click()
  Form11.Show
  Unload Me
End Sub

Private Sub Command6_Click()
  Form6.Show
  Unload Me
End Sub

Private Sub Command7_Click()
  Form10.Show
  Unload Me
End Sub

Private Sub Command8_Click()
  Form7.Show
  Unload Me
End Sub

Private Sub Command9_Click()
  Form1.Show
  Unload Me
End Sub

Private Sub Form_Load()

  Form5.WindowState = 2
  Dim cmd As String
  Dim sql As String
 Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
   cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
  sql = "select school_name from school_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   Combo1.AddItem rs("school_name")
     rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub

Private Sub List1_Click()
   Dim cmd As String
  Dim sql As String
  Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
  cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
  sql = "select * from dpt_master where dpt_id='" + List1.Text + "'"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
   Text2.Text = rs("dpt_id")
  Text1.Text = rs("dpt_name")
  Combo1.Text = rs("school_name")
  Text4.Text = rs("dpt_phone")
  Text5.Text = rs("dpt_add")
  Text6.Text = rs("dpt_stab")
  Text7.Text = rs("no_of_course")
  Text8.Text = rs("dpt_email")
  Text9.Text = rs("dpt_web")
  .Close
  End With
  Set rs = Nothing
  cn.Close
  Set cn = Nothing
End Sub
Public Sub clr()
  Text2.Text = ""
Text1.Text = ""
Text4.Text = ""
  Text5.Text = ""
  Text6.Text = ""
  Text7.Text = ""
  Text8.Text = ""
  Text9.Text = ""
End Sub
