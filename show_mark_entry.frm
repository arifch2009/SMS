VERSION 5.00
Begin VB.Form Form19 
   BackColor       =   &H8000000B&
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19005
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   Picture         =   "show_mark_entry.frx":0000
   ScaleHeight     =   697.949
   ScaleMode       =   0  'User
   ScaleWidth      =   23537.96
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Caption         =   "Frame3"
      Height          =   6015
      Left            =   1200
      TabIndex        =   27
      Top             =   4560
      Width           =   17775
      Begin VB.PictureBox ListView1 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   240
         ScaleHeight     =   4635
         ScaleWidth      =   16875
         TabIndex        =   44
         Top             =   1200
         Width           =   16935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         TabIndex        =   40
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frame9 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   13320
         TabIndex        =   38
         Top             =   240
         Width           =   2055
         Begin VB.ComboBox Combo16 
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
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Subject Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   11160
         TabIndex        =   36
         Top             =   240
         Width           =   2175
         Begin VB.ComboBox Combo15 
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
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Semester"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9480
         TabIndex        =   34
         Top             =   240
         Width           =   1695
         Begin VB.ComboBox Combo14 
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
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Exam Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5400
         TabIndex        =   32
         Top             =   240
         Width           =   2175
         Begin VB.ComboBox Combo13 
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
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Exam ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3000
         TabIndex        =   30
         Top             =   240
         Width           =   2415
         Begin VB.ComboBox Combo12 
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
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Student ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   2895
         Begin VB.ComboBox Combo11 
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
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   2415
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "Exam Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   15615
      Begin VB.ComboBox Combo4 
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
         Left            =   12960
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   10080
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   5760
         TabIndex        =   6
         Top             =   360
         Width           =   2415
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
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Category :"
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
         Left            =   11400
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Semester :"
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
         Left            =   8520
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Student ID :"
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
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Exam Roll :"
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Mark Entry"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1320
      TabIndex        =   1
      Top             =   2280
      Width           =   15375
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
         Height          =   330
         Left            =   12000
         TabIndex        =   43
         Top             =   435
         Width           =   2655
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Refresh"
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
         Left            =   11280
         TabIndex        =   41
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   26
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   25
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   24
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1680
         TabIndex        =   23
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox Combo10 
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
         Left            =   12720
         TabIndex        =   22
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox Combo9 
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
         Left            =   6120
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox Combo8 
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
         Left            =   9960
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox Combo7 
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
         Left            =   2160
         TabIndex        =   16
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo6 
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
         Left            =   6240
         TabIndex        =   14
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   2160
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Student Name:"
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
         Left            =   9960
         TabIndex        =   42
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Status :"
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
         Left            =   11400
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Pass Mark :"
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
         Left            =   4320
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Mark Obtained :"
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
         Left            =   7800
         TabIndex        =   17
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Total Mark :"
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
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Subject Name :"
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
         Left            =   4320
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Subject Code :"
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
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Student Mark Entry :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo5_Change()
loadlist1
End Sub









Private Sub subname()
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
  sql = "select sub_name from sub_master where sub_id='" + Combo5.Text + "';"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   
   Combo6.AddItem rs("sub_name")
   
   rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub

Private Sub Command1_Click()
 On Error GoTo SaveErr
    Dim sSQL As String
    
    sSQL = "select * from mark_master"
        Dim cmd As String
    cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=trial"
    Set dbConn = New ADODB.Connection
    Set dbRec = New ADODB.Recordset
    
    dbConn.ConnectionString = cmd
    dbConn.Open
    
    dbRec.Open sSQL, dbConn, adOpenDynamic, adLockOptimistic
        
    With dbRec
        .AddNew
        .Fields("stu_id") = Combo2.Text
        .Fields("exam_id") = Combo1.Text
        .Fields("exam_sem") = Combo3.Text
        
        .Fields("exam_category") = Combo4.Text
        
        .Fields("sub_code") = Combo5.Text
        
        .Fields("total_mark") = Combo7.Text
        .Fields("pass_mark") = Combo8.Text
        .Fields("mark_obtained") = Combo9.Text
        .Fields("status") = Combo10.Text
        .Update
    End With
    
    dbRec.Close
    dbConn.Close
        
    Set dbConn = Nothing
    Set dbRec = Nothing
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command4_Click()
 Call clr
End Sub
Public Sub clr()
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
Combo6.Text = ""
Combo7.Text = ""
Combo8.Text = ""
Combo9.Text = ""
Combo10.Text = ""
Combo11.Text = ""
Combo12.Text = ""
Combo13.Text = ""
Combo14.Text = ""
Combo15.Text = ""
Combo16.Text = ""
Text1.Text = ""
loadlist

End Sub

Private Sub Command5_Click()
loadlist1
End Sub

Private Sub Command6_Click()

  Call sname
  Call com5
End Sub

Private Sub Form_Load()
  Command1.Enabled = False
   Command2.Enabled = False
 Command3.Enabled = False
  Form19.WindowState = 2
  Call roll1
  Call id1
  Dim j As Integer
  For j = 1 To 12
  Combo3.AddItem j
  Combo14.AddItem j
  Next j
   Dim i As Integer
    For i = 1 To 150
    Combo7.AddItem i
    Combo8.AddItem i
    Combo9.AddItem i
    Next i
    Combo10.AddItem "Pass"
    Combo10.AddItem "Fail"
    Combo10.AddItem "---"
    Combo16.AddItem "Pass"
    Combo16.AddItem "Fail"
    Combo16.AddItem "---"
  Call com4
  Call com5
  loadlist
  
End Sub
Public Sub sname()
  On Error GoTo SaveErr
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
  sql = "select * from stu_personal where stu_id='" + Combo2.Text + "'"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 
  Text1.Text = rs("stu_name")
  .Close
  End With
  Set rs = Nothing
  cn.Close
  Set cn = Nothing
   
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Public Sub com5()
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
  sql = "select distinct sub_master.sub_code,sub_master.sub_name from sub_master,stu_aca where stu_aca.dpt_name=sub_master.dpt_name and stu_aca.stu_id like '" + Combo2.Text + "%' and sub_master.sub_sem like '" + Combo3.Text + "%';"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   Combo5.AddItem rs("sub_code")
   Combo6.AddItem rs("sub_name")
   Combo15.AddItem rs("sub_code")
   rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
   
    End Sub
Public Sub com4()
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
  sql = "select distinct exam_category from exam_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   Combo4.AddItem rs("exam_category")
    Combo13.AddItem rs("exam_category")
   rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub
Public Sub id1()
On Error GoTo SaveErr
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
  sql = "select stu_id from stu_personal"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   Combo2.AddItem rs("stu_id")
   Combo11.AddItem rs("stu_id")
   
   rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
SaveErr:
    MsgBox Err.Description

End Sub
Public Sub roll1()
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
  sql = "select distinct exam_id from exam_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   Combo1.AddItem rs("exam_id")
   Combo12.AddItem rs("exam_id")
   rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
End Sub
Public Sub loadlist()
On Error Resume Next
    Dim li As ListItem
    Dim LV As ListView
     Dim cmd As String
  Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
  cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=trial"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
 '   Dim rec As Integer
    Dim sSQL  As String
    sSQL = "select * from mark_master"
    
   Set rs = New ADODB.Recordset
   With rs
   .Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
   
     Set LV = ListView1
    LV.ListItems.Clear
LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "Student code", 1700, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Exam Code", 1400, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Exam category", 1700, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Semester", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Code", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Total Mark", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Pass Mark", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Mark Obtained", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Status", 2000, lvwColumnLeft
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
   '     rec = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("stu_id") & "")
              li.ListSubItems.Add , , rs("exam_id") & ""
              li.ListSubItems.Add , , rs("exam_category") & ""
              li.ListSubItems.Add , , rs("exam_sem") & ""
              li.ListSubItems.Add , , rs("sub_code") & ""
              li.ListSubItems.Add , , rs("total_mark") & ""
              li.ListSubItems.Add , , rs("pass_mark") & ""
              li.ListSubItems.Add , , rs("mark_obtained") & ""
              li.ListSubItems.Add , , rs("status") & ""
            rs.MoveNext
        Loop
        
    End If
    
    rs.Close
    End With
    
    cn.Close
    
End Sub
Public Sub loadlist1()
On Error Resume Next
    Dim li As ListItem
    Dim LV As ListView
     Dim cmd As String
  Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
  cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=trial"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
 '   Dim rec As Integer
    Dim sSQL  As String
    sSQL = "select * from mark_master where exam_id like '" + Combo12.Text + "%' and exam_category like '" + Combo13.Text + "%' and stu_id like '" + Combo11.Text + "%' and exam_sem like '" + Combo14.Text + "%' and sub_code like '" + Combo15.Text + "%' and status like '" + Combo16.Text + "%'"
    
   Set rs = New ADODB.Recordset
   With rs
   .Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
   
     Set LV = ListView1
    LV.ListItems.Clear
LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "Student code", 1700, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Exam Code", 1400, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Exam category", 1700, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Semester", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Code", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Total Mark", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Pass Mark", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Mark Obtained", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Status", 2000, lvwColumnLeft
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
   '     rec = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("stu_id") & "")
              li.ListSubItems.Add , , rs("exam_id") & ""
              li.ListSubItems.Add , , rs("exam_category") & ""
              li.ListSubItems.Add , , rs("exam_sem") & ""
              li.ListSubItems.Add , , rs("sub_code") & ""
              li.ListSubItems.Add , , rs("total_mark") & ""
              li.ListSubItems.Add , , rs("pass_mark") & ""
              li.ListSubItems.Add , , rs("mark_obtained") & ""
              li.ListSubItems.Add , , rs("status") & ""
              
            rs.MoveNext
        Loop
        
    End If
    
    rs.Close
    End With
    
    cn.Close
End Sub
Private Sub ListView1_DblClick()

     Combo2.Text = ListView1.SelectedItem.Text
    Combo1.Text = ListView1.SelectedItem.SubItems(1)
     Combo4.Text = ListView1.SelectedItem.SubItems(2)
      Combo3.Text = ListView1.SelectedItem.SubItems(3)
       Combo5.Text = ListView1.SelectedItem.SubItems(4)
      Combo7.Text = ListView1.SelectedItem.SubItems(5)
      Combo8.Text = ListView1.SelectedItem.SubItems(6)
      Combo9.Text = ListView1.SelectedItem.SubItems(7)
      Combo10.Text = ListView1.SelectedItem.SubItems(8)
End Sub
