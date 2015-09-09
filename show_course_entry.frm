VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form18 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course-Info"
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19005
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   ScaleHeight     =   697.949
   ScaleMode       =   0  'User
   ScaleWidth      =   24273.52
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000B&
      Caption         =   "Frame4"
      Height          =   3255
      Left            =   1320
      TabIndex        =   7
      Top             =   1320
      Width           =   12375
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
         Left            =   9000
         TabIndex        =   21
         Top             =   2520
         Width           =   1455
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
         TabIndex        =   20
         Top             =   2520
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
         Left            =   4200
         TabIndex        =   19
         Top             =   2520
         Width           =   1335
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
         Height          =   495
         Left            =   1920
         TabIndex        =   18
         Top             =   2520
         Width           =   1335
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
         Left            =   8640
         TabIndex        =   17
         Top             =   1680
         Width           =   1455
      End
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
         Left            =   2280
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
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
         Left            =   8640
         TabIndex        =   13
         Top             =   1080
         Width           =   3375
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
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   1080
         Width           =   3495
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
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Total Semester :"
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
         Left            =   6600
         TabIndex        =   16
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Duration(yr) :"
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
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Department :"
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
         Left            =   6600
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Course Name :"
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
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Course ID :"
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
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Frame1"
      Height          =   5175
      Left            =   2160
      TabIndex        =   1
      Top             =   4800
      Width           =   10455
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7011
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Frame Frame3 
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6120
         TabIndex        =   5
         Top             =   240
         Width           =   3135
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
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton Command12 
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
         Left            =   4560
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   4095
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
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   3375
         End
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Course Information :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
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
  cn.Execute "insert into course_master(course_id,course_name,dpt_name,course_duration,course_tsem) values('" + Text5.Text + "','" + Text1.Text + "','" + Combo1.Text + "','" + Combo4.Text + "','" + Combo5.Text + "');"
  MsgBox "Saved..."
  Call clr
  loadlist
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command11_Click()
  End
End Sub



Private Sub Command12_Click()
  loadlist
End Sub

Private Sub Command13_Click()
  Call clr
  loadlist
End Sub

Private Sub Command2_Click()
Dim cmd As String
On Error GoTo SaveErr
  Dim sql As String
  Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
  cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
   cn.Execute "update course_master set course_id='" + Text5.Text + "',course_name='" + Text1.Text + "'" + ",dpt_name='" + Combo1.Text + "'" + ",course_duration='" + Combo4.Text + "'" + ",course_tsem='" + Combo5.Text + "' where course_id='" + Text5.Text + "'"
  MsgBox "Saved............."
  Call clr
  loadlist
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command3_Click()
On Error GoTo SaveErr
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
  dcom.CommandText = "delete from course_master where course_id='" + Text5.Text + "'"
  dcom.CommandType = adCmdText
  Set dcom.ActiveConnection = cn
  dcom.Execute
  Call clr
  loadlist
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command4_Click()
  Call clr
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
   Form4.Show
   Unload Me
End Sub

Private Sub Command9_Click()
  Form1.Show
  Unload Me
End Sub

Private Sub DataGrid1_OnAddNew()
  Call clr
End Sub

Private Sub Form_Load()
  Command1.Enabled = False
   Command2.Enabled = False
 Command3.Enabled = False

Form7.WindowState = 2
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
  sql = "select dpt_name from dpt_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   Combo1.AddItem rs("dpt_name")
   Combo2.AddItem rs("dpt_name")
     rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    
    For i = 1 To 6
      Combo3.AddItem i
      Combo4.AddItem i
    Next i
    For j = 1 To 12
    Combo5.AddItem j
    Next j
    
    loadlist
End Sub
Public Sub clr()
Text5.Text = ""
Text1.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo2.Text = ""
Combo5.Text = ""
loadlist
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
  sql = "select * from course_master where course_id='" + List1.Text + "'"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
   Text5.Text = rs("course_id")
  Text1.Text = rs("course_name")
  Combo1.Text = rs("dpt_name")
  Text3.Text = rs("course_duration")
  Text4.Text = rs("course_tsem")
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
    sSQL = "select * from course_master where dpt_name like '" & Combo2.Text & "%' and course_duration like '" & Combo3.Text & "%'"
    
   Set rs = New ADODB.Recordset
   With rs
   .Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
   
     Set LV = ListView1
    LV.ListItems.Clear
LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "course ID", 1300, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Course name", 1800, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Department name", 3500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "course Duration", 1900, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Total sem", 1500, lvwColumnLeft
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
   '     rec = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("course_id") & "")
              li.ListSubItems.Add , , rs("course_name") & ""
              li.ListSubItems.Add , , rs("dpt_name") & ""
              li.ListSubItems.Add , , rs("course_duration") & ""
              li.ListSubItems.Add , , rs("course_tsem") & ""
            rs.MoveNext
        Loop
        
    End If
    
    rs.Close
    End With
    
    cn.Close
    
End Sub



Private Sub ListView1_DblClick()
 Text5.Text = ListView1.SelectedItem.Text
    Text1.Text = ListView1.SelectedItem.SubItems(1)
     Combo1.Text = ListView1.SelectedItem.SubItems(2)
      Combo4.Text = ListView1.SelectedItem.SubItems(3)
       Combo5.Text = ListView1.SelectedItem.SubItems(4)
End Sub
