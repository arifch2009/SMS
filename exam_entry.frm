VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form11 
   BackColor       =   &H8000000B&
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19005
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   ScaleHeight     =   697.949
   ScaleMode       =   0  'User
   ScaleWidth      =   23537.96
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "Frame2"
      Height          =   6255
      Left            =   1800
      TabIndex        =   15
      Top             =   4320
      Width           =   13215
      Begin MSComctlLib.ListView ListView1 
         Height          =   4695
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   8281
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
      Begin VB.Frame Frame4 
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
         Height          =   735
         Left            =   7080
         TabIndex        =   19
         Top             =   240
         Width           =   3495
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
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.CommandButton Command5 
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
         Left            =   4800
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Exam Roll"
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
         TabIndex        =   16
         Top             =   240
         Width           =   3735
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
            Left            =   360
            TabIndex        =   17
            Top             =   240
            Width           =   2895
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   13095
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
         Height          =   495
         Left            =   1920
         TabIndex        =   22
         Top             =   360
         Width           =   2415
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
         Left            =   7920
         TabIndex        =   14
         Top             =   1560
         Width           =   1695
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
         Height          =   465
         Left            =   5520
         TabIndex        =   13
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
         Left            =   3240
         TabIndex        =   12
         Top             =   1560
         Width           =   1575
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
         Left            =   720
         TabIndex        =   11
         Top             =   1590
         Width           =   1815
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
         Left            =   2040
         TabIndex        =   10
         Top             =   960
         Width           =   2895
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
         Left            =   6240
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   11040
         TabIndex        =   6
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   85327873
         CurrentDate     =   40488
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
         Height          =   375
         Left            =   7080
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Exam Name :"
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
         TabIndex        =   9
         Top             =   960
         Width           =   1815
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
         Height          =   495
         Left            =   4680
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Commence on :"
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
         Left            =   9000
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Exam Session :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   3
         Top             =   960
         Width           =   1935
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
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Exam Schedule Entry :"
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
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  

  On Error GoTo SaveErr
    Dim sSQL As String
    
    sSQL = "select * from exam_master"
        Dim cmd As String
    cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=trial"
    Set dbConn = New ADODB.Connection
    Set dbRec = New ADODB.Recordset
    
    dbConn.ConnectionString = cmd
    dbConn.Open
    
    dbRec.Open sSQL, dbConn, adOpenDynamic, adLockOptimistic
        
    With dbRec
        .AddNew
        .Fields("exam_id") = Text1.Text
        .Fields("exam_name") = Combo2.Text
        
        .Fields("exam_category") = Combo1.Text
        
        .Fields("exam_schedule") = DTPicker1.Value
        
        .Fields("exam_session") = Text2.Text
        .Update
    End With
    
    dbRec.Close
    dbConn.Close
        
    Set dbConn = Nothing
    Set dbRec = Nothing
    loadlist
    Call clr
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command2_Click()
  On Error GoTo SaveErr
    Dim sSQL As String
    
    sSQL = "select * from exam_master where exam_id='" + Text1.Text + "' and exam_category='" + Combo1.Text + "'"
        Dim cmd As String
    cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=trial"
    Set dbConn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    dbConn.ConnectionString = cmd
    dbConn.Open
    
    rs.Open sSQL, dbConn, adOpenDynamic, adLockOptimistic
        
   
        rs("exam_id") = Text1.Text
        rs("exam_name") = Combo2.Text
        
        rs("exam_category") = Combo1.Text
        
        rs("exam_schedule") = DTPicker1.Value
        
        rs("exam_session") = Text2.Text
        rs.Update
       
    rs.Close
    dbConn.Close
        
    Set dbConn = Nothing
    Set rs = Nothing
    loadlist
Exit Sub
  Call clr
SaveErr:
    MsgBox Err.Description
    
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
  dcom.CommandText = "delete from exam_master where exam_id='" + Text1.Text + "' and exam_category='" + Combo1.Text + "'"
  dcom.CommandType = adCmdText
  Set dcom.ActiveConnection = cn
  dcom.Execute
  Call clr
 
  cn.Close
  Set cn = Nothing
  Call clr
End Sub

Private Sub Command4_Click()
  Call clr
  loadlist
End Sub

Private Sub Command5_Click()
  loadlist
  Call clr
End Sub

Private Sub Form_Load()
Form11.WindowState = 2
loadlist
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
  sql = "select course_name from course_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   Combo2.AddItem rs("course_name") + "(Even Sem.)"
    Combo2.AddItem rs("course_name") + "(Odd Sem.)"
   rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    
  Combo1.AddItem "1st Internal"
  Combo1.AddItem "2nd Internal"
  Combo1.AddItem "External"
 Call com3
 Call com4
End Sub
Public Sub com4()
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
  sql = "select distinct exam_category from exam_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
    Combo4.AddItem rs("exam_category")
   rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub
Public Sub com3()
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
  sql = "select distinct exam_id from exam_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
    Combo3.AddItem rs("exam_id")
   rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
  Exit Sub
SaveErr:
    MsgBox Err.Description
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
    sSQL = "select * from exam_master where exam_id like '" & Combo3.Text & "%' and exam_category like '" & Combo4.Text & "%'"
    
   Set rs = New ADODB.Recordset
   With rs
   .Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
   
     Set LV = ListView1
    LV.ListItems.Clear
LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "Exam code", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Exam Name", 4000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Exam Category", 3000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Exam Session", 1700, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Exam Schedule", 2000, lvwColumnLeft
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
   '     rec = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("exam_id") & "")
              li.ListSubItems.Add , , rs("exam_name") & ""
              li.ListSubItems.Add , , rs("exam_category") & ""
              li.ListSubItems.Add , , rs("exam_session") & ""
              li.ListSubItems.Add , , rs("exam_schedule") & ""
            rs.MoveNext
        Loop
        
    End If
    
    rs.Close
    End With
    
    cn.Close
    
End Sub
Public Sub clr()
 Text1.Text = ""
 Text2.Text = ""
 Combo1.Text = ""
 Combo2.Text = ""
 DTPicker1.Value = Value
 loadlist
End Sub
Private Sub ListView1_DblClick()

     Text1.Text = ListView1.SelectedItem.Text
    Combo2.Text = ListView1.SelectedItem.SubItems(1)
     Combo1.Text = ListView1.SelectedItem.SubItems(2)
      Text2.Text = ListView1.SelectedItem.SubItems(3)
     DTPicker1.Value = ListView1.SelectedItem.SubItems(4)
End Sub

