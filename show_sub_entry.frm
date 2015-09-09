VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form21 
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
      Height          =   7335
      Left            =   840
      TabIndex        =   15
      Top             =   3600
      Width           =   13095
      Begin MSComctlLib.ListView ListView1 
         Height          =   5775
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   10186
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
      Begin VB.CommandButton Command5 
         Caption         =   "Refresh"
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
         Left            =   5160
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   735
         Left            =   7440
         TabIndex        =   18
         Top             =   240
         Width           =   3255
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
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   735
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   4335
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
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   3735
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   12495
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
         Left            =   11040
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
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
         Left            =   8400
         TabIndex        =   10
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text3 
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
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   3615
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
         Left            =   5640
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Sub Category :"
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
         Left            =   6000
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Sub_Name :"
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
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1575
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
         Left            =   9360
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
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
         Left            =   3840
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "sub_code   :"
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
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Subject Information :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form21"
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
  cn.Execute "insert into sub_master(sub_code,sub_name,sub_sem,dpt_name,sub_category) values('" + Text1.Text + "','" + Text3.Text + "','" + Combo5.Text + "','" + Combo1.Text + "','" + Combo2.Text + "');"
  MsgBox "Saved..."
  Call clr
  loadlist
  
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command2_Click()
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
   cn.Execute "update sub_master set sub_code='" + Text1.Text + "',sub_name='" + Text3.Text + "'" + ",dpt_name='" + Combo1.Text + "'" + ",sub_sem='" + Combo5.Text + "'" + ",sub_category='" + Combo2.Text + "' where sub_code='" + Text1.Text + "'"
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
  dcom.CommandText = "delete from sub_master where sub_code='" + Text1.Text + "'"
  dcom.CommandType = adCmdText
  Set dcom.ActiveConnection = cn
  dcom.Execute
  Call clr
 
  cn.Close
  Set cn = Nothing
  loadlist
  Call clr
  
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command4_Click()
 Call clr
 loadlist
End Sub

Private Sub Command5_Click()
  loadlist
End Sub

Private Sub Form_Load()
Command1.Enabled = False
   Command2.Enabled = False
 Command3.Enabled = False
Form10.WindowState = 2
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
   Combo3.AddItem rs("dpt_name")
   Combo1.AddItem rs("dpt_name")
     rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
Combo2.AddItem "Theory"
Combo2.AddItem "Practical"
Combo2.AddItem "Project"
Combo2.AddItem "Workshop"
For i = 1 To 12
  Combo4.AddItem i
  Combo5.AddItem i
Next i

loadlist



End Sub

Private Sub clr()
Text1.Text = ""
Combo5.Text = ""
Text3.Text = ""
Combo1.Text = ""
Combo2.Text = ""
loadlist
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
    sSQL = "select * from sub_master where dpt_name like '" & Combo3.Text & "%' and sub_sem like '" & Combo4.Text & "%'"
    
   Set rs = New ADODB.Recordset
   With rs
   .Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
   
     Set LV = ListView1
    LV.ListItems.Clear
LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "subject code", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "subject name", 4000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Department name", 3500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Semester", 1200, lvwColumnLeft
    LV.ColumnHeaders.Add , , "subjec Category", 2000, lvwColumnLeft
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
   '     rec = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("sub_code") & "")
              li.ListSubItems.Add , , rs("sub_name") & ""
              li.ListSubItems.Add , , rs("dpt_name") & ""
              li.ListSubItems.Add , , rs("sub_sem") & ""
              li.ListSubItems.Add , , rs("sub_category") & ""
            rs.MoveNext
        Loop
        
    End If
    
    rs.Close
    End With
    
    cn.Close
    

    
End Sub




Private Sub ListView1_DblClick()

     Text1.Text = ListView1.SelectedItem.Text
    Text3.Text = ListView1.SelectedItem.SubItems(1)
     Combo1.Text = ListView1.SelectedItem.SubItems(2)
      Combo5.Text = ListView1.SelectedItem.SubItems(3)
       Combo2.Text = ListView1.SelectedItem.SubItems(4)
End Sub

