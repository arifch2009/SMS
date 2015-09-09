VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "School-Entry"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18855
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form4"
   ScaleHeight     =   696.949
   ScaleMode       =   0  'User
   ScaleWidth      =   19005
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Caption         =   "Frame2"
      Height          =   5295
      Left            =   960
      TabIndex        =   24
      Top             =   5160
      Width           =   17895
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
         Height          =   495
         Left            =   5280
         TabIndex        =   28
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "University"
         Height          =   735
         Left            =   600
         TabIndex        =   26
         Top             =   120
         Width           =   4215
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
            Left            =   360
            TabIndex        =   27
            Top             =   240
            Width           =   3495
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   17055
         _ExtentX        =   30083
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
         NumItems        =   20
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "school_id"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "school_name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "univ_name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "school_phone"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "school_add"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "no_of_dpt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "school_stab"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "school_email"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "school_web"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "school_name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "univ_name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "school_name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "univ_name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   14535
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
         Height          =   375
         Index           =   1
         Left            =   8520
         TabIndex        =   23
         Top             =   960
         Width           =   3255
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
         Left            =   8520
         TabIndex        =   21
         Top             =   3240
         Width           =   1335
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
         Left            =   6120
         TabIndex        =   20
         Top             =   3240
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
         Left            =   3600
         TabIndex        =   19
         Top             =   3240
         Width           =   1815
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
         Left            =   1440
         TabIndex        =   18
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox Text10 
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
         Left            =   2400
         TabIndex        =   17
         Top             =   2670
         Width           =   3615
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   2400
         TabIndex        =   15
         Top             =   2100
         Width           =   3495
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
         Height          =   405
         Left            =   2400
         TabIndex        =   13
         Top             =   1560
         Width           =   2895
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
         Height          =   1575
         Index           =   0
         Left            =   8520
         TabIndex        =   11
         Top             =   1440
         Width           =   4095
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
         Index           =   0
         Left            =   2400
         TabIndex        =   9
         Top             =   960
         Width           =   3015
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
         Left            =   12840
         TabIndex        =   8
         Top             =   360
         Width           =   1335
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
         Height          =   495
         Left            =   6600
         TabIndex        =   5
         Top             =   240
         Width           =   3375
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
         Height          =   495
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Phone :"
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
         Left            =   6720
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Website :"
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
         TabIndex        =   16
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Mail ID (If Any):"
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
         TabIndex        =   14
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Established(Year):"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Address :"
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
         Left            =   6840
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "No. of Departments :"
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
         Left            =   10080
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Under University :"
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
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "School Name :"
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
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "School ID :"
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "School Info :"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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
  cn.Execute "insert into school_master(school_id,school_name,univ_name,school_phone,school_add,no_of_dpt,school_stab,school_email,school_web) values('" + Text3(1).Text + "','" + Text1.Text + "','" + Text2(0).Text + "','" + Text4(1).Text + "','" + Text6(0).Text + "','" + Combo1.Text + "','" + Text8.Text + "','" + Text9(1).Text + "','" + Text10.Text + "');"
  MsgBox "Saved..."
  Call clr
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command11_Click()
  End
End Sub


Private Sub Command13_Click()
  Form8.Show
  Unload Me
End Sub

Private Sub Command14_Click()
  Call clr
  
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
   cn.Execute "update school_master set school_id='" + Text3(1).Text + "',school_name='" + Text1.Text + "'" + ",univ_name='" + Text2(0).Text + "'" + ",school_phone='" + Text4(1).Text + "'" + ",school_add='" + Text6(0).Text + "'" + ",no_of_dpt='" + Combo1.Text + "'" + ",school_stab='" + Text8.Text + "'" + ",school_email='" + Text9(1).Text + "'" + ",school_web='" + Text10.Text + "'" + " where school_id='" + Text3(1).Text + "';"
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
  dcom.CommandText = "delete from school_master where school_id='" + Text3(1).Text + "'"
  dcom.CommandType = adCmdText
  Set dcom.ActiveConnection = cn
  dcom.Execute
  Call clr
  
  cn.Close
  Set cn = Nothing
  Call clr
  loadlist
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command4_Click()
   Form5.Show
   Unload Me
End Sub



Private Sub Command5_Click()
 loadlist
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
  Form4.WindowState = 2
  Dim i As Integer
  For i = 1 To 50
  Combo1.AddItem i
  Next i
  Call com2
  loadlist
End Sub
Public Sub com2()
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
  sql = "select distinct univ_name from school_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
    Combo2.AddItem rs("univ_name")
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




Public Sub clr()
   Text3(1).Text = ""
Text1.Text = ""
Text2(0).Text = ""
  Text4(1).Text = ""
  Text6(0).Text = ""
  Combo1.Text = ""
  Text8.Text = ""
  Text9(1).Text = ""
  Text10.Text = ""
End Sub
Public Sub loadlist()
    On Error Resume Next
    Dim li As ListItem
    Dim LV As ListView
     Dim cmd As String
  Dim cn As ADODB.Connection
  Dim rs As ADODB.Recordset
  cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
  Set cn = New ADODB.Connection
  With cn
  .ConnectionString = cmd
  .Open
  End With
 '   Dim rec As Integer
    Dim sSQL  As String
    sSQL = "select * from school_master where univ_name like '" & Combo2.Text & "%'"
    
   Set rs = New ADODB.Recordset
   With rs
   .Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
   
     Set LV = ListView1
    LV.ListItems.Clear
LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "School ID", 1200, lvwColumnLeft
    LV.ColumnHeaders.Add , , "School name", 2500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "University name", 2500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "No. of Dept", 1200, lvwColumnLeft
   
    LV.ColumnHeaders.Add , , "Established", 1500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Phone", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Mail ID", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "website", 2000, lvwColumnLeft
     LV.ColumnHeaders.Add , , "Address", 10000, lvwColumnLeft
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
   '     rec = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("school_id") & "")
              li.ListSubItems.Add , , rs("school_name") & ""
              li.ListSubItems.Add , , rs("univ_name") & ""
              li.ListSubItems.Add , , rs("no_of_dpt") & ""
             
              li.ListSubItems.Add , , rs("school_stab") & ""
              li.ListSubItems.Add , , rs("school_phone") & ""
          li.ListSubItems.Add , , rs("school_email") & ""
            li.ListSubItems.Add , , rs("school_web") & ""
             li.ListSubItems.Add , , rs("school_add") & ""
        rs.MoveNext
        Loop
        
    End If
    
    rs.Close
    End With
    
    cn.Close
    
End Sub
Private Sub ListView1_DblClick()

     Text3(1).Text = ListView1.SelectedItem.Text
    Text1.Text = ListView1.SelectedItem.SubItems(1)
     Text2(0).Text = ListView1.SelectedItem.SubItems(2)
      Combo1.Text = ListView1.SelectedItem.SubItems(3)
     Text8.Text = ListView1.SelectedItem.SubItems(4)
     Text4(1).Text = ListView1.SelectedItem.SubItems(5)
     Text9(1).Text = ListView1.SelectedItem.SubItems(6)
     Text10.Text = ListView1.SelectedItem.SubItems(7)
     Text6(0).Text = ListView1.SelectedItem.SubItems(8)
End Sub


