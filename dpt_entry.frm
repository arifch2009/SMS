VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form50 
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
      Height          =   5535
      Left            =   960
      TabIndex        =   23
      Top             =   5160
      Width           =   17895
      Begin MSComctlLib.ListView ListView1 
         Height          =   4215
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   17175
         _ExtentX        =   30295
         _ExtentY        =   7435
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
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "School"
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
         Left            =   600
         TabIndex        =   24
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
            TabIndex        =   25
            Top             =   240
            Width           =   3495
         End
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
         Left            =   2400
         TabIndex        =   27
         Top             =   960
         Width           =   3735
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
         Height          =   375
         Index           =   1
         Left            =   8520
         TabIndex        =   22
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   14
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
         TabIndex        =   12
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
         TabIndex        =   10
         Top             =   1440
         Width           =   4095
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
         Height          =   375
         Left            =   7080
         TabIndex        =   5
         Top             =   240
         Width           =   3735
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
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Website(If Any):"
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
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "No. of courses :"
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
         Left            =   10920
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Under School Of :"
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
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Departmentl ID :"
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
         Width           =   2175
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
Attribute VB_Name = "Form50"
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
  cn.Execute "insert into dpt_master(dpt_id,dpt_name,school_name,dpt_phone,dpt_add,no_of_course,dpt_stab,dpt_email,dpt_web) values('" + Text3(1).Text + "','" + Text1.Text + "','" + Combo2.Text + "','" + Text4(1).Text + "','" + Text6(0).Text + "','" + Combo1.Text + "','" + Text8.Text + "','" + Text9(1).Text + "','" + Text10.Text + "');"
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
   cn.Execute "update dpt_master set dpt_id='" + Text3(1).Text + "',dpt_name='" + Text1.Text + "'" + ",school_name='" + Combo3.Text + "'" + ",dpt_phone='" + Text4(1).Text + "'" + ",dpt_add='" + Text6(0).Text + "'" + ",no_of_course='" + Combo1.Text + "'" + ",dpt_stab='" + Text8.Text + "'" + ",dpt_email='" + Text9(1).Text + "'" + ",dpt_web='" + Text10.Text + "'" + " where dpt_id='" + Text3(1).Text + "';"
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
  dcom.CommandText = "delete from dpt_master where dpt_id='" + Text3(1).Text + "'"
  dcom.CommandType = adCmdText
  Set dcom.ActiveConnection = cn
  dcom.Execute
  Call clr
  
  cn.Close
  Set cn = Nothing
  
  loadlist
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command4_Click()
  Call clr
End Sub



Private Sub Command5_Click()
 loadlist

End Sub




Private Sub Form_Load()
  Form50.WindowState = 2
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
  sql = "select distinct school_name from school_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
 Combo2.Text = ""
Combo3.Text = ""
    Combo3.AddItem rs("school_name")
    Combo2.AddItem rs("school_name")
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
Combo2.Text = ""
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
    sSQL = "select * from dpt_master where school_name like '" & Combo2.Text & "%'"
    
   Set rs = New ADODB.Recordset
   With rs
   .Open sSQL, cn, adOpenForwardOnly, adLockReadOnly
   
     Set LV = ListView1
    LV.ListItems.Clear
LV.ColumnHeaders.Clear
    LV.ColumnHeaders.Add , , "Department ID", 1200, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Department name", 2500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "School name", 2500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "No. of course", 1200, lvwColumnLeft
   
    LV.ColumnHeaders.Add , , "Established", 1500, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Phone", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "Mail ID", 2000, lvwColumnLeft
    LV.ColumnHeaders.Add , , "website", 2000, lvwColumnLeft
     LV.ColumnHeaders.Add , , "Address", 10000, lvwColumnLeft
    If (rs.RecordCount <> 0) Or Not (rs.RecordCount = -1) Then
   '     rec = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            Set li = LV.ListItems.Add(, , rs("dpt_id") & "")
              li.ListSubItems.Add , , rs("dpt_name") & ""
              li.ListSubItems.Add , , rs("school_name") & ""
              li.ListSubItems.Add , , rs("no_of_course") & ""
             
              li.ListSubItems.Add , , rs("dpt_stab") & ""
              li.ListSubItems.Add , , rs("dpt_phone") & ""
          li.ListSubItems.Add , , rs("dpt_email") & ""
            li.ListSubItems.Add , , rs("dpt_web") & ""
             li.ListSubItems.Add , , rs("dpt_add") & ""
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
     Combo3.Text = ListView1.SelectedItem.SubItems(2)
      Combo1.Text = ListView1.SelectedItem.SubItems(3)
     Text8.Text = ListView1.SelectedItem.SubItems(4)
     Text4(1).Text = ListView1.SelectedItem.SubItems(5)
     Text9(1).Text = ListView1.SelectedItem.SubItems(6)
     Text10.Text = ListView1.SelectedItem.SubItems(7)
     Text6(0).Text = ListView1.SelectedItem.SubItems(8)
End Sub


