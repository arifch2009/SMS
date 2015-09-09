VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   10635
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   19005
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   ScaleHeight     =   697.949
   ScaleMode       =   0  'User
   ScaleWidth      =   19005
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   15720
      Top             =   8040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=vb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
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
      Left            =   5400
      TabIndex        =   8
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
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
      Left            =   2520
      TabIndex        =   7
      Top             =   8520
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   7200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3960
      TabIndex        =   4
      Top             =   6360
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3960
      TabIndex        =   2
      Top             =   5760
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Password    :"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Username    :"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Login Type : "
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
      Left            =   1920
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "Student Management System-Developed by Arif && Purbankan"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   10815
   End
   Begin VB.Menu a 
      Caption         =   "&Exit"
      Index           =   1
      WindowList      =   -1  'True
      Begin VB.Menu b 
         Caption         =   "&Exit"
         Index           =   1
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu c 
      Caption         =   "&About"
      Begin VB.Menu d 
         Caption         =   "Software"
      End
      Begin VB.Menu e 
         Caption         =   "websites"
      End
   End
   Begin VB.Menu f 
      Caption         =   "Help"
      Begin VB.Menu g 
         Caption         =   "Software"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public username As String
Public pwd As String






Private Sub About_Click()

End Sub

Private Sub b_Click(Index As Integer)
   End
End Sub

Private Sub Combo1_Click()
   Select Case Combo1.ListIndex
        Case 0
               username = "Admin"
                Text1.Text = username
                
        Case 1
           username = "User"
                Text1.Text = username
        End Select
End Sub



Private Sub Command1_Click()
    If Text1.Text = "Admin" Then
        If pwd = "" Then
           Form2.Show
           Unload Me
        Else
          MsgBox "Incorrect Username & Password Combination.."
        End If
    End If
    If Text1.Text = "User" Then
       If pwd = "" Then
          Form3.Show
          Unload Me
      Else
        MsgBox "Incorrect Username & Password Combination.."
      End If
    End If
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Form_Load()
   Combo1.TabIndex = 1
   Text2.TabIndex = 2
   Command1.TabIndex = 3
   Command2.TabIndex = 4
   Combo1.AddItem "Administrator"
    Combo1.AddItem "User"
    Form1.WindowState = 2
End Sub

Private Sub Software_Click()

End Sub

Private Sub Text2_Change()
  pwd = Text2.Text
  
End Sub
