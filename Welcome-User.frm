VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome-User"
   ClientHeight    =   11235
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   19005
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form3"
   ScaleHeight     =   907.188
   ScaleMode       =   0  'User
   ScaleWidth      =   15270
   Begin VB.CommandButton Command9 
      Caption         =   "Exam Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   6720
      Width           =   2895
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Go TO login Screen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   8
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "course Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Subject Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Department Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Student Mark Query"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   4
      Top             =   4920
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "School Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Student Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   3615
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000018&
      Caption         =   "Choose your operation from below options :"
      BeginProperty Font 
         Name            =   "Modern"
         Size            =   18
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   3120
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      Caption         =   "       Welcome,You are in User Section"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   0
      Top             =   1320
      Width           =   7335
   End
   Begin VB.Menu a 
      Caption         =   "&Exit"
      Index           =   1
      Begin VB.Menu b 
         Caption         =   "Exit"
         Index           =   2
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu c 
      Caption         =   "About"
      Index           =   3
      Begin VB.Menu d 
         Caption         =   "Software"
         Index           =   4
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_Click(Index As Integer)
   End
End Sub

Private Sub Command1_Click()
Form22.Show
End Sub

Private Sub Command2_Click()
Form15.Show
End Sub

Private Sub Command3_Click()
Form19.Show
End Sub

Private Sub Command4_Click()
Form16.Show
End Sub

Private Sub Command5_Click()
Form21.Show
End Sub

Private Sub Command6_Click()
Form18.Show
End Sub

Private Sub Command7_Click()
  Form1.Show
  Unload Me
End Sub

Private Sub Command8_Click()
  End
End Sub

Private Sub Command9_Click()
Form20.Show
End Sub

Private Sub Form_Load()
  Command1.TabIndex = 1
  Command2.TabIndex = 2
  Command3.TabIndex = 3
  Command4.TabIndex = 4
  Command5.TabIndex = 5
  Command6.TabIndex = 6
  Command8.TabIndex = 7
  Command7.TabIndex = 8
  Form3.WindowState = 2
End Sub
