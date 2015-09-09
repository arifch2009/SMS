VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome-Admin"
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   19005
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   697.949
   ScaleMode       =   0  'User
   ScaleWidth      =   19005
   Begin VB.CommandButton Command10 
      Caption         =   "Mark Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   11
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
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
      Height          =   615
      Left            =   7320
      TabIndex        =   10
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Change Setting"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Course Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   8
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
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
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   5520
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
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
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   5520
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Goto Login Screeen"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exam Information"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   4
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      Height          =   4095
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   10215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      Caption         =   "       Welcome,You are in Administrative Section"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Width           =   9255
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
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   7215
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
      Caption         =   "&About"
      Index           =   3
      Begin VB.Menu d 
         Caption         =   "Software"
         Index           =   4
      End
      Begin VB.Menu e 
         Caption         =   "product"
         Index           =   5
      End
   End
   Begin VB.Menu h 
      Caption         =   "Options"
      Index           =   7
      Begin VB.Menu i 
         Caption         =   "Student Information"
      End
      Begin VB.Menu j 
         Caption         =   "Department Information"
      End
      Begin VB.Menu k 
         Caption         =   "Exam Information"
      End
      Begin VB.Menu l 
         Caption         =   "School Information"
      End
      Begin VB.Menu m 
         Caption         =   "Course Information"
      End
      Begin VB.Menu n 
         Caption         =   "Login Screen"
      End
   End
   Begin VB.Menu f 
      Caption         =   "Help"
      Index           =   6
      Begin VB.Menu g 
         Caption         =   "Software"
         Index           =   7
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_Click(Index As Integer)
  End
End Sub

Private Sub Command1_Click()
  Form4.Show

End Sub

Private Sub Command10_Click()
  Form12.Show
End Sub

Private Sub Command2_Click()
   Form50.Show

End Sub

Private Sub Command3_Click()
  Form11.Show

End Sub

Private Sub Command4_Click()
  Form1.Show
End Sub

Private Sub Command5_Click()
  Form6.Show
End Sub


Private Sub Command6_Click()
  Form10.Show

End Sub

Private Sub Command7_Click()
  Form7.Show
 
End Sub

Private Sub Command9_Click()
  End
End Sub

Private Sub Form_Load()
   Command1.TabIndex = 1
   Command2.TabIndex = 2
   Command3.TabIndex = 3
   Command5.TabIndex = 4
   Command6.TabIndex = 5
   Command7.TabIndex = 6
   Command8.TabIndex = 7
   Command9.TabIndex = 8
   Command4.TabIndex = 9
   
   Form2.WindowState = 2
End Sub
