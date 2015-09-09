VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form6 
   Caption         =   "Student-Entry"
   ClientHeight    =   10890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18855
   BeginProperty Font 
      Name            =   "Algerian"
      Size            =   21.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form6"
   ScaleHeight     =   704.785
   ScaleMode       =   0  'User
   ScaleWidth      =   19005
   Begin VB.CommandButton Command7 
      Caption         =   "Load Student List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7920
      TabIndex        =   71
      Top             =   960
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Caption         =   "Please Choose"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13080
      TabIndex        =   63
      Top             =   240
      Width           =   4215
      Begin VB.CommandButton Command8 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2400
         TabIndex        =   64
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Academic View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   10800
      TabIndex        =   4
      Top             =   1440
      Width           =   7695
      Begin VB.CommandButton Command10 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6000
         TabIndex        =   73
         Top             =   8520
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   68
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   67
         Top             =   600
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3000
         TabIndex        =   66
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   85458947
         CurrentDate     =   40480
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4200
         TabIndex        =   60
         Top             =   8520
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         TabIndex        =   59
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   600
         TabIndex        =   58
         Top             =   8520
         Width           =   1335
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2400
         TabIndex        =   57
         Top             =   7800
         Width           =   2055
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5880
         TabIndex        =   55
         Top             =   7200
         Width           =   1455
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2400
         TabIndex        =   53
         Top             =   7200
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "Qualified Exams"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   360
         TabIndex        =   31
         Top             =   3120
         Width           =   6975
         Begin VB.TextBox Text19 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5520
            TabIndex        =   49
            Top             =   3120
            Width           =   1095
         End
         Begin VB.TextBox Text18 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   5040
            TabIndex        =   47
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox Text17 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   5040
            TabIndex        =   45
            Top             =   1920
            Width           =   1815
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   5040
            TabIndex        =   43
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   2040
            TabIndex        =   40
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "Algerian"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   38
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1800
            TabIndex        =   36
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1800
            TabIndex        =   34
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "Percentage :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   48
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label Label27 
            Caption         =   "Year :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   46
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label Label26 
            Caption         =   "Board :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   44
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "Name :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            TabIndex        =   42
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Exam2 : "
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4440
            TabIndex        =   41
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Percentage :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   39
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "Year :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Board :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Name :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Exam1 :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1200
            TabIndex        =   32
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   30
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label36 
         Caption         =   "Rank :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   56
         Top             =   7800
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   54
         Top             =   7200
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "Entrance Name :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   52
         Top             =   7200
         Width           =   1935
      End
      Begin VB.Label Label33 
         Height          =   375
         Left            =   480
         TabIndex        =   51
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Date Of Admission :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   50
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label18 
         Caption         =   "Year Of Joining :"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   29
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label17 
         Caption         =   "Course Name :"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   28
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label16 
         Caption         =   "Department :"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Personal View"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   9975
      Begin VB.CommandButton Command9 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7920
         TabIndex        =   72
         Top             =   8640
         Width           =   1575
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6960
         TabIndex        =   70
         Top             =   7320
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2880
         TabIndex        =   65
         Top             =   7320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   85458945
         CurrentDate     =   40480
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2880
         TabIndex        =   62
         Top             =   8040
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5760
         TabIndex        =   26
         Top             =   8640
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3600
         TabIndex        =   25
         Top             =   8640
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1440
         TabIndex        =   24
         Top             =   8640
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   7320
         TabIndex        =   22
         Top             =   6480
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2880
         TabIndex        =   20
         Top             =   6480
         Width           =   2535
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
         Left            =   7680
         TabIndex        =   18
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2880
         TabIndex        =   16
         Top             =   5760
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3840
         Width           =   4335
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   4335
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2880
         TabIndex        =   10
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2880
         TabIndex        =   8
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2880
         TabIndex        =   6
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label13 
         Caption         =   "Cast :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   69
         Top             =   7320
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "Nationality :"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   61
         Top             =   8040
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Date Of Birth :"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   7320
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Religion :"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   21
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Family Income : (Per Annum)"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   19
         Top             =   6480
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Sex :"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   5640
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Phone (Father):"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Address : (Temporary)"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   13
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Address : (permanent) "
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Mother's Name :"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Father's Name :"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Student ID :"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Student Detail :"
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
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error GoTo SaveErr
    Dim sSQL As String
    
    sSQL = "select * from stu_personal"
        Dim cmd As String
    cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=trial"
    Set dbConn = New ADODB.Connection
    Set dbRec = New ADODB.Recordset
    
    dbConn.ConnectionString = cmd
    dbConn.Open
    
    dbRec.Open sSQL, dbConn, adOpenDynamic, adLockOptimistic
        
    With dbRec
        .AddNew
        .Fields("stu_id") = Text1.Text
        .Fields("stu_name") = Text2.Text
        
        .Fields("stu_fname") = Text3.Text
        
        .Fields("stu_mname") = Text4.Text
        
        .Fields("stu_tadd") = Text6.Text
        .Fields("stu_padd") = Text5.Text
        .Fields("stu_phone") = Text7.Text
        .Fields("stu_sex") = Combo1.Text
        .Fields("stu_income") = Text8.Text
        .Fields("stu_religion") = Combo2.Text
        .Fields("stu_dob") = DTPicker1.Value
        .Fields("stu_cast") = Combo5.Text
        .Fields("stu_nationality") = Text23.Text
        .Update
    End With
    
    dbRec.Close
    dbConn.Close
        
    Set dbConn = Nothing
    Set dbRec = Nothing
   MsgBox "Saved"
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command10_Click()
  Call clr2
End Sub
Public Sub clr2()
Text1.Text = ""
         Combo3.Text = ""
        Combo4.Text = ""
        Text11.Text = ""
        DTPicker2.Value = Date
        Text12.Text = ""
         Text13.Text = ""
         Text14.Text = ""
         Text15.Text = ""
        Text16.Text = ""
         Text16.Text = ""
         Text17.Text = ""
         Text18.Text = ""
        Text20.Text = ""
        Text21.Text = ""
         Text22.Text = ""
End Sub

Private Sub Command2_Click()
On Error GoTo SaveErr
    Dim sSQL As String
    
    sSQL = "select * from stu_personal where stu_id='" + Text1.Text + "'"
        Dim cmd As String
    cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=trial"
    Set dbConn = New ADODB.Connection
    Set dbRec = New ADODB.Recordset
    
    dbConn.ConnectionString = cmd
    dbConn.Open
    
    dbRec.Open sSQL, dbConn, adOpenDynamic, adLockOptimistic
        
    With dbRec
        .AddNew
        .Fields("stu_id") = Text1.Text
        .Fields("stu_name") = Text2.Text
        
        .Fields("stu_fname") = Text3.Text
        
        .Fields("stu_mname") = Text4.Text
        
        .Fields("stu_tadd") = Text6.Text
        .Fields("stu_padd") = Text5.Text
        .Fields("stu_phone") = Text7.Text
        .Fields("stu_sex") = Combo1.Text
        .Fields("stu_income") = Text8.Text
        .Fields("stu_religion") = Combo2.Text
        .Fields("stu_dob") = DTPicker1.Value
        .Fields("stu_cast") = Combo5.Text
        .Fields("stu_nationality") = Text23.Text
        .Update
    End With
    
    dbRec.Close
    dbConn.Close
        
    Set dbConn = Nothing
    Set dbRec = Nothing
   MsgBox "Saved"
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
  dcom.CommandText = "delete from stu_personal where stu_id='" + Text1.Text + "'"
  dcom.CommandType = adCmdText
  Set dcom.ActiveConnection = cn
  dcom.Execute
 
  cn.Close
  Set cn = Nothing
  Call clr
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command4_Click()
On Error GoTo SaveErr
    Dim sSQL As String
    
    sSQL = "select * from stu_aca where stu_id='" + Text1.Text + "'"
        Dim cmd As String
    cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=trial"
    Set dbConn = New ADODB.Connection
    Set dbRec = New ADODB.Recordset
    
    dbConn.ConnectionString = cmd
    dbConn.Open
    
    dbRec.Open sSQL, dbConn, adOpenDynamic, adLockOptimistic
        
    With dbRec
        .AddNew
        .Fields("stu_id") = Text1.Text
        .Fields("dpt_name") = Combo3.Text
        
        .Fields("course_name") = Combo4.Text
        .Fields("stu_year") = Text11.Text
        .Fields("stu_Dadd") = DTPicker2.Value
        .Fields("stu_q1name") = Text12.Text
        .Fields("stu_q1board") = Text13.Text
        .Fields("stu_q1year") = Text14.Text
        .Fields("stu_q1percent") = Text15.Text
        .Fields("stu_q2name") = Text16.Text
        .Fields("stu_q2board") = Text17.Text
        .Fields("stu_q2year") = Text18.Text
        .Fields("stu_q2percent") = Text19.Text
        .Fields("stu_ename") = Text20.Text
        .Fields("stu_eyear") = Text21.Text
        .Fields("stu_erank") = Text22.Text
        .Update
    End With
    
    dbRec.Close
    dbConn.Close
        
    Set dbConn = Nothing
    Set dbRec = Nothing
    MsgBox "Saved"
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command5_Click()
On Error GoTo SaveErr
    Dim sSQL As String
    
    sSQL = "select * from stu_aca"
        Dim cmd As String
    cmd = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=dsn1;Initial Catalog=trial"
    Set dbConn = New ADODB.Connection
    Set dbRec = New ADODB.Recordset
    
    dbConn.ConnectionString = cmd
    dbConn.Open
    
    dbRec.Open sSQL, dbConn, adOpenDynamic, adLockOptimistic
        
    With dbRec
        .AddNew
        .Fields("stu_id") = Text1.Text
        .Fields("dpt_name") = Combo3.Text
        
        .Fields("course_name") = Combo4.Text
        .Fields("stu_year") = Text11.Text
        .Fields("stu_Dadd") = DTPicker2.Value
        .Fields("stu_q1name") = Text12.Text
        .Fields("stu_q1board") = Text13.Text
        .Fields("stu_q1year") = Text14.Text
        .Fields("stu_q1percent") = Text15.Text
        .Fields("stu_q2name") = Text16.Text
        .Fields("stu_q2board") = Text17.Text
        .Fields("stu_q2year") = Text18.Text
        .Fields("stu_q2percent") = Text19.Text
        .Fields("stu_ename") = Text20.Text
        .Fields("stu_eyear") = Text21.Text
        .Fields("stu_erank") = Text22.Text
        .Update
    End With
    
    dbRec.Close
    dbConn.Close
        
    Set dbConn = Nothing
    Set dbRec = Nothing
    MsgBox "Saved"
Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command6_Click()
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
  dcom.CommandText = "delete from stu_aca where stu_id='" + Text1.Text + "' and exam_category='" + Combo1.Text + "'"
  dcom.CommandType = adCmdText
  Set dcom.ActiveConnection = cn
  dcom.Execute
 
  cn.Close
  Set cn = Nothing
  Call clr
  Exit Sub
SaveErr:
    MsgBox Err.Description
End Sub

Private Sub Command7_Click()
  Form2.Show
  Unload Me
End Sub

Private Sub Command8_Click()
  End
End Sub

Private Sub Command9_Click()
       Call clr1
    
End Sub
Public Sub clr1()
 Text1.Text = ""
         Text2.Text = ""
        
        Text3.Text = ""
        
        Text4.Text = ""
        
         Text6.Text = ""
         Text5.Text = ""
         Text7.Text = ""
         Combo1.Text = ""
         Text8.Text = ""
         Combo2.Text = ""
         DTPicker1.Value = Date
         Combo5.Text = ""
         Text23.Text = ""
End Sub
Private Sub Form_Load()
  
Form6.WindowState = 2
  
  Combo1.AddItem "Male"
  Combo1.AddItem "Female"
  
  Combo2.AddItem "Hindu"
  Combo2.AddItem "Islam"
  Combo2.AddItem "Christian"
  Combo2.AddItem "Budhist"
  Combo2.AddItem "Jainism"
  Combo5.AddItem "General"
  Combo5.AddItem "OBC"
  Combo5.AddItem "SC"
  Combo5.AddItem "ST"
  
  
  
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
     rs.MoveNext
    Loop
    .Close
    End With
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
  Call com4
End Sub

Private Sub com4()
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
  sql = "select distinct course_name from course_master"
  Set rs = New ADODB.Recordset
  With rs
  .Open sql, cn, adOpenForwardOnly, adLockReadOnly
 Do While Not rs.EOF
   Combo4.AddItem rs("course_name")
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


