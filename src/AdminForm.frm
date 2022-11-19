VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form StaffForm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8025
   ClientLeft      =   5955
   ClientTop       =   2925
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12255
   Begin VB.Frame Sidebar 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton UsersBtn 
         Caption         =   "Users"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CommandButton AccountBtn 
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CommandButton LogoutBtn 
         Caption         =   "Logout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   6840
         Width           =   2295
      End
      Begin VB.CommandButton HomeBtn 
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "12345"
         Height          =   195
         Left            =   1395
         TabIndex        =   8
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "STAFFID:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Logged in as:"
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame HomeFrame 
      Caption         =   "HomeFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   285
         Left            =   7920
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Open Student Form"
         Height          =   525
         Left            =   8040
         TabIndex        =   12
         Top             =   7200
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5415
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9551
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Search"
         Height          =   285
         Left            =   6720
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1080
         Width           =   6255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Results:"
         Height          =   195
         Left            =   3600
         TabIndex        =   16
         Top             =   7200
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Page:"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Top             =   7200
         Width           =   420
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Manage Enrollees"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   13
         Top             =   360
         Width           =   2580
      End
   End
   Begin VB.Frame UsersFrame 
      Caption         =   "UsersFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.Label Label8 
         Caption         =   "This is Users panel."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3120
         TabIndex        =   19
         Top             =   3360
         Width           =   4335
      End
   End
   Begin VB.Frame AccountFrame 
      Caption         =   "AccountFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.Label Label9 
         Caption         =   "This is Account panel."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2760
         TabIndex        =   21
         Top             =   3240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "StaffForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    HomeFrame.Caption = ""
    UsersFrame.Caption = ""
    AccountFrame.Caption = ""
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub HomeBtn_Click()
    HomeFrame.Visible = True
    UsersFrame.Visible = False
    AccountFrame.Visible = False
End Sub


Private Sub UsersBtn_Click()
    HomeFrame.Visible = False
    UsersFrame.Visible = True
    AccountFrame.Visible = False
End Sub

Private Sub AccountBtn_Click()
    HomeFrame.Visible = False
    UsersFrame.Visible = False
    AccountFrame.Visible = True
End Sub
