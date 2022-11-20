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
         TabIndex        =   15
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
      Begin VB.CommandButton EnrolleesBtn 
         Caption         =   "Enrollees"
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
         Caption         =   "<username>"
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
         Caption         =   "Logged in as"
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame EnrolleesFrame 
      Caption         =   "EnrolleesFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton Command7 
         Caption         =   "Sort by"
         Height          =   285
         Left            =   7920
         TabIndex        =   33
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add Enrollee"
         Height          =   525
         Left            =   6960
         TabIndex        =   14
         Top             =   7200
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
         Left            =   6360
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   480
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "(first index) - (last index)"
         Height          =   195
         Left            =   3480
         TabIndex        =   32
         Top             =   7560
         Width           =   1635
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Showing:"
         Height          =   195
         Left            =   2640
         TabIndex        =   31
         Top             =   7560
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "(total queries)"
         Height          =   195
         Left            =   1200
         TabIndex        =   30
         Top             =   7560
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "<< < 1 2 3 4 5 > >> "
         Height          =   195
         Left            =   1440
         TabIndex        =   29
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Results:"
         Height          =   195
         Left            =   480
         TabIndex        =   28
         Top             =   7560
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Go to page:"
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Top             =   7200
         Width           =   840
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
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   480
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   1080
         Width           =   5655
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Search"
         Height          =   285
         Left            =   6360
         TabIndex        =   50
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Sort by"
         Height          =   285
         Left            =   7920
         TabIndex        =   49
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Create Account"
         Height          =   525
         Left            =   8280
         TabIndex        =   18
         Top             =   7200
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5415
         Left            =   240
         TabIndex        =   19
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "(first index) - (last index)"
         Height          =   195
         Left            =   3480
         TabIndex        =   26
         Top             =   7560
         Width           =   1635
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Showing:"
         Height          =   195
         Left            =   2640
         TabIndex        =   25
         Top             =   7560
         Width           =   660
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Manage Users"
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
         Left            =   3630
         TabIndex        =   24
         Top             =   360
         Width           =   2040
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Go to page:"
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   7200
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Results:"
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   7560
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "<< < 1 2 3 4 5 > >> "
         Height          =   195
         Left            =   1440
         TabIndex        =   21
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(total queries)"
         Height          =   195
         Left            =   1200
         TabIndex        =   20
         Top             =   7560
         Width           =   945
      End
   End
   Begin VB.Frame AccountFrame 
      Caption         =   "AccountFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton Command8 
         Caption         =   "Submit"
         Height          =   255
         Left            =   720
         TabIndex        =   48
         Top             =   5880
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2880
         TabIndex        =   47
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2880
         TabIndex        =   40
         Top             =   4680
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2880
         TabIndex        =   39
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Submit"
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2400
         TabIndex        =   36
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Confirm password:"
         Height          =   195
         Left            =   720
         TabIndex        =   46
         Top             =   5280
         Width           =   1290
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "New password:"
         Height          =   195
         Left            =   720
         TabIndex        =   45
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Enter current password:"
         Height          =   195
         Left            =   720
         TabIndex        =   44
         Top             =   4080
         Width           =   1680
      End
      Begin VB.Label Label24 
         Caption         =   "<username>"
         Height          =   255
         Left            =   2400
         TabIndex        =   43
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "New username:"
         Height          =   195
         Left            =   720
         TabIndex        =   42
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Current username:"
         Height          =   195
         Left            =   720
         TabIndex        =   41
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label Label21 
         Caption         =   "Change Password"
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
         Left            =   600
         TabIndex        =   37
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Change username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   35
         Top             =   1200
         Width           =   1890
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Manage Account"
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
         Left            =   3480
         TabIndex        =   34
         Top             =   360
         Width           =   2400
      End
   End
End
Attribute VB_Name = "StaffForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
    CreateAccForm.Show
End Sub

Private Sub Command5_Click()
    StudentForm.Show
End Sub

Private Sub Form_Load()
    EnrolleesFrame.Caption = ""
    UsersFrame.Caption = ""
    AccountFrame.Caption = ""
End Sub

Private Sub EnrolleesBtn_Click()
    EnrolleesFrame.Visible = True
    UsersFrame.Visible = False
    AccountFrame.Visible = False
End Sub

Private Sub UsersBtn_Click()
    EnrolleesFrame.Visible = False
    UsersFrame.Visible = True
    AccountFrame.Visible = False
End Sub

Private Sub AccountBtn_Click()
    EnrolleesFrame.Visible = False
    UsersFrame.Visible = False
    AccountFrame.Visible = True
End Sub
