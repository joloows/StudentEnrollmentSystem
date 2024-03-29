VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form StaffForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AES Enrollment System"
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
      Begin VB.CommandButton RegistrarBtn 
         Caption         =   "Registrars"
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
         TabIndex        =   44
         Top             =   4560
         Width           =   2295
      End
      Begin VB.CommandButton AdminBtn 
         Caption         =   "Admins"
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
         TabIndex        =   14
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
         Top             =   5640
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
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   7320
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
      Begin VB.Label lblStaffId 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "<staff id>"
         Height          =   195
         Left            =   1290
         TabIndex        =   8
         Top             =   240
         Width           =   675
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
      Begin VB.Label lblUsername 
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
   Begin VB.Frame AccountFrame 
      Caption         =   "AccountFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox txtAuth1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   35
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CommandButton btnChangePassword 
         Caption         =   "Submit"
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   6360
         Width           =   855
      End
      Begin VB.TextBox txtConfirm 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   33
         Top             =   5760
         Width           =   2535
      End
      Begin VB.TextBox txtNewPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   26
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox txtAuth2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton btnChangeUsername 
         Caption         =   "Submit"
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtNewUsername 
         Height          =   285
         Left            =   2880
         TabIndex        =   22
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter current password:"
         Height          =   195
         Left            =   720
         TabIndex        =   36
         Top             =   2760
         Width           =   1680
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Confirm password:"
         Height          =   195
         Left            =   720
         TabIndex        =   32
         Top             =   5760
         Width           =   1290
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "New password:"
         Height          =   195
         Left            =   720
         TabIndex        =   31
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Enter current password:"
         Height          =   195
         Left            =   720
         TabIndex        =   30
         Top             =   4560
         Width           =   1680
      End
      Begin VB.Label lblAccountUsername 
         Caption         =   "<username>"
         Height          =   255
         Left            =   2880
         TabIndex        =   29
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "New username:"
         Height          =   195
         Left            =   720
         TabIndex        =   28
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Current username:"
         Height          =   195
         Left            =   720
         TabIndex        =   27
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
         TabIndex        =   23
         Top             =   3960
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   360
         Width           =   2400
      End
   End
   Begin VB.Frame RegistrarFrame 
      Caption         =   "RegistrarFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   63
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton btnRefreshReg 
         DisabledPicture =   "StaffForm.frx":0000
         Height          =   375
         Left            =   240
         Picture         =   "StaffForm.frx":424A
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   1050
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   71
         Top             =   1100
         Width           =   3495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Search"
         Height          =   285
         Left            =   4320
         TabIndex        =   70
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton btnCreateAcc 
         Caption         =   "Create Account"
         Height          =   525
         Index           =   0
         Left            =   7920
         TabIndex        =   69
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton btnRfirst 
         Caption         =   "<< First"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   68
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnRprev 
         Caption         =   "< Previous"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   67
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Go to page"
         Height          =   375
         Left            =   3960
         TabIndex        =   66
         Top             =   7080
         Width           =   1335
      End
      Begin VB.CommandButton btnRnext 
         Caption         =   "Next >"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   65
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnRlast 
         Caption         =   "Last >>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   64
         Top             =   7080
         Width           =   975
      End
      Begin MSComctlLib.ListView RegistrarLV 
         Height          =   5415
         Left            =   1440
         TabIndex        =   72
         Top             =   1560
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   9551
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Manage Registrars"
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
         Left            =   3330
         TabIndex        =   79
         Top             =   360
         Width           =   2640
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Results:"
         Height          =   195
         Left            =   480
         TabIndex        =   78
         Top             =   7560
         Width           =   570
      End
      Begin VB.Label txtRResult 
         AutoSize        =   -1  'True
         Caption         =   "total"
         Height          =   195
         Left            =   1200
         TabIndex        =   77
         Top             =   7560
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Showing:"
         Height          =   195
         Left            =   7320
         TabIndex        =   76
         Top             =   7560
         Width           =   660
      End
      Begin VB.Label txtRIndex 
         AutoSize        =   -1  'True
         Caption         =   "record"
         Height          =   195
         Left            =   8160
         TabIndex        =   75
         Top             =   7560
         Width           =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Page:"
         Height          =   195
         Left            =   3600
         TabIndex        =   74
         Top             =   7560
         Width           =   420
      End
      Begin VB.Label txtRPages 
         AutoSize        =   -1  'True
         Caption         =   "(cur p.) of (total p.)"
         Height          =   195
         Left            =   4200
         TabIndex        =   73
         Top             =   7560
         Width           =   1290
      End
   End
   Begin VB.Frame AdminFrame 
      Caption         =   "AdminFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   45
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton btnRefreshAdmin 
         DisabledPicture =   "StaffForm.frx":46D4
         Height          =   375
         Left            =   240
         Picture         =   "StaffForm.frx":891E
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1050
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   54
         Top             =   1100
         Width           =   3495
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Search"
         Height          =   285
         Left            =   4320
         TabIndex        =   53
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton btnCreateAcc 
         Caption         =   "Create Account"
         Height          =   525
         Index           =   1
         Left            =   7920
         TabIndex        =   51
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton btnAfirst 
         Caption         =   "<< First"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   50
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnAprev 
         Caption         =   "< Previous"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   49
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Go to page"
         Height          =   375
         Left            =   3960
         TabIndex        =   48
         Top             =   7080
         Width           =   1335
      End
      Begin VB.CommandButton btnAnext 
         Caption         =   "Next >"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   47
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnAlast 
         Caption         =   "Last >>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   46
         Top             =   7080
         Width           =   975
      End
      Begin MSComctlLib.ListView AdminLV 
         Height          =   5415
         Left            =   1440
         TabIndex        =   52
         Top             =   1560
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   9551
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Manage Admins"
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
         Left            =   3510
         TabIndex        =   61
         Top             =   360
         Width           =   2280
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Results:"
         Height          =   195
         Left            =   480
         TabIndex        =   60
         Top             =   7560
         Width           =   570
      End
      Begin VB.Label txtAResult 
         AutoSize        =   -1  'True
         Caption         =   "total"
         Height          =   195
         Left            =   1200
         TabIndex        =   59
         Top             =   7560
         Width           =   300
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Showing:"
         Height          =   195
         Left            =   7320
         TabIndex        =   58
         Top             =   7560
         Width           =   660
      End
      Begin VB.Label txtAIndex 
         AutoSize        =   -1  'True
         Caption         =   "record"
         Height          =   195
         Left            =   8160
         TabIndex        =   57
         Top             =   7560
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Page:"
         Height          =   195
         Left            =   3600
         TabIndex        =   56
         Top             =   7560
         Width           =   420
      End
      Begin VB.Label txtAPages 
         AutoSize        =   -1  'True
         Caption         =   "(cur p.) of (total p.)"
         Height          =   195
         Left            =   4200
         TabIndex        =   55
         Top             =   7560
         Width           =   1290
      End
   End
   Begin VB.Frame EnrolleesFrame 
      Caption         =   "EnrolleesFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton btnRefreshEn 
         DisabledPicture =   "StaffForm.frx":8DA8
         Height          =   375
         Left            =   240
         Picture         =   "StaffForm.frx":CFF2
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1050
         Width           =   375
      End
      Begin VB.CommandButton btnElast 
         Caption         =   "Last >>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   43
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnEnext 
         Caption         =   "Next >"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   42
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnEGoto 
         Caption         =   "Go to page"
         Height          =   375
         Left            =   3960
         TabIndex        =   41
         Top             =   7080
         Width           =   1335
      End
      Begin VB.CommandButton btnEprev 
         Caption         =   "< Previous"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   40
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnEfirst 
         Caption         =   "<< First"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   39
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnOpnStudForm 
         Caption         =   "Open Student Form"
         Height          =   525
         Left            =   7920
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin MSComctlLib.ListView EnrolleeLV 
         Height          =   5415
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9551
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Search"
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   1100
         Width           =   3495
      End
      Begin VB.Label txtEPages 
         AutoSize        =   -1  'True
         Caption         =   "<<cur p.> of <total p.>>"
         Height          =   195
         Left            =   4200
         TabIndex        =   38
         Top             =   7560
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Page:"
         Height          =   195
         Left            =   3600
         TabIndex        =   37
         Top             =   7560
         Width           =   420
      End
      Begin VB.Label txtEIndex 
         AutoSize        =   -1  'True
         Caption         =   "<record>"
         Height          =   195
         Left            =   8160
         TabIndex        =   19
         Top             =   7560
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Showing:"
         Height          =   195
         Left            =   7320
         TabIndex        =   18
         Top             =   7560
         Width           =   660
      End
      Begin VB.Label txtEResult 
         AutoSize        =   -1  'True
         Caption         =   "<total>"
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   7560
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Results:"
         Height          =   195
         Left            =   480
         TabIndex        =   16
         Top             =   7560
         Width           =   570
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
End
Attribute VB_Name = "StaffForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private result As Collection
Public search As String

Public selectedEnrollee As Enrollee
Public selectedUser As user

Public eCurrentPage As Integer
Public eTotalPage As Integer
Public aCurrentPage As Integer
Public aTotalPage As Integer
Public rCurrentPage As Integer
Public rTotalPage As Integer



' #################### ENROLLEE BUTTONS ####################
Private Sub btnEGoto_Click()
    GoToForm.Show vbModal
    
    Set result = GetEnrollee(eCurrentPage, search)
    Call InitPagination("enrollee", result)
End Sub

Private Sub btnEfirst_Click()
    eCurrentPage = 1
    
    Set result = GetEnrollee(eCurrentPage, search)
    Call InitPagination("enrollee", result)
End Sub

Private Sub btnElast_Click()
    eCurrentPage = eTotalPage
    
    Set result = GetEnrollee(eCurrentPage, search)
    Call InitPagination("enrollee", result)
End Sub

Private Sub btnEnext_Click()
    eCurrentPage = eCurrentPage + 1
    
    Set result = GetEnrollee(eCurrentPage, search)
    Call InitPagination("enrollee", result)
End Sub

Private Sub btnEprev_Click()
    eCurrentPage = eCurrentPage - 1
    
    Set result = GetEnrollee(eCurrentPage, search)
    Call InitPagination("enrollee", result)
End Sub

' #################### ADMIN BUTTONS ####################
Private Sub btnAfirst_Click()
    aCurrentPage = 1
    
    Set result = GetUser(True, aCurrentPage, search)
    Call InitPagination("admin", result)
End Sub

Private Sub btnAlast_Click()
    aCurrentPage = aTotalPage
    
    Set result = GetUser(True, aCurrentPage, search)
    Call InitPagination("admin", result)
End Sub

Private Sub btnAnext_Click()
    aCurrentPage = aCurrentPage + 1
    
    Set result = GetUser(True, aCurrentPage, search)
    Call InitPagination("admin", result)
End Sub

Private Sub btnAprev_Click()
    aCurrentPage = aCurrentPage - 1
    
    Set result = GetUser(True, aCurrentPage, search)
    Call InitPagination("admin", result)
End Sub

' #################### REGISTRAR BUTTONS ####################
Private Sub btnRfirst_Click()
    rCurrentPage = 1
    
    Set result = GetUser(False, rCurrentPage, search)
    Call InitPagination("registrar", result)
End Sub

Private Sub btnRlast_Click()
    rCurrentPage = rTotalPage
    
    Set result = GetUser(False, rCurrentPage, search)
    Call InitPagination("registrar", result)
End Sub

Private Sub btnRnext_Click()
    rCurrentPage = rCurrentPage + 1
    
    Set result = GetUser(False, rCurrentPage, search)
    Call InitPagination("registrar", result)
End Sub

Private Sub btnRprev_Click()
    rCurrentPage = rCurrentPage - 1
    
    Set result = GetUser(False, rCurrentPage, search)
    Call InitPagination("registrar", result)
End Sub

Private Sub btnRefreshEn_Click()
    eCurrentPage = 1
    Set result = GetEnrollee()
    Call InitPagination("enrollee", result)
    
    txtSearch.Text = ""
End Sub

Private Sub btnRefreshAdmin_Click()
    aCurrentPage = 1
    Set result = GetUser(True)
    Call InitPagination("admin", result)
    
    Text1.Text = ""
End Sub

Private Sub btnRefreshReg_Click()
    rCurrentPage = 1
    Set result = GetUser(False)
    Call InitPagination("registrar", result)
    
    Text2.Text = ""
End Sub

Private Sub btnSearch_Click()
    eCurrentPage = 1
    search = txtSearch.Text
    
    Set result = GetEnrollee(, search)
    Call InitPagination("enrollee", result)
End Sub
' ##########################################################

' TODO: CODE FOR ADMIN AND REGISTRAR PAGINATION BUTTONS

Private Sub btnOpenStudForm_Click()
    StudentForm.Show
End Sub

Private Sub btnCreateAcc_Click(Index As Integer)
    CreateAccForm.Show
End Sub

Private Sub btnOpnStudForm_Click()
    StudentForm.Show
End Sub

Private Sub btnChangeUsername_Click()
    
    If CurrentUser.password = txtAuth1.Text Then
        CurrentUser.username = txtNewUsername.Text
        
        Call UpdateUser(CurrentUser.id, CurrentUser)
        
        lblUsername.Caption = CurrentUser.username
        lblAccountUsername.Caption = lblUsername.Caption
        
        txtNewUsername.Text = ""
        txtAuth1.Text = ""
    Else
        MsgBox "Current password field input does not match the current user's password", vbExclamation, "Error"
    End If
End Sub

Private Sub btnChangePassword_Click()
    If Len(txtNewPassword.Text) < 8 Then
        GoTo PasswordInvalidError
    End If
    If CurrentUser.password = txtAuth2.Text Then
        If txtNewPassword.Text = txtConfirm.Text Then
            CurrentUser.password = txtNewPassword.Text
            
            Call UpdateCurrentUser(CurrentUser.id, CurrentUser)
            
            txtNewPassword.Text = ""
            txtAuth2.Text = ""
            txtConfirm.Text = ""
        Else
            MsgBox "Passwords does not match.", vbExclamation, "Error"
            Exit Sub
        End If
    Else
        MsgBox "Current password field input does not match the current user's password", vbExclamation, "Error"
    End If
    Exit Sub
PasswordInvalidError:
    MsgBox "Password should be at least 8 characters.", vbExclamation, "Invalid password"
End Sub

Private Sub LogoutBtn_Click()
    
    x = MsgBox("Are you sure you want to logout?", vbYesNo + vbQuestion, "Logout")
    If x = 6 Then ' If yes
        Call LogoutUser
    End If
    
End Sub

' Handles the Frame to be shown/hide to the user based on the button clicked
Private Sub EnrolleesBtn_Click()
    EnrolleesFrame.Visible = True
    AdminFrame.Visible = False
    RegistrarFrame.Visible = False
    AccountFrame.Visible = False
End Sub

Private Sub AdminBtn_Click()
    EnrolleesFrame.Visible = False
    AdminFrame.Visible = True
    RegistrarFrame.Visible = False
    AccountFrame.Visible = False
End Sub

Private Sub RegistrarBtn_Click()
    EnrolleesFrame.Visible = False
    AdminFrame.Visible = False
    RegistrarFrame.Visible = True
    AccountFrame.Visible = False
End Sub

Private Sub AccountBtn_Click()
    EnrolleesFrame.Visible = False
    AdminFrame.Visible = False
    RegistrarFrame.Visible = False
    AccountFrame.Visible = True
End Sub

' Yung pagination doon nakadisplay yung mga records ng database but paginated.
' i.e. pages 1-44 yung records ng enrollees. Pagination ang tawag po sa way ng
' pag display ng data. ignore
Public Sub InitPagination(whatToPaginate As String, items As Collection)
    Dim Li As ListItem
    
    If whatToPaginate = "enrollee" Then
        eTotalPage = items("pages")
        
        Call ButtonPaginationHelper(eCurrentPage, eTotalPage, btnEfirst, _
        btnEprev, btnEnext, btnElast)
    
        EnrolleeLV.ListItems.Clear
        
        For Each En In items("enrollees")
            Set Li = EnrolleeLV.ListItems.Add(, , En.id)
            With Li
                .SubItems(1) = IIf(En.Enrolled, "Y", "N")
                .SubItems(2) = En.Grade
                .SubItems(3) = En.Section
                .SubItems(4) = En.Lname
                .SubItems(5) = En.Fname
                .SubItems(6) = En.Mname
                .SubItems(7) = En.TotalFee
                .SubItems(8) = En.payment
                .SubItems(9) = IIf(En.WithUniform, "Y", "N")
                .SubItems(10) = En.PaymentType
                .SubItems(11) = En.Sex
                .SubItems(12) = En.Age
                .SubItems(13) = En.Birthdate
                .SubItems(14) = En.Birthplace
                .SubItems(15) = En.Mt
                .SubItems(16) = En.Address
                .SubItems(17) = En.Submission
                .SubItems(18) = En.Fathername
                .SubItems(19) = En.Fnum
                .SubItems(20) = En.MotherName
                .SubItems(21) = En.Mnum
                .SubItems(22) = En.GuardianName
                .SubItems(23) = En.Gnum
            End With
        Next
        
        Call PaginationInfo("enrollee", items, EnrolleeLV, txtEResult, _
        txtEIndex, txtEPages)
        
    ElseIf whatToPaginate = "admin" Then
        aTotalPage = items("pages")
        Call ButtonPaginationHelper(aCurrentPage, aTotalPage, btnAfirst, _
        btnAprev, btnAnext, btnAlast)
        
        Call UserPaginationHelper(AdminLV, items, "admin")
        
        Call PaginationInfo("admin", items, AdminLV, txtAResult, _
        txtAIndex, txtAPages)
    ElseIf whatToPaginate = "registrar" Then
        rTotalPage = items("pages")
        Call ButtonPaginationHelper(rCurrentPage, rTotalPage, btnRfirst, _
        btnRprev, btnRnext, btnRlast)
        
        Call UserPaginationHelper(RegistrarLV, items, "registrar")
        
        Call PaginationInfo("registrar", items, RegistrarLV, txtRResult, _
        txtRIndex, txtRPages)
    End If
    
End Sub

' Setups the AdminLV and RegistrarLV for pagination. ignore
Private Sub UserPaginationHelper(LV As ListView, items As Collection, userType As String)
    LV.ListItems.Clear
    If userType = "admin" Then
        For Each u In items("users")
            Set Li = LV.ListItems.Add(, , u.id)
            With Li
                .SubItems(1) = u.username
                .SubItems(2) = u.isAdmin
                .SubItems(3) = u.dateCreated
            End With
        Next
    
    ElseIf userType = "registrar" Then
        For Each u In items("users")
            Set Li = LV.ListItems.Add(, , u.id)
            With Li
                .SubItems(1) = u.username
                .SubItems(2) = u.password
                .SubItems(3) = u.isAdmin
                .SubItems(4) = u.dateCreated
            End With
        Next
    End If
    
End Sub

' Sets up the (first, prev, next, last) buttons on the pagination
Sub ButtonPaginationHelper(currentPage As Integer, totalPage As Integer, _
firstBtn As CommandButton, prevBtn As CommandButton, nextBtn As CommandButton, _
lastBtn As CommandButton)
    
    ' If at first page, disable first and previous buttons
    If currentPage <= 1 Then
        firstBtn.Enabled = False
        prevBtn.Enabled = False
    Else
        firstBtn.Enabled = True
        prevBtn.Enabled = True
    End If
    
    ' If at last page, disable next and last buttons
    Debug.Print currentPage
    Debug.Print totalPage
    If currentPage = totalPage Then
        nextBtn.Enabled = False
        lastBtn.Enabled = False
    Else
        nextBtn.Enabled = True
        lastBtn.Enabled = True
    End If
End Sub

Sub PaginationInfo(sender As String, items As Collection, LV As ListView, txtResult As Label, _
txtIndex As Label, txtPages As Label)
        
        ' Check what recordset are we processing
        If sender = "enrollee" Then
            currentPage = eCurrentPage
            totalPage = eTotalPage
        ElseIf sender = "admin" Then
            currentPage = aCurrentPage
            totalPage = aTotalPage
        ElseIf sender = "registrar" Then
            currentPage = rCurrentPage
            totalPage = rTotalPage
        End If
        
        ' Update the info at the bottom of pagination
        txtResult.Caption = items("recordCount")
        If LV.ListItems.Count = 0 Or LV.ListItems.Count = 1 Then
            txtIndex.Caption = txtResult.Caption
        Else
            txtIndex.Caption = items("startIndex") & "-" & items("stopIndex")
        End If
        
        txtPages.Caption = currentPage & " of " & totalPage
End Sub

Private Sub Form_Load()
    
    ' Remove frame controls caption
    EnrolleesFrame.Caption = ""
    AdminFrame.Caption = ""
    RegistrarFrame.Caption = ""
    AccountFrame.Caption = ""
    
    ' Replace placeholder captions
    lblUsername.Caption = CurrentUser.username
    lblStaffId.Caption = CurrentUser.id
    lblAccountUsername.Caption = lblUsername.Caption
    
    EnrolleeLV.FullRowSelect = True
    AdminLV.FullRowSelect = True
    RegistrarLV.FullRowSelect = True
    
    ' Enables or disables StaffForm buttons depending on CurrentUser privileges.
    If CurrentUser.isAdmin = False Then
        AdminBtn.Enabled = False
        RegistrarBtn.Enabled = False
        AccountBtn.Enabled = False
    End If
    
    ' Add listview column headers to EnrolleeLV

    With EnrolleeLV.ColumnHeaders
        .Add , , "Id", 500
        .Add , , "Enrolled", 900, lvwColumnCenter
        .Add , , "Grade", 900, lvwColumnCenter
        .Add , , "Section", 1200, lvwColumnCenter
        .Add , , "Last Name", 1200, lvwColumnCenter
        .Add , , "First Name", 1200, lvwColumnCenter
        .Add , , "Middle Name", 1200, lvwColumnCenter
        .Add , , "Total Fees", 1200, lvwColumnCenter
        .Add , , "Payment", 1200, lvwColumnCenter
        .Add , , "With Uniform", 1200, lvwColumnCenter
        .Add , , "Payment Type", 1400, lvwColumnCenter
        .Add , , "Sex", 700
        .Add , , "Age", 700, lvwColumnCenter
        .Add , , "Birthdate", 1200, lvwColumnCenter
        .Add , , "Birthplace", 1200, lvwColumnCenter
        .Add , , "Mother Toungue", 1400, lvwColumnCenter
        .Add , , "Address", 1200
        .Add , , "Date Enrolled", 1200
        .Add , , "Father Name", 1200
        .Add , , "Father No.", 1200
        .Add , , "Mother Name", 1200
        .Add , , "Mother No.", 1200
        .Add , , "Guardian Name", 1200
        .Add , , "Guardian No.", 1200
    End With
    
    Call SetupHelper(AdminLV, "admin")
    Call SetupHelper(RegistrarLV, "registrar")
    
    ' Set CurrentPage for all paginations to 1
    eCurrentPage = 1
    aCurrentPage = 1
    rCurrentPage = 1
    
    Set result = GetEnrollee()
    Call InitPagination("enrollee", result)
    
    ' GetUser(True) gets the admin users
    Set result = GetUser(True)
    Call InitPagination("admin", result)
    
    ' GetUser(False) gets the registrar users
    Set result = GetUser(False)
    Call InitPagination("registrar", result)
End Sub

Sub SetupHelper(LV As ListView, userType As String)
    If userType = "admin" Then
        With LV.ColumnHeaders
            .Add , , "Id", 500
            .Add , , "username", 1950, lvwColumnCenter
            .Add , , "IsAdmin", 1850, lvwColumnCenter
            .Add , , "Date Created", 2000, lvwColumnCenter
        End With
    ElseIf userType = "registrar" Then
        With LV.ColumnHeaders
            .Add , , "Id", 500
            .Add , , "username", 1400, lvwColumnCenter
            .Add , , "password", 1350, lvwColumnCenter
            .Add , , "IsAdmin", 1350, lvwColumnCenter
            .Add , , "Date Created", 1700, lvwColumnCenter
        End With
    End If
End Sub

Private Sub EnrolleeLV_DblClick()
    Dim En As New Enrollee
    With En
        .id = EnrolleeLV.SelectedItem
        .Enrolled = IIf(EnrolleeLV.SelectedItem.SubItems(1) = "Y", True, False)
        .Grade = EnrolleeLV.SelectedItem.SubItems(2)
        .Section = EnrolleeLV.SelectedItem.SubItems(3)
        .Lname = EnrolleeLV.SelectedItem.SubItems(4)
        .Fname = EnrolleeLV.SelectedItem.SubItems(5)
        .Mname = EnrolleeLV.SelectedItem.SubItems(6)
        .TotalFee = CCur(EnrolleeLV.SelectedItem.SubItems(7))
        .payment = CCur(EnrolleeLV.SelectedItem.SubItems(8))
        .WithUniform = IIf(EnrolleeLV.SelectedItem.SubItems(9) = "Y", True, False)
        .PaymentType = EnrolleeLV.SelectedItem.SubItems(10)
        .Sex = EnrolleeLV.SelectedItem.SubItems(11)
        .Age = EnrolleeLV.SelectedItem.SubItems(12)
        .Birthdate = EnrolleeLV.SelectedItem.SubItems(13)
        .Birthplace = EnrolleeLV.SelectedItem.SubItems(14)
        .Mt = EnrolleeLV.SelectedItem.SubItems(15)
        .Address = EnrolleeLV.SelectedItem.SubItems(16)
        .Submission = EnrolleeLV.SelectedItem.SubItems(17)
        .Fathername = EnrolleeLV.SelectedItem.SubItems(18)
        .Fnum = EnrolleeLV.SelectedItem.SubItems(19)
        .MotherName = EnrolleeLV.SelectedItem.SubItems(20)
        .Mnum = EnrolleeLV.SelectedItem.SubItems(21)
        .GuardianName = EnrolleeLV.SelectedItem.SubItems(22)
        .Gnum = EnrolleeLV.SelectedItem.SubItems(23)
    End With
    Set selectedEnrollee = En
    
    EnSelectForm.Show
End Sub

Private Sub AdminLV_DblClick()
    Debug.Print ("AdminLV")
    Call UserSelect(AdminLV.SelectedItem, "admin")
End Sub

Private Sub RegistrarLV_DblClick()
    Debug.Print ("RegistrarLV")
    Call UserSelect(RegistrarLV.SelectedItem, "registrar")
End Sub

Private Sub UserSelect(item As IListItem, userType As String)
    Dim u As New user
    
    If userType = "admin" Then
        MsgBox "You have no permission to manage admin accounts.", vbExclamation, "No permission"
        Exit Sub
    ElseIf userType = "registrar" Then
        With u
            .id = item
            .username = item.SubItems(1)
            .password = item.SubItems(2)
            .isAdmin = item.SubItems(3)
        End With
    End If
    Set selectedUser = u
    
    UserSelectForm.Show
End Sub



