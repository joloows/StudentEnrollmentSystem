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
      Begin VB.Label lblStaffId 
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
   Begin VB.Frame EnrolleesFrame 
      Caption         =   "EnrolleesFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton Command4 
         Caption         =   "Sort by"
         Height          =   285
         Left            =   6720
         TabIndex        =   58
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search by"
         Height          =   285
         Left            =   5520
         TabIndex        =   57
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Height          =   285
         Left            =   240
         TabIndex        =   56
         Top             =   1080
         Width           =   285
      End
      Begin VB.CommandButton btnElast 
         Caption         =   "Last >>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   55
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnEnext 
         Caption         =   "Next >"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5880
         TabIndex        =   54
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnEGoto 
         Caption         =   "Go to page"
         Height          =   375
         Left            =   3960
         TabIndex        =   53
         Top             =   7080
         Width           =   1335
      End
      Begin VB.CommandButton btnEprev 
         Caption         =   "< Previous"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   52
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton btnEfirst 
         Caption         =   "<< First"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   51
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
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label txtPages 
         AutoSize        =   -1  'True
         Caption         =   "(cur p.) of (total p.)"
         Height          =   195
         Left            =   4200
         TabIndex        =   50
         Top             =   7560
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Page:"
         Height          =   195
         Left            =   3600
         TabIndex        =   49
         Top             =   7560
         Width           =   420
      End
      Begin VB.Label txtIndex 
         AutoSize        =   -1  'True
         Caption         =   "record"
         Height          =   195
         Left            =   8160
         TabIndex        =   29
         Top             =   7560
         Width           =   450
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Showing:"
         Height          =   195
         Left            =   7320
         TabIndex        =   28
         Top             =   7560
         Width           =   660
      End
      Begin VB.Label txtResult 
         AutoSize        =   -1  'True
         Caption         =   "total"
         Height          =   195
         Left            =   1200
         TabIndex        =   27
         Top             =   7560
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Results:"
         Height          =   195
         Left            =   480
         TabIndex        =   26
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
   Begin VB.Frame UsersFrame 
      Caption         =   "UsersFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   480
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   1080
         Width           =   5655
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Search"
         Height          =   285
         Left            =   6360
         TabIndex        =   45
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton btnCreateAcc 
         Caption         =   "Create Account"
         Height          =   525
         Left            =   8280
         TabIndex        =   17
         Top             =   7200
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5415
         Left            =   240
         TabIndex        =   18
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
         TabIndex        =   25
         Top             =   7560
         Width           =   1635
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Showing:"
         Height          =   195
         Left            =   2640
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   360
         Width           =   2040
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Go to page:"
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   7200
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Results:"
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   7560
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "<< < 1 2 3 4 5 > >> "
         Height          =   195
         Left            =   1440
         TabIndex        =   20
         Top             =   7200
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(total queries)"
         Height          =   195
         Left            =   1200
         TabIndex        =   19
         Top             =   7560
         Width           =   945
      End
   End
   Begin VB.Frame AccountFrame 
      Caption         =   "AccountFrame"
      Height          =   7935
      Left            =   2760
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2880
         TabIndex        =   47
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Submit"
         Height          =   255
         Left            =   720
         TabIndex        =   44
         Top             =   6360
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2880
         TabIndex        =   43
         Top             =   5760
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2880
         TabIndex        =   36
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2880
         TabIndex        =   35
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Submit"
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2880
         TabIndex        =   32
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter current password:"
         Height          =   195
         Left            =   720
         TabIndex        =   48
         Top             =   2760
         Width           =   1680
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Confirm password:"
         Height          =   195
         Left            =   720
         TabIndex        =   42
         Top             =   5760
         Width           =   1290
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "New password:"
         Height          =   195
         Left            =   720
         TabIndex        =   41
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Enter current password:"
         Height          =   195
         Left            =   720
         TabIndex        =   40
         Top             =   4560
         Width           =   1680
      End
      Begin VB.Label lblAccountUsername 
         Caption         =   "<username>"
         Height          =   255
         Left            =   2880
         TabIndex        =   39
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "New username:"
         Height          =   195
         Left            =   720
         TabIndex        =   38
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Current username:"
         Height          =   195
         Left            =   720
         TabIndex        =   37
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
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   30
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
Private result As Collection
Private search As String
Public current_page As Integer
Public total_page As Integer

Private Sub btnEGoto_Click()
    GoToForm.Show vbModal
    
    Set result = Get_Enrollee(current_page, search)
    Call InitPagination(result)
End Sub

Private Sub btnEfirst_Click()
    current_page = 1
    
    Set result = Get_Enrollee(current_page, search)
    Call InitPagination(result)
End Sub

Private Sub btnElast_Click()
    current_page = total_page
    
    Set result = Get_Enrollee(current_page, search)
    Call InitPagination(result)
End Sub

Private Sub btnEnext_Click()
    current_page = current_page + 1
    
    Set result = Get_Enrollee(current_page, search)
    Call InitPagination(result)
End Sub

Private Sub btnEprev_Click()
    Dim result As Collection
    current_page = current_page - 1
    
    Set result = Get_Enrollee(current_page, search)
    Call InitPagination(result)
End Sub

Private Sub btnOpnStudForm_Click()
    StudentForm.Show
End Sub

Private Sub btnSearch_Click()
    current_page = 1
    search = txtSearch.Text
    
    Set result = Get_Enrollee(, search)
    Call InitPagination(result)
End Sub

Private Sub btnCreateAcc_Click()
    If CurrentUser.isAdmin Then
        CreateAccForm.Show
    Else
        MsgBox "Non-admin users cannot access this feature.", vbExclamation, "Denied access"
    End If
End Sub

Private Sub btnOpenStudForm_Click()
    StudentForm.Show
End Sub

Private Sub EnrolleeLV_DblClick()
    Debug.Print "Double clicked item: " & EnrolleeLV.SelectedItem.Index
    id = EnrolleeLV.SelectedItem
    Column1 = EnrolleeLV.SelectedItem.SubItems(1)
    Column2 = EnrolleeLV.SelectedItem.SubItems(2)
    Column3 = EnrolleeLV.SelectedItem.SubItems(3)
    Column4 = EnrolleeLV.SelectedItem.SubItems(4)
    Column5 = EnrolleeLV.SelectedItem.SubItems(5)
    MsgBox (id & vbNewLine & Column1 & vbNewLine & Column2 & vbNewLine & Column3 & vbNewLine & Column4 & vbNewLine & Column5)
End Sub

Private Sub LogoutBtn_Click()
    
    x = MsgBox("Are you sure you want to logout?", vbYesNo + vbQuestion, "Logout")
    If x = 6 Then ' If yes
        Call LogoutUser
        StudentForm.Show
        Unload StaffForm
    End If
    
End Sub

' Handles the Frame to be shown/hide to the user based on the button clicked

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

Public Sub InitPagination(items As Collection)
    Dim En As Enrollee
    Dim Li As ListItem
    
    total_page = items("pages")
    ' If at first page
    If current_page <= 1 Then
        btnEfirst.Enabled = False
        btnEprev.Enabled = False
    Else
        btnEfirst.Enabled = True
        btnEprev.Enabled = True
    End If
    
    If total_page > 1 Then
        ' If at last page
        If current_page = total_page Then
            btnEnext.Enabled = False
            btnElast.Enabled = False
        Else
            btnEnext.Enabled = True
            btnElast.Enabled = True
        End If
    Else
        btnEnext.Enabled = False
        btnElast.Enabled = False
    End If
    
    EnrolleeLV.ListItems.Clear

    For Each En In items("enrollees")
        Set Li = EnrolleeLV.ListItems.Add(, , En.id)
        With Li
            .SubItems(1) = En.Grade
            .SubItems(2) = En.Lname
            .SubItems(3) = En.Fname
            .SubItems(4) = En.Mname
            .SubItems(5) = En.Sex
            .SubItems(6) = En.Age
            .SubItems(7) = En.Birthdate
            .SubItems(8) = En.Birthplace
            .SubItems(9) = En.Mt
            .SubItems(10) = En.Address
            .SubItems(11) = En.Submission
            .SubItems(12) = En.Fathername
            .SubItems(13) = En.Fnum
            .SubItems(14) = En.MotherName
            .SubItems(15) = En.Mnum
            .SubItems(16) = En.GuardianName
            .SubItems(17) = En.Gnum
        End With
    Next
    
    txtResult.Caption = items("record_count")
    If EnrolleeLV.ListItems.Count = 0 Or EnrolleeLV.ListItems.Count = 1 Then
        txtIndex.Caption = txtResult.Caption
    Else
        txtIndex.Caption = items("start_index") & "-" & items("stop_index")
    End If
    txtPages.Caption = current_page & " of " & total_page
End Sub

Private Sub Form_Load()
    
    ' Remove frame controls caption
    EnrolleesFrame.Caption = ""
    UsersFrame.Caption = ""
    AccountFrame.Caption = ""
    
    ' Replace placeholder captions
    lblUsername.Caption = CurrentUser.username
    lblStaffId.Caption = CurrentUser.id
    lblAccountUsername.Caption = lblUsername.Caption
    
    EnrolleeLV.FullRowSelect = True
    
    ' Add listview column headers
    With EnrolleeLV.ColumnHeaders
        .Add , , "Id", 500
        .Add , , "Grade", 900, lvwColumnCenter
        .Add , , "Last Name", 1200, lvwColumnCenter
        .Add , , "First Name", 1200, lvwColumnCenter
        .Add , , "Middle Name", 1200, lvwColumnCenter
        .Add , , "Sex", 700, lvwColumnCenter
        .Add , , "Age", 700, lvwColumnCenter
        .Add , , "Birthdate", 1200, lvwColumnCenter
        .Add , , "Birthplace", 1200, lvwColumnCenter
        .Add , , "Mother Toungue", 1400, lvwColumnCenter
        .Add , , "Address", 1200, lvwColumnCenter
        .Add , , "Date Enrolled", 1200, lvwColumnCenter
        .Add , , "Father Name", 1200, lvwColumnCenter
        .Add , , "Father No.", 1200, lvwColumnCenter
        .Add , , "Mother Name", 1200, lvwColumnCenter
        .Add , , "Mother No.", 1200, lvwColumnCenter
        .Add , , "Guardian Name", 1200, lvwColumnCenter
        .Add , , "Guardian No.", 1200, lvwColumnCenter
    End With
    
    current_page = 1
    Set result = Get_Enrollee()
    Call InitPagination(result)
End Sub
