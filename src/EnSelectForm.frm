VERSION 5.00
Begin VB.Form EnSelectForm 
   Caption         =   "Manage Enrollee"
   ClientHeight    =   4200
   ClientLeft      =   9210
   ClientTop       =   4950
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4575
   Begin VB.CommandButton btnAssignSection 
      Caption         =   "Assign Section"
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CheckBox chkEnrolled 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton btnEnDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton btnEnUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Enrolled:"
      Height          =   195
      Left            =   1680
      TabIndex        =   8
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "<section name>"
      Height          =   195
      Left            =   3060
      TabIndex        =   5
      Top             =   2040
      Width           =   1155
   End
   Begin VB.Label lblGrade 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "<num>"
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Section:"
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Grade Level: "
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblSelected 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "<full name of selected enrollee>"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   2385
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Manage Enrollee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "EnSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private En As Enrollee
Private result As Collection

Private Sub btnEnDelete_Click()

    x = MsgBox("Are you sure you want to delete this enrollee?", vbQuestion + vbYesNo, "Confirm")
    If x = 6 Then
        Call DeleteEnrollee(StaffForm.selectedEnrollee.id)
        MsgBox "Enrollee is successfuly deleted.", vbInformation, "Success"
        
        Set result = GetEnrollee(StaffForm.eCurrentPage, StaffForm.search)
        Call StaffForm.InitPagination("enrollee", result)
        
        Unload Me
    End If
End Sub

Private Sub btnEnUpdate_Click()
    Dim Address() As String
    Dim x() As String
    Dim zip As String
    Dim Fathername() As String
    Dim MotherName() As String
    Dim GuardianName() As String
    Dim mop As Integer
    
    StudentForm.Hide
    
    
    If En.Tuition = 11550 Then
        mop = 2
    ElseIf En.Tuition = 11025 Then
        mop = 1
    Else
        mop = 0
    End If
    
    Address = Split(En.Address, ", ")
    x = Split(Address(UBound(Address)), " ")
    zip = x(UBound(x))
    Fathername = Split(En.Fathername, " ")
    MotherName = Split(En.MotherName, " ")
    GuardianName = Split(En.GuardianName, " ")
    With StudentForm
        .cbxMOP.ListIndex = mop
        .txtLname.Text = En.Lname
        .txtFname.Text = En.Fname
        .txtMname.Text = En.Mname
        .txtGrade.Text = En.Grade
        .optMale.Value = IIf(En.Sex = "Male", True, False)
        .optFemale.Value = IIf(En.Sex = "Female", True, False)
        .txtAge.Text = En.Age
        .txtBm.Text = Month(En.Birthdate)
        .txtBd.Text = Day(En.Birthdate)
        .txtBy.Text = Year(En.Birthdate)
        .txtBirth.Text = En.Birthplace
        .txtMt.Text = En.Mt
        .txtHno.Text = Address(0)
        .txtSt.Text = Address(1)
        .txtBrgy.Text = Address(2)
        .txtCity.Text = Address(3)
        .txtProv.Text = x(0)
        .txtZip.Text = zip
        .txtfLname.Text = Fathername(0)
        .txtfFname.Text = Fathername(2)
        .txtfMname.Text = Fathername(1)
        .txtfNum.Text = En.Fnum
        .txtmLname.Text = MotherName(2)
        .txtmFname.Text = MotherName(0)
        .txtmMname.Text = MotherName(1)
        .txtmNum.Text = En.Mnum
        .txtgLname.Text = GuardianName(2)
        .txtgFname.Text = GuardianName(0)
        .txtgMname.Text = GuardianName(1)
        .txtgNum.Text = En.Gnum
    End With
    

    StudentForm.inputMode = 1
    StudentForm.Show vbModal
    StudentForm.inputMode = 0
End Sub

' fix error handler
Private Sub chkEnrolled_Click()
    Dim status As Boolean
    
    
    status = CBool(chkEnrolled.Value)
    Call UpdateEnrolleeStatus(status, StaffForm.selectedEnrollee.id)
    
    Set result = GetEnrollee(StaffForm.eCurrentPage, StaffForm.search)
    Call StaffForm.InitPagination("enrollee", result)
End Sub

Private Sub btnAssignSection_Click()
    Dim Section As String
    
    Section = StrConv(InputBox("Enter the name of the section:", "Assign Section"), vbProperCase)
    
    Call AssignEnrolleeSection(Section, StaffForm.selectedEnrollee.id)
    
    lblSection.Caption = Section
    
    Set result = GetEnrollee(StaffForm.eCurrentPage, StaffForm.search)
    Call StaffForm.InitPagination("enrollee", result)
End Sub

Private Sub Form_Load()
    Set En = StaffForm.selectedEnrollee

    MI = UCase(mId(En.Mname, 1, 1))
    chkEnrolled.Value = IIf(En.Enrolled, 1, 0)
    lblSelected.Caption = En.Fname & " " & MI & ". " & En.Lname
    lblGrade.Caption = En.Grade
    lblSection.Caption = IIf(En.Section <> "", En.Section, "not assigned")
End Sub
