VERSION 5.00
Begin VB.Form EnSelectForm 
   Caption         =   "Manage Enrollee"
   ClientHeight    =   4200
   ClientLeft      =   7860
   ClientTop       =   4905
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   8505
   Begin VB.CommandButton btnUpdateFee 
      Caption         =   "Update Fees"
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtAddPayment 
      Height          =   285
      Left            =   6480
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
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
   Begin VB.Label Label7 
      Caption         =   "Add payment: "
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblEnFeeLeft 
      Alignment       =   1  'Right Justify
      Caption         =   "Label6"
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Fees Left: "
      Height          =   195
      Left            =   4680
      TabIndex        =   13
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label lblEnTotalFee 
      Alignment       =   1  'Right Justify
      Caption         =   "Label6"
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Total Fees:"
      Height          =   195
      Left            =   4680
      TabIndex        =   11
      Top             =   1080
      Width           =   795
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
      Height          =   495
      Left            =   2040
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
    
    StudentForm.Hide
    
    Address = Split(En.Address, ", ")
    x = Split(Address(UBound(Address)), " ")
    zip = x(UBound(x))
    Fathername = Split(En.Fathername, " ")
    MotherName = Split(En.MotherName, " ")
    GuardianName = Split(En.GuardianName, " ")
    With StudentForm
        .txtLname.Text = En.Lname
        .txtFname.Text = En.Fname
        .txtMname.Text = En.Mname
        .cmbGrade.Text = En.Grade
        .optMale.Value = IIf(En.Sex = "Male", True, False)
        .optFemale.Value = IIf(En.Sex = "Female", True, False)
        .txtAge.Text = En.Age
        .dtBirthdate = En.Birthdate
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
    
    StudentForm.Width = 9960
    StudentForm.inputMode = 1
    StudentForm.Show vbModal
    StudentForm.Width = 13995
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

Private Sub btnUpdateFee_Click()
    On Err GoTo FeeErrorHandler
    Dim payment As Currency
    Dim response As Integer
    
    If lblEnFeeLeft.Caption = 0 Then
        GoTo EnrolleeAlreadyPaid
    End If
    
    payment = CCur(txtAddPayment.Text)
    
    If payment > CCur(lblEnFeeLeft) Then
        payment = CCur(lblEnFeeLeft)
    End If
    
    response = AddEnrolleeFeePayment(payment, En.id)
    If response = 0 Then
        MsgBox "Fee has been updated succesfully.", vbInformation, "Success"
        If payment = CCur(lblEnFeeLeft) Then
            lblEnFeeLeft.Caption = 0
        Else
            lblEnFeeLeft.Caption = CInt(lblEnFeeLeft.Caption) - payment
        End If
        txtAddPayment.Text = ""
    Else
        MsgBox "An Unexpected error has occured.", vbCritical, "Error"
    End If
    
    Set result = GetEnrollee(StaffForm.eCurrentPage, StaffForm.search)
    Call StaffForm.InitPagination("enrollee", result)
    Exit Sub
FeeErrorHandler:
    MsgBox "Invalid Input. Please only use numbers (0-9) in the fields.", vbCritical, "Error"
    txtAddPayment.Text = ""
    Exit Sub
EnrolleeAlreadyPaid:
    MsgBox "Enrollee is already fully paid.", vbExclamation, "Fully paid"
End Sub

Private Sub Form_Load()
    Set En = StaffForm.selectedEnrollee

    MI = UCase(mId(En.Mname, 1, 1))
    chkEnrolled.Value = IIf(En.Enrolled, 1, 0)
    lblSelected.Caption = En.Fname & " " & MI & ". " & En.Lname
    lblGrade.Caption = En.Grade
    lblSection.Caption = IIf(En.Section <> "", En.Section, "not assigned")
    lblEnTotalFee.Caption = En.TotalFee
    Debug.Print En.TotalFee & " - " & En.payment
    lblEnFeeLeft.Caption = En.TotalFee - En.payment
End Sub

