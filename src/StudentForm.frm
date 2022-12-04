VERSION 5.00
Begin VB.Form StudentForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AES Enrollment System"
   ClientHeight    =   10515
   ClientLeft      =   6810
   ClientTop       =   1410
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   9975
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   8640
      TabIndex        =   68
      Top             =   9720
      Width           =   735
   End
   Begin VB.CommandButton btnSubmitRecord 
      Caption         =   "Submit"
      Height          =   495
      Left            =   4200
      TabIndex        =   66
      Top             =   9720
      Width           =   1695
   End
   Begin VB.TextBox txtgNum 
      Height          =   285
      Left            =   7440
      TabIndex        =   31
      Top             =   9000
      Width           =   1935
   End
   Begin VB.TextBox txtgLname 
      Height          =   285
      Left            =   600
      TabIndex        =   28
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox txtgFname 
      Height          =   285
      Left            =   2880
      TabIndex        =   29
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox txtgMname 
      Height          =   285
      Left            =   5160
      TabIndex        =   30
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox txtmNum 
      Height          =   285
      Left            =   7440
      TabIndex        =   27
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox txtfNum 
      Height          =   285
      Left            =   7440
      TabIndex        =   23
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox txtmLname 
      Height          =   285
      Left            =   600
      TabIndex        =   24
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox txtmFname 
      Height          =   285
      Left            =   2880
      TabIndex        =   25
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox txtmMname 
      Height          =   285
      Left            =   5160
      TabIndex        =   26
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox txtfLname 
      Height          =   285
      Left            =   600
      TabIndex        =   20
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtfFname 
      Height          =   285
      Left            =   2880
      TabIndex        =   21
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtfMname 
      Height          =   285
      Left            =   5160
      TabIndex        =   22
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   8640
      TabIndex        =   19
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtProv 
      Height          =   285
      Left            =   6960
      TabIndex        =   18
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   5280
      TabIndex        =   17
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtBrgy 
      Height          =   285
      Left            =   3600
      TabIndex        =   16
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtSt 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtHno 
      Height          =   285
      Left            =   600
      TabIndex        =   14
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtMt 
      Height          =   285
      Left            =   7920
      TabIndex        =   13
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   8640
      TabIndex        =   41
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtBirth 
      Height          =   285
      Left            =   5520
      TabIndex        =   12
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtBy 
      Height          =   285
      Left            =   4800
      TabIndex        =   11
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtBd 
      Height          =   285
      Left            =   4200
      TabIndex        =   10
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtBm 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   3720
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sex"
      Height          =   615
      Left            =   600
      TabIndex        =   35
      Top             =   3360
      Width           =   2175
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtAge 
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtMname 
      Height          =   285
      Left            =   5640
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtFname 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txtEx 
      Height          =   285
      Left            =   8160
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtLname 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "Student Enrollment System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   67
      Top             =   1320
      Width           =   2910
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "Contact Number:"
      Height          =   195
      Left            =   7440
      TabIndex        =   65
      Top             =   8640
      Width           =   1200
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   64
      Top             =   8640
      Width           =   810
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   2880
      TabIndex        =   63
      Top             =   8640
      Width           =   795
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5160
      TabIndex        =   62
      Top             =   8640
      Width           =   945
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "Guardian's Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   61
      Top             =   8280
      Width           =   1515
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Contact Number:"
      Height          =   195
      Left            =   7440
      TabIndex        =   60
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Contact Number:"
      Height          =   195
      Left            =   7440
      TabIndex        =   59
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   58
      Top             =   7440
      Width           =   810
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   2880
      TabIndex        =   57
      Top             =   7440
      Width           =   795
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5160
      TabIndex        =   56
      Top             =   7440
      Width           =   945
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Mother's Maiden Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   55
      Top             =   7080
      Width           =   2010
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   54
      Top             =   6240
      Width           =   810
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   2880
      TabIndex        =   53
      Top             =   6240
      Width           =   795
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5160
      TabIndex        =   52
      Top             =   6240
      Width           =   945
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Father's Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   51
      Top             =   5880
      Width           =   1290
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parent/Guardian Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -120
      TabIndex        =   50
      Top             =   5160
      Width           =   10335
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Learner Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -120
      TabIndex        =   49
      Top             =   1800
      Width           =   10335
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "ZIP code:"
      Height          =   195
      Left            =   8640
      TabIndex        =   48
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Province:"
      Height          =   195
      Left            =   6960
      TabIndex        =   47
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Municipality/City:"
      Height          =   195
      Left            =   5280
      TabIndex        =   46
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Barangay:"
      Height          =   195
      Left            =   3600
      TabIndex        =   45
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Street Name:"
      Height          =   195
      Left            =   1920
      TabIndex        =   44
      Top             =   4200
      Width           =   930
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "House No.:"
      Height          =   195
      Left            =   600
      TabIndex        =   43
      Top             =   4200
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Mother Tongue:"
      Height          =   195
      Left            =   7920
      TabIndex        =   42
      Top             =   3360
      Width           =   1140
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "School Logo, Name, details here "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   40
      Top             =   600
      Width           =   5355
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Place of birth (Municipality/City):"
      Height          =   195
      Left            =   5520
      TabIndex        =   39
      Top             =   3360
      Width           =   2265
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4680
      TabIndex        =   38
      Top             =   3720
      Width           =   60
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   37
      Top             =   3720
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Birthdate (mm/dd/yyyy):"
      Height          =   195
      Left            =   3600
      TabIndex        =   36
      Top             =   3360
      Width           =   1680
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Age:"
      Height          =   195
      Left            =   3000
      TabIndex        =   34
      Top             =   3360
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Extension Name:"
      Height          =   195
      Left            =   8160
      TabIndex        =   33
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5640
      TabIndex        =   32
      Top             =   2520
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   3000
      TabIndex        =   1
      Top             =   2520
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   2520
      Width           =   810
   End
End
Attribute VB_Name = "StudentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClear_Click()
    Dim En As putangina
    
    Debug.Print TypeName(En)
End Sub

Private Sub Command1_Click()
    If CurrentUser.isAuthenticated Then
        StaffForm.Show
    Else
        LoginForm.Show
    End If
End Sub

Private Sub Command3_Click()
    StaffForm.Show
End Sub

Private Sub btnSubmitRecord_Click()
    Dim En As Enrollee
    
    En.Lname = txtLname.Text
    En.Fname = txtFname.Text
    En.Mname = txtMname.Text
    En.Ex = txtEx.Text
    
    If optMale Then
        En.Sex = optMale.Caption
    ElseIf optFemale Then
        En.Sex = optFemale.Caption
    End If
    
    En.Age = CInt(txtAge.Text)
    En.Birthdate = CDate(txtBm.Text & "/" & txtBd.Text & "/" & txtBy.Text)
    En.Birthplace = txtBirth.Text
    En.Mt = txtMt.Text
    En.Address = txtHno.Text & "," & txtSt.Text & "," & txtBrgy.Text & "," & txtCity.Text & "," & txtProv.Text & " " & txtZip.Text
    En.FatherName = txtfFname.Text & " " & txtfMname.Text & " " & txtfLname.Text
    En.FNum = txtfNum.Text
    En.MotherName = txtmFname.Text & " " & txtmMname.Text & " " & txtmLname.Text
    En.Mnum = txtmNum.Text
    En.GuardianName = txtgFname.Text & " " & txtgMname.Text & " " & txtgLname.Text
    En.Gnum = txtgNum.Text
    En.Submission = Format(Now, "mm/dd/yyyy")
    
    Debug.Print En.Lname
    Call AddEnrollee(En)
    
    
End Sub

Private Sub ClearForm()
    txtLname.Text = ""
    txtFname.Text = ""
    txtMname.Text = ""
    optMale.Value = False
    optFemale.Value = False
    txtAge.Text = ""
    txtBm.Text = ""
    txtBd.Text = ""
    txtBy.Text = ""
    txtBirth.Text = ""
    txtMt.Text = ""
    txtHno.Text = ""
    txtSt.Text = ""
    txtBrgy.Text = ""
    txtCity.Text = ""
    txtProv.Text = ""
    txtZip.Text = ""
    txtfLname.Text = ""
    txtfFname.Text = ""
    txtfMname.Text = ""
    txtfNum.Text = ""
    txtmLname.Text = ""
    txtmFname.Text = ""
    txtmMname.Text = ""
    txtmNum.Text = ""
    txtgLname.Text = ""
    txtgFname.Text = ""
    txtgMname.Text = ""
    txtgNum.Text = ""
End Sub


