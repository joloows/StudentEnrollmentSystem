VERSION 5.00
Begin VB.Form StudentForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10515
   ClientLeft      =   6810
   ClientTop       =   1410
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   9975
   Begin VB.CommandButton Command3 
      Caption         =   "Show Staff Form"
      Height          =   495
      Left            =   240
      TabIndex        =   68
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Submit"
      Height          =   495
      Left            =   4200
      TabIndex        =   66
      Top             =   9720
      Width           =   1695
   End
   Begin VB.TextBox Text28 
      Height          =   285
      Left            =   7440
      TabIndex        =   64
      Top             =   9000
      Width           =   1935
   End
   Begin VB.TextBox Text27 
      Height          =   285
      Left            =   600
      TabIndex        =   60
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox Text26 
      Height          =   285
      Left            =   2880
      TabIndex        =   59
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   5160
      TabIndex        =   58
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   7440
      TabIndex        =   55
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   7440
      TabIndex        =   53
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   600
      TabIndex        =   49
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   2880
      TabIndex        =   48
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   5160
      TabIndex        =   47
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   600
      TabIndex        =   42
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   2880
      TabIndex        =   41
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   5160
      TabIndex        =   40
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   8640
      TabIndex        =   37
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   6960
      TabIndex        =   36
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   5280
      TabIndex        =   35
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   3600
      TabIndex        =   34
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1920
      TabIndex        =   33
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   600
      TabIndex        =   32
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   7920
      TabIndex        =   24
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   8640
      TabIndex        =   22
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5520
      TabIndex        =   20
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   4800
      TabIndex        =   16
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4200
      TabIndex        =   15
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3600
      TabIndex        =   14
      Top             =   3720
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sex"
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   3360
      Width           =   2175
      Begin VB.OptionButton Option2 
         Caption         =   "Female"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Male"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7440
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
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
      TabIndex        =   63
      Top             =   8640
      Width           =   810
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   2880
      TabIndex        =   62
      Top             =   8640
      Width           =   795
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5160
      TabIndex        =   61
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
      TabIndex        =   57
      Top             =   8280
      Width           =   1515
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Contact Number:"
      Height          =   195
      Left            =   7440
      TabIndex        =   56
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Contact Number:"
      Height          =   195
      Left            =   7440
      TabIndex        =   54
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   52
      Top             =   7440
      Width           =   810
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   2880
      TabIndex        =   51
      Top             =   7440
      Width           =   795
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5160
      TabIndex        =   50
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
      TabIndex        =   46
      Top             =   7080
      Width           =   2010
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   45
      Top             =   6240
      Width           =   810
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   2880
      TabIndex        =   44
      Top             =   6240
      Width           =   795
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5160
      TabIndex        =   43
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      TabIndex        =   31
      Top             =   1800
      Width           =   10335
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "ZIP code:"
      Height          =   195
      Left            =   8640
      TabIndex        =   30
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Province:"
      Height          =   195
      Left            =   6960
      TabIndex        =   29
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Municipality/City:"
      Height          =   195
      Left            =   5280
      TabIndex        =   28
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Barangay:"
      Height          =   195
      Left            =   3600
      TabIndex        =   27
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Street Name:"
      Height          =   195
      Left            =   1920
      TabIndex        =   26
      Top             =   4200
      Width           =   930
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "House No.:"
      Height          =   195
      Left            =   600
      TabIndex        =   25
      Top             =   4200
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Mother Tongue:"
      Height          =   195
      Left            =   7920
      TabIndex        =   23
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
      TabIndex        =   21
      Top             =   600
      Width           =   5355
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Place of birth (Municipality/City):"
      Height          =   195
      Left            =   5520
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   3720
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Birthdate (mm/dd/yyyy):"
      Height          =   195
      Left            =   3600
      TabIndex        =   13
      Top             =   3360
      Width           =   1680
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Age:"
      Height          =   195
      Left            =   3000
      TabIndex        =   8
      Top             =   3360
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Extension Name: e.g. Jr,  III"
      Height          =   195
      Left            =   7440
      TabIndex        =   7
      Top             =   2520
      Width           =   1950
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5160
      TabIndex        =   6
      Top             =   2520
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   2880
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

Private Sub Command1_Click()
    LoginForm.Show
End Sub

Private Sub Command3_Click()
    StaffForm.Show
End Sub

