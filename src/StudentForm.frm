VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form StudentForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AES Enrollment System"
   ClientHeight    =   10515
   ClientLeft      =   5145
   ClientTop       =   1590
   ClientWidth     =   13905
   Icon            =   "StudentForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleMode       =   0  'User
   ScaleWidth      =   13995
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   8160
      TabIndex        =   87
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtDownPayment 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   84
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton calculateFee 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   11400
      TabIndex        =   80
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txtPayment 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   78
      Top             =   5400
      Width           =   735
   End
   Begin VB.OptionButton optDown 
      Caption         =   "Down payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   77
      Top             =   4080
      Width           =   1815
   End
   Begin VB.OptionButton optFull 
      Caption         =   "Full"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   76
      Top             =   4080
      Width           =   855
   End
   Begin VB.CheckBox chkUniform 
      Caption         =   "Check1"
      Height          =   255
      Left            =   10320
      TabIndex        =   70
      Top             =   2550
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtBirthdate 
      Height          =   285
      Left            =   3600
      TabIndex        =   67
      Top             =   3720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   503
      _Version        =   393216
      Format          =   109838337
      CurrentDate     =   40179
   End
   Begin VB.ComboBox cmbGrade 
      Height          =   315
      ItemData        =   "StudentForm.frx":2EAEE
      Left            =   8160
      List            =   "StudentForm.frx":2EB04
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   360
      Picture         =   "StudentForm.frx":2EB1A
      ScaleHeight     =   1575
      ScaleWidth      =   1695
      TabIndex        =   63
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   8640
      TabIndex        =   62
      Top             =   9720
      Width           =   735
   End
   Begin VB.CommandButton btnSubmitRecord 
      Caption         =   "Submit"
      Height          =   495
      Left            =   4200
      TabIndex        =   60
      Top             =   9720
      Width           =   1695
   End
   Begin VB.TextBox txtgNum 
      Height          =   285
      Left            =   7440
      TabIndex        =   27
      Top             =   9000
      Width           =   1935
   End
   Begin VB.TextBox txtgLname 
      Height          =   285
      Left            =   600
      TabIndex        =   24
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox txtgFname 
      Height          =   285
      Left            =   2880
      TabIndex        =   25
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox txtgMname 
      Height          =   285
      Left            =   5160
      TabIndex        =   26
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox txtmNum 
      Height          =   285
      Left            =   7440
      TabIndex        =   23
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox txtfNum 
      Height          =   285
      Left            =   7440
      TabIndex        =   19
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox txtmLname 
      Height          =   285
      Left            =   600
      TabIndex        =   20
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox txtmFname 
      Height          =   285
      Left            =   2880
      TabIndex        =   21
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox txtmMname 
      Height          =   285
      Left            =   5160
      TabIndex        =   22
      Top             =   7800
      Width           =   2055
   End
   Begin VB.TextBox txtfLname 
      Height          =   285
      Left            =   600
      TabIndex        =   16
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtfFname 
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtfMname 
      Height          =   285
      Left            =   5160
      TabIndex        =   18
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   8640
      TabIndex        =   15
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtProv 
      Height          =   285
      Left            =   6960
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   5280
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtBrgy 
      Height          =   285
      Left            =   3600
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtSt 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtHno 
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtMt 
      Height          =   285
      Left            =   7920
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton btnLgn 
      Caption         =   "Login"
      Height          =   375
      Left            =   8640
      TabIndex        =   35
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtBirth 
      Height          =   285
      Left            =   5520
      TabIndex        =   8
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sex"
      Height          =   615
      Left            =   600
      TabIndex        =   31
      Top             =   3360
      Width           =   2175
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtAge 
      Height          =   285
      Left            =   3000
      TabIndex        =   7
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
   Begin VB.TextBox txtLname 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblFeeLeft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fee Left Here"
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
      Left            =   11895
      TabIndex        =   86
      Top             =   7560
      Width           =   1485
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Fee Left: "
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
      Left            =   10290
      TabIndex        =   85
      Top             =   7560
      Width           =   1035
   End
   Begin VB.Label lblDownPayment 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Down Payment:"
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
      Left            =   10320
      TabIndex        =   83
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblChange 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Change here"
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
      Left            =   12000
      TabIndex        =   82
      Top             =   6840
      Width           =   1380
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Change: "
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
      Left            =   10320
      TabIndex        =   81
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label lblPayment 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Payment:"
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
      Left            =   10320
      TabIndex        =   79
      Top             =   5400
      Width           =   1005
   End
   Begin VB.Label lblTotalFee 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "400"
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
      Left            =   12240
      TabIndex        =   75
      Top             =   3360
      Width           =   435
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total:"
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
      Left            =   10800
      TabIndex        =   74
      Top             =   3360
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   10266.02
      X2              =   13768.54
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calculate Fees"
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
      Left            =   9960
      TabIndex        =   73
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "ID Card:           100"
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
      Left            =   10740
      TabIndex        =   72
      Top             =   1920
      Width           =   1965
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Uniform:          400"
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
      Left            =   10775
      TabIndex        =   71
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Tuition:            300"
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
      Left            =   10800
      TabIndex        =   69
      Top             =   1320
      Width           =   1905
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Calculate Tuition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   68
      Top             =   360
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   10024.47
      X2              =   10024.47
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      Caption         =   "Kap. M. S. Victa, Kawit, Philippines, 4104"
      Height          =   195
      Left            =   2640
      TabIndex        =   65
      Top             =   960
      Width           =   4875
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Caption         =   "depedcavite.aguinadoes108016@gmail.com |  (046) 484 7623"
      Height          =   195
      Left            =   2640
      TabIndex        =   64
      Top             =   720
      Width           =   4890
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
      TabIndex        =   61
      Top             =   1320
      Width           =   2910
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "Contact Number:"
      Height          =   195
      Left            =   7440
      TabIndex        =   59
      Top             =   8640
      Width           =   1200
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   58
      Top             =   8640
      Width           =   810
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   2880
      TabIndex        =   57
      Top             =   8640
      Width           =   795
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5160
      TabIndex        =   56
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
      TabIndex        =   55
      Top             =   8280
      Width           =   1515
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Contact Number:"
      Height          =   195
      Left            =   7440
      TabIndex        =   54
      Top             =   7440
      Width           =   1200
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Contact Number:"
      Height          =   195
      Left            =   7440
      TabIndex        =   53
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
      TabIndex        =   49
      Top             =   7080
      Width           =   2010
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   48
      Top             =   6240
      Width           =   810
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Left            =   2880
      TabIndex        =   47
      Top             =   6240
      Width           =   795
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5160
      TabIndex        =   46
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
      TabIndex        =   45
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
      TabIndex        =   44
      Top             =   5160
      Width           =   10095
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
      TabIndex        =   43
      Top             =   1800
      Width           =   10095
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "ZIP code:"
      Height          =   195
      Left            =   8640
      TabIndex        =   42
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Province:"
      Height          =   195
      Left            =   6960
      TabIndex        =   41
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Municipality/City:"
      Height          =   195
      Left            =   5280
      TabIndex        =   40
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Barangay:"
      Height          =   195
      Left            =   3600
      TabIndex        =   39
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Street Name:"
      Height          =   195
      Left            =   1920
      TabIndex        =   38
      Top             =   4200
      Width           =   930
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "House No.:"
      Height          =   195
      Left            =   600
      TabIndex        =   37
      Top             =   4200
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Mother Tongue:"
      Height          =   195
      Left            =   7920
      TabIndex        =   36
      Top             =   3360
      Width           =   1140
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Aguinaldo Elementary School"
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
      Left            =   2640
      TabIndex        =   34
      Top             =   240
      Width           =   4860
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Place of birth (Municipality/City):"
      Height          =   195
      Left            =   5520
      TabIndex        =   33
      Top             =   3360
      Width           =   2265
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Birthdate (mm/dd/yyyy):"
      Height          =   195
      Left            =   3600
      TabIndex        =   32
      Top             =   3360
      Width           =   1680
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Age:"
      Height          =   195
      Left            =   3000
      TabIndex        =   30
      Top             =   3360
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Grade Level"
      Height          =   195
      Left            =   8160
      TabIndex        =   29
      Top             =   2520
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Middle name:"
      Height          =   195
      Left            =   5640
      TabIndex        =   28
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
' For determining the action to be taken by btnSubmitRecord_Click()
' 0 for create, 1 for update
Public inputMode As Integer
Private feeHasBeenCalculated As Boolean

Private Sub btnClear_Click()
    Call ClearForm
End Sub

Private Sub btnLgn_Click()
    If CurrentUser.isAuthenticated Then
        StaffForm.Show
    Else
        LoginForm.Show vbModal
    End If
End Sub

Private Sub btnSubmitRecord_Click()
    On Error GoTo Handler
    Dim En As New Enrollee
    Dim result As Collection
    Dim s As String
    Dim totalTuition
    
    If optMale Then
        s = optMale.Caption
    ElseIf optFemale Then
        s = optFemale.Caption
    End If
    
    If inputMode = 1 Then 'update enrollee
        
        x = MsgBox("Are you sure the updated details are correct?", vbYesNo + vbExclamation, "Confirm")
        
        If x = 6 Then
            
            With En
                .Lname = txtLname.Text
                .Fname = txtFname.Text
                .Mname = txtMname.Text
                .Grade = CInt(cmbGrade.Text)
                .Sex = s
                .Age = CInt(txtAge.Text)
                .Birthdate = Format(dtBirthdate.Value, "mm/dd/yyyy")
                .Birthplace = txtBirth.Text
                .Mt = txtMt.Text
                .Address = txtHno.Text & ", " & txtSt.Text & ", " & txtBrgy.Text & ", " & txtCity.Text & ", " & txtProv.Text & " " & txtZip.Text
                .Fathername = txtfFname.Text & " " & txtfMname.Text & " " & txtfLname.Text
                .Fnum = txtfNum.Text
                .MotherName = txtmFname.Text & " " & txtmMname.Text & " " & txtmLname.Text
                .Mnum = txtmNum.Text
                .GuardianName = txtgFname.Text & " " & txtgMname.Text & " " & txtgLname.Text
                .Gnum = txtgNum.Text
            End With
            
            Call UpdateEnrollee(StaffForm.selectedEnrollee.id, En)
            
            Set result = GetEnrollee(StaffForm.eCurrentPage, StaffForm.search)
            Call StaffForm.InitPagination("enrollee", result)
            
            Unload EnSelectForm
            Unload Me
        End If
        
    Else 'create enrollee
        
        x = MsgBox("Are you sure the details are correct?", vbYesNo + vbExclamation, "Confirm")

        ' If yes, Add record to database
        If x = 6 Then
            If feeHasBeenCalculated Then
                With En
                    .Lname = txtLname.Text
                    .Fname = txtFname.Text
                    .Mname = txtMname.Text
                    .Grade = CInt(cmbGrade.Text)
                    .TotalFee = CInt(lblTotalFee.Caption)
                    .WithUniform = chkUniform.Value
                    .PaymentType = IIf(optFull, "full", "down_payment")
                    .payment = CInt(txtPayment.Text) - CInt(lblChange.Caption)
                    .Sex = s
                    .Age = CInt(txtAge.Text)
                    .Birthdate = Format(dtBirthdate.Value, "mm/dd/yyyy")
                    .Birthplace = txtBirth.Text
                    .Mt = txtMt.Text
                    .Address = txtHno.Text & ", " & txtSt.Text & ", " & txtBrgy.Text & ", " & txtCity.Text & ", " & txtProv.Text & " " & txtZip.Text
                    .Fathername = txtfFname.Text & " " & txtfMname.Text & " " & txtfLname.Text
                    .Fnum = txtfNum.Text
                    .MotherName = txtmFname.Text & " " & txtmMname.Text & " " & txtmLname.Text
                    .Mnum = txtmNum.Text
                    .GuardianName = txtgFname.Text & " " & txtgMname.Text & " " & txtgLname.Text
                    .Gnum = txtgNum.Text
                    .Submission = Format(Now, "mm/dd/yyyy")
                End With
            Else
                GoTo FeeHasNotBeenCalculatedHandler
            End If
            
            Call AddEnrollee(En)
        End If
    End If
    Exit Sub
Handler:
    If Err.Number = 13 Then
        MsgBox "Input mismatch. Please only only put numbers in age or birthdate input fields.", vbCritical, "Error"
    End If
    Exit Sub
FeeHasNotBeenCalculatedHandler:
    MsgBox "Fee has not been calculated.", vbCritical, "Fee not calculated"
End Sub

Private Sub ClearForm()
    txtLname.Text = ""
    txtFname.Text = ""
    txtMname.Text = ""
    optMale.Value = True
    optFemale.Value = False
    txtAge.Text = ""
    dtBirthdate.Value = "01/01/2010"
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
    txtDownPayment.Text = ""
    txtPayment.Text = ""
    optFull.Value = True
    optDown.Value = False
    chkUniform.Value = 0
    lblChange.Caption = ""
    lblFeeLeft.Caption = ""
    feeHasBeenCalculated = False
End Sub

Private Sub calculateFee_Click()
    On Error GoTo FeeErrorHandler
    Dim downPayment As Integer
    Dim payment As Integer
    Dim change As Currency
    Dim feeLeft As Currency
    Dim TotalFee As Integer
    
    downPayment = 0
    
    TotalFee = CInt(lblTotalFee.Caption)
    If optDown Then
        downPayment = CInt(txtDownPayment.Text)
    End If
    payment = CInt(txtPayment.Text)
    
    If optFull Then
        If payment < TotalFee Then
            GoTo PaymentNotEnoughHandler
            Exit Sub
        End If
        change = Abs(TotalFee - payment)
        feeLeft = 0
    ElseIf optDown Then
        If payment < downPayment Then
            GoTo PaymentNotEnoughHandler
            Exit Sub
        End If
        change = Abs(downPayment - payment)
        feeLeft = TotalFee - downPayment
    End If
    
    lblChange.Caption = change
    lblFeeLeft.Caption = feeLeft
    
    feeHasBeenCalculated = True
    Exit Sub
    
FeeErrorHandler:
    MsgBox "Invalid Input. Please only use numbers (0-9) in the fields.", vbCritical, "Error"
    txtDownPayment.Text = ""
    txtPayment.Text = ""
    Exit Sub
    
PaymentNotEnoughHandler:
    MsgBox "Payment should be equal or over the total fee/down payment.", vbExclamation, "Error"
    txtDownPayment.Text = ""
    txtPayment.Text = ""
End Sub

Private Sub chkUniform_Click()
    txtDownPayment.Text = ""
    If chkUniform.Value = 1 Then
        Label41.Enabled = True
        lblTotalFee.Caption = "800"
    ElseIf chkUniform.Value = 0 Then
        Label41.Enabled = False
        lblTotalFee.Caption = "400"
    End If
End Sub

Private Sub Command1_Click()
    StaffForm.Show
End Sub

Private Sub Form_Load()
    Picture1.Picture = LoadPicture(App.Path & "\aes-ico.jpg")
    
    Picture1.ScaleMode = 3
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture1.Picture, _
    0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    0, 0, Picture1.Picture.Width / 26.46, _
    Picture1.Picture.Height / 26.46
    
    Picture1.Picture = Picture1.Image
    
    P = App.Path & "\hand.ico"
    
    lblTotalFee.Caption = "400"
    Label41.Enabled = False
    optFull.Value = True
    optMale.Value = True
    lblChange.Caption = ""
    lblFeeLeft.Caption = ""
    
    ' Temporary autofill student form
    txtLname.Text = "Antonio"
    txtFname.Text = "Angelo"
    txtMname.Text = "Bueneventura"
    cmbGrade.Text = "5"
    optMale.Value = True
    optFemale.Value = False
    txtAge.Text = "10"
    txtBirth.Text = "Cavite City"
    txtMt.Text = "Filipino"
    txtHno.Text = "101"
    txtSt.Text = "Maharlika St."
    txtBrgy.Text = "Poblacion"
    txtCity.Text = "Kawit"
    txtProv.Text = "Marulas"
    txtZip.Text = "1031"
    txtfLname.Text = "Antonio"
    txtfFname.Text = "Denzel"
    txtfMname.Text = "Velcuz"
    txtfNum.Text = "09222222222"
    txtmLname.Text = "Bueneventura"
    txtmFname.Text = "Stephanie"
    txtmMname.Text = "Villuaneva"
    txtmNum.Text = "09333333333"
    txtgLname.Text = "Denzel"
    txtgFname.Text = "Antonio"
    txtgMname.Text = "Velcus"
    txtgNum.Text = "09222222222"
End Sub

Private Sub Label39_Click()
    tuitionDialog.Show vbModal
End Sub

Private Sub optDown_Click()
    lblDownPayment.Visible = True
    txtDownPayment.Visible = True
End Sub

Private Sub optFull_Click()
    lblDownPayment.Visible = False
    txtDownPayment.Visible = False
End Sub

Private Sub txtDownPayment_Change()
    On Error GoTo ErrorHandler
    If txtDownPayment <> "" Then
        If CInt(txtDownPayment) > lblTotalFee Then
            MsgBox "Entered down payment exceeds the total fee. Please choose the 'full' option if you plan to pay the fee in whole.", vbExclamation, "Error"
            txtDownPayment.Text = ""
        End If
    End If
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub txtPayment_Change()
    On Error GoTo ErrorHandler
    If txtPayment <> "" Then
        If CInt(txtPayment) > 1000 Then
            MsgBox "Valid payment does not exceed 1000.", vbExclamation, "Error"
            txtPayment.Text = ""
        End If
    End If
    Exit Sub
    
ErrorHandler:
End Sub
