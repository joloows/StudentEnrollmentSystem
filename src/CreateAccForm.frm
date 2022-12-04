VERSION 5.00
Begin VB.Form CreateAccForm 
   Caption         =   "AES Enrollment System"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Has admin permissions"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Confirm Password:"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Create Account"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2310
   End
End
Attribute VB_Name = "CreateAccForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
