VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4395
   ClientLeft      =   7650
   ClientTop       =   4440
   ClientWidth     =   7980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7980
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2640
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Label Label4 
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
      Height          =   735
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()

End Sub
