VERSION 5.00
Begin VB.Form ConfirmPasswordForm 
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   9810
   ClientTop       =   5955
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter password to confirm:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1875
   End
End
Attribute VB_Name = "ConfirmPasswordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = CurrentUser.password Then
        UserSelectForm.PasswordConfirmed = True
        Unload Me
    Else
        MsgBox "Password does not match the password of the current user.", vbCritical, "Password not match"
        Text1.Text = ""
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
