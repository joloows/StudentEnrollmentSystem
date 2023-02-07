VERSION 5.00
Begin VB.Form UserSelectForm 
   Caption         =   "Edit User"
   ClientHeight    =   4770
   ClientLeft      =   9810
   ClientTop       =   4500
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   4815
   Begin VB.TextBox txtPasswordEdit 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "<current pass of selected user>"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtUsernameEdit 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Text            =   "<username of selected user>"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton btnChange 
      Caption         =   "Change"
      Height          =   300
      Left            =   2520
      TabIndex        =   6
      Top             =   1520
      Width           =   1095
   End
   Begin VB.CommandButton btnUpdateUser 
      Caption         =   "Update"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton btnDeleteUser 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Change password"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Change username:"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   1350
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Edit User"
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblPermission 
      AutoSize        =   -1  'True
      Caption         =   "<permission>"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Permission: "
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label lblSelected 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "<username of selected user>"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "UserSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private u As User
Private result As Collection

Private Sub btnChange_Click()
    If lblPermission.Caption = "admin" Then
        lblPermission.Caption = "registrar"
    Else
        lblPermission.Caption = "admin"
    End If
End Sub

Private Sub btnDeleteUser_Click()
    Dim isAdmin As Boolean
    Dim page As Integer
    
    isAdmin = IIf(StaffForm.selectedUser.isAdmin, True, False)
    page = IIf(StaffForm.selectedUser.isAdmin, StaffForm.aCurrentPage, StaffForm.rCurrentPage)
    
    x = InputBox("Enter your password to confirm account deletion")
    If x = CurrentUser.password Then
        Call DeleteUser(StaffForm.selectedUser.id)
        MsgBox "User successfully deleted.", vbInformation, "Success"
        
        Set result = GetUser(isAdmin, page)
        Call StaffForm.InitPagination(IIf(isAdmin, "admin", "registrar"), result)
        
        Unload Me
    Else
        MsgBox "The password entered doesn't match the current user password.", vbExclamation, "Error"
    End If
End Sub

Private Sub btnUpdateUser_Click()
    Dim NewUser As New User
    Dim isAdmin As Boolean
    Dim page As Integer
    
    isAdmin = IIf(StaffForm.selectedUser.isAdmin, True, False)
    page = IIf(StaffForm.selectedUser.isAdmin, StaffForm.aCurrentPage, StaffForm.rCurrentPage)
    
    With NewUser
        .username = txtUsernameEdit.Text
        .password = txtPasswordEdit.Text
        .isAdmin = IIf(lblPermission = "admin", True, False)
    End With
    
    x = InputBox("Enter your password to confirm account deletion")
    If x = CurrentUser.password Then
        Call UpdateUser(StaffForm.selectedUser.id, NewUser)
        MsgBox "User successfully updated.", vbInformation, "Success"
        
        Set result = GetUser(isAdmin, page)
        Call StaffForm.InitPagination(IIf(isAdmin, "admin", "registrar"), result)
        
        Unload Me
    Else
        MsgBox "The password entered doesn't match the current user password.", vbExclamation, "Error"
    End If
    
    
End Sub

Private Sub Form_Load()
    ' StaffForm.SelectedUser contains the information
    ' of the double clicked user on the list.
    Set u = StaffForm.selectedUser
    
    lblSelected.Caption = u.username
    txtUsernameEdit.Text = u.username
    txtPasswordEdit.Text = u.password
    If u.isAdmin Then
        lblPermission.Caption = "admin"
    Else
        lblPermission.Caption = "registrar"
    End If
End Sub
