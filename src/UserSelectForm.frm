VERSION 5.00
Begin VB.Form UserSelectForm 
   Caption         =   "Manage User"
   ClientHeight    =   4035
   ClientLeft      =   9810
   ClientTop       =   4500
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   4815
   Begin VB.TextBox txtPasswordEdit 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Text            =   "<current pass of selected user>"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtUsernameEdit 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "<username of selected user>"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton btnUpdateUser 
      Caption         =   "Update"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton btnDeleteUser 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Change password"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Change username:"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1350
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Manage User"
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
      TabIndex        =   3
      Top             =   240
      Width           =   4455
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
Private u As user
Public PasswordConfirmed As Boolean

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
    Dim result As Collection
    
    isAdmin = IIf(StaffForm.selectedUser.isAdmin, True, False)
    page = IIf(StaffForm.selectedUser.isAdmin, StaffForm.aCurrentPage, StaffForm.rCurrentPage)
    
    ConfirmPasswordForm.Show vbModal
    If PasswordConfirmed Then
        Call DeleteUser(StaffForm.selectedUser.id)
        PasswordConfirmed = False
        
        Set result = GetUser(isAdmin, page)
        Call StaffForm.InitPagination(IIf(isAdmin, "admin", "registrar"), result)
        
        Unload Me
    End If
    
End Sub

Private Sub btnUpdateUser_Click()
    Dim NewUser As New user
    Dim isAdmin As Boolean
    Dim page As Integer
    Dim adminResult As Collection
    Dim registrarResult As Collection
    
    If Len(txtPasswordEdit.Text) < 8 Then
        GoTo PasswordInvalidError
    End If
    isAdmin = IIf(StaffForm.selectedUser.isAdmin, True, False)
    page = IIf(StaffForm.selectedUser.isAdmin, StaffForm.aCurrentPage, StaffForm.rCurrentPage)
    
    With NewUser
        .username = txtUsernameEdit.Text
        .password = txtPasswordEdit.Text
        .isAdmin = IIf(lblPermission = "admin", True, False)
    End With
    
    ConfirmPasswordForm.Show vbModal
    If PasswordConfirmed Then
        Call UpdateUser(StaffForm.selectedUser.id, NewUser)
        
        Set adminResult = GetUser(True, StaffForm.aCurrentPage)
        Set registrarResult = GetUser(False, StaffForm.rCurrentPage)
        
        Call StaffForm.InitPagination("admin", adminResult)
        Call StaffForm.InitPagination("registrar", registrarResult)
        
    End If
PasswordInvalidError:
    MsgBox "Password should be at least 8 characters.", vbExclamation, "Invalid password"
End Sub

Private Sub Form_Load()
    ' StaffForm.SelectedUser contains the information
    ' of the double clicked user on the list.
    Set u = StaffForm.selectedUser
    
    lblSelected.Caption = u.username
    txtUsernameEdit.Text = u.username
    txtPasswordEdit.Text = u.password
End Sub
