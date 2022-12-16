VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4395
   ClientLeft      =   7650
   ClientTop       =   4440
   ClientWidth     =   7980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7980
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      Picture         =   "LoginForm.frx":0000
      ScaleHeight     =   1575
      ScaleWidth      =   1695
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton loginBtn 
      Caption         =   "Login"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   5895
   End
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2160
      Width           =   5895
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
      Left            =   2400
      TabIndex        =   8
      Top             =   360
      Width           =   4860
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Caption         =   "depedcavite.aguinadoes108016@gmail.com |  (046) 484 7623"
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   840
      Width           =   4890
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      Caption         =   "Kap. M. S. Victa, Kawit, Philippines, 4104"
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   4875
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub loginBtn_Click()
    Dim username As String
    Dim password As String
    username = Me.txtUsername
    password = Me.txtPassword
    
    'validation to check if the user entered the username in the username field
    If IsNull(username) Or username = "" Then
        MsgBox "You must enter a username.", vbExclamation + vbOKOnly, "Required Data"
        txtUsername.SetFocus
        Exit Sub
    End If
    
    'validation to check if the user entered the password in the password field
    If IsNull(password) Or password = "" Then
        MsgBox "You must enter a password.", vbExclamation + vbOKOnly, "Required Data"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    x = LoginUser(username, password)
    If x = 1 Then ' If successful login
        ' Show staff form
        StaffForm.Show
        Unload LoginForm
    End If
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
End Sub
