VERSION 5.00
Begin VB.Form GotoForm 
   Caption         =   "Go to"
   ClientHeight    =   1770
   ClientLeft      =   10110
   ClientTop       =   6165
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   ScaleHeight     =   1770
   ScaleWidth      =   3735
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   120
      X2              =   3600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label GoToForm 
      AutoSize        =   -1  'True
      Caption         =   "Page:"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   420
   End
End
Attribute VB_Name = "GotoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnGo_Click()
    If CInt(txtPage.Text) > StaffForm.eTotalPage Or CInt(txtPage.Text) < 1 Then
        MsgBox "Page input either exceeds total pages or input is less than 1.", vbCritical, "Invalid Page"
    Else
        StaffForm.eCurrentPage = CInt(txtPage.Text)
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
