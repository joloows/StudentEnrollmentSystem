VERSION 5.00
Begin VB.Form DebugForm 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   3360
   ClientTop       =   2655
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Create Database"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "DebugForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call InitDatabase
End Sub
