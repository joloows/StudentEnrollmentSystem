VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIsAuthenticated
Private mId As Long
Private mUsername As String
Private mPassword As String
Private mIsAdmin As Boolean
Private mDateCreated As Date


' Same thing with Enrollee Class.
Public Property Get isAuthenticated() As Boolean
    isAuthenticated = mIsAuthenticated
End Property

Public Property Let isAuthenticated(Value As Boolean)
    mIsAuthenticated = Value
End Property

Public Property Get id() As Long
    id = mId
End Property

Public Property Let id(Value As Long)
    mId = Value
End Property

Public Property Get username() As String
    username = mUsername
End Property

Public Property Let username(Value As String)
    mUsername = Value
End Property

Public Property Get password() As String
    password = mPassword
End Property

Public Property Let password(Value As String)
    mPassword = Value
End Property

Public Property Get isAdmin() As Boolean
    isAdmin = mIsAdmin
End Property

Public Property Let isAdmin(Value As Boolean)
    mIsAdmin = Value
End Property
    
Public Property Get dateCreated() As Date
    dateCreated = mDateCreated
End Property

Public Property Let dateCreated(Value As Date)
    mDateCreated = Value
End Property

