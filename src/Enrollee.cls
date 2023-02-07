VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enrollee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mId As Long
Private mEnrolled As Boolean
Private mGrade As Integer
Private mSection As String
Private mLname As String
Private mFname As String
Private mMname As String
Private mSex As String
Private mTuition As Integer
Private mAge As Integer
Private mBirthdate As Date
Private mBirthplace As String
Private mMt As String
Private mAddress As String
Private mFatherName As String
Private mFNum As String
Private mMotherName As String
Private mMnum As String
Private mGuardianName As String
Private mGnum As String
Private mSubmission As Date

' Get
' Usage po nento is Enrollee.id, or Enrollee.Lname, etc.
' Return value is kung ano po yung nakasave na
' value ng mId.
Public Property Get id() As Long
    id = mId
End Property

Public Property Get Enrolled() As Boolean
    Enrolled = mEnrolled
End Property

Public Property Get Lname() As String
    Lname = mLname
End Property

Public Property Get Fname() As String
    Fname = mFname
End Property

Public Property Get Mname() As String
    Mname = mMname
End Property

Public Property Get Grade() As Integer
    Grade = mGrade
End Property

Public Property Get Section() As String
    Section = mSection
End Property

Public Property Get Sex() As String
    Sex = mSex
End Property

Public Property Get Tuition() As Integer
    Tuition = mTuition
End Property

Public Property Get Age() As Integer
    Age = mAge
End Property

Public Property Get Birthdate() As Date
    Birthdate = mBirthdate
End Property

Public Property Get Birthplace() As String
    Birthplace = mBirthplace
End Property

Public Property Get Mt() As String
    Mt = mMt
End Property

Public Property Get Address() As String
    Address = mAddress
End Property

Public Property Get Fathername() As String
    Fathername = mFatherName
End Property

Public Property Get Fnum() As String
    Fnum = mFNum
End Property

Public Property Get MotherName() As String
    MotherName = mMotherName
End Property

Public Property Get Mnum() As String
    Mnum = mMnum
End Property

Public Property Get GuardianName() As String
    GuardianName = mGuardianName
End Property

Public Property Get Gnum() As String
    Gnum = mGnum
End Property

Public Property Get Submission() As Date
    Submission = mSubmission
End Property

' Let
' Dito po natin inaasign value ni mId
' To assign a value to the property of Enrollee,
' we do Enrollee.<kung anong property ng Enrollee> = <value na gusto nating i-assign.>
Public Property Let id(Value As Long)
    mId = Value
End Property

Public Property Let Enrolled(Value As Boolean)
    mEnrolled = Value
End Property

Public Property Let Lname(Value As String)
    mLname = Value
End Property

Public Property Let Fname(Value As String)
    mFname = Value
End Property

Public Property Let Mname(Value As String)
    mMname = Value
End Property

Public Property Let Grade(Value As Integer)
    mGrade = Value
End Property

Public Property Let Section(Value As String)
    mSection = Value
End Property

Public Property Let Sex(Value As String)
    mSex = Value
End Property

Public Property Let Tuition(Value As Integer)
    mTuition = Value
End Property

Public Property Let Age(Value As Integer)
    mAge = Value
End Property

Public Property Let Birthdate(Value As Date)
    mBirthdate = Value
End Property

Public Property Let Birthplace(Value As String)
    mBirthplace = Value
End Property

Public Property Let Mt(Value As String)
    mMt = Value
End Property

Public Property Let Address(Value As String)
    mAddress = Value
End Property

Public Property Let Fathername(Value As String)
    mFatherName = Value
End Property

Public Property Let Fnum(Value As String)
    mFNum = Value
End Property

Public Property Let MotherName(Value As String)
    mMotherName = Value
End Property

Public Property Let Mnum(Value As String)
    mMnum = Value
End Property

Public Property Let GuardianName(Value As String)
    mGuardianName = Value
End Property

Public Property Let Gnum(Value As String)
    mGnum = Value
End Property

Public Property Let Submission(Value As Date)
    mSubmission = Value
End Property

