Attribute VB_Name = "DatabaseModule"
Public db As Database
Public rs As Recordset

Public Sub InitDatabase()
Attribute InitDatabase.VB_Description = "Initializes the database needed for the application. Creates the database if the database does not exist. If the database exists, it just opens the database"
    Dim dbPath As String
    dbPath = App.Path & "\database.accdb"
    
    If Dir(dbPath) <> "" Then ' If database exist
        
        ' Connect to database
        Set db = OpenDatabase(dbPath)
        Debug.Print "Successfully opened " & db.Name
        
    Else ' If database does not exist
    
        Dim EnrolleeTable, StaffTable As TableDef
        Dim EnrolleeIdField, GradeLevelField, SectionField As Field
        Dim LastNameField, FirstNameField, MiddleNameField, ExNameField As Field
        Dim SexField, AddressField, TotalFeeField, WithUniformField As Field
        Dim PaymentTypeField, PaymentField, IsEnrolledField, DateEnrolledField As Field
        Dim FatherNameField, MotherNameField, GuardianNameField As Field
        Dim FatherNumField, MotherNumField, GuardianNumField As Field
        Dim StaffIdField, UsernameField, PasswordField, IsAdminField, DateCreatedField As Field

        ' Create and Connect to the database
        Set db = CreateDatabase(dbPath, dbLangGeneral, dbEncrypted)
        
        ' Creates table "enrollee"
        Set EnrolleeTable = db.CreateTableDef("enrollee")
        
        ' Creates fields for table enrollee
        With EnrolleeTable
            Set EnrolleeIdField = .CreateField("enrollee_id", dbLong)
            EnrolleeIdField.Attributes = dbAutoIncrField
            Set GradeLevelField = .CreateField("grade_level", dbInteger)
            Set IsEnrolledField = .CreateField("is_enrolled", dbBoolean)
            Set SectionField = .CreateField("section", dbText)
            Set LastNameField = .CreateField("last_name", dbText)
            Set FirstNameField = .CreateField("first_name", dbText)
            Set MiddleNameField = .CreateField("middle_name", dbText)
            Set SexField = .CreateField("sex", dbText)
            Set AgeField = .CreateField("age", dbInteger)
            Set TotalFeeField = .CreateField("total_fee", dbCurrency)
            Set WithUniformField = .CreateField("with_uniform", dbBoolean)
            Set PaymentTypeField = .CreateField("payment_type", dbText)
            Set PaymentField = .CreateField("payment", dbCurrency)
            Set BirthdateField = .CreateField("birthdate", dbDate)
            Set BirthplaceField = .CreateField("birthplace", dbText)
            Set MtField = .CreateField("mother_tongue", dbText)
            Set AddressField = .CreateField("address", dbText)
            Set DateEnrolledField = .CreateField("date_enrolled", dbDate)
            Set FatherNameField = .CreateField("father_name", dbText)
            Set FatherNumField = .CreateField("father_no", dbText)
            Set MotherNameField = .CreateField("mother_name", dbText)
            Set MotherNumField = .CreateField("mother_no", dbText)
            Set GuardianNameField = .CreateField("guardian_name", dbText)
            Set GuardianNumField = .CreateField("guardian_no", dbText)
        End With
        
        ' Append fields to table enrollee
        With EnrolleeTable.Fields
            .Append EnrolleeIdField
            .Append GradeLevelField
            .Append IsEnrolledField
            .Append SectionField
            .Append LastNameField
            .Append FirstNameField
            .Append MiddleNameField
            .Append TotalFeeField
            .Append WithUniformField
            .Append PaymentTypeField
            .Append PaymentField
            .Append SexField
            .Append AgeField
            .Append BirthdateField
            .Append BirthplaceField
            .Append MtField
            .Append AddressField
            .Append DateEnrolledField
            .Append FatherNameField
            .Append FatherNumField
            .Append MotherNameField
            .Append MotherNumField
            .Append GuardianNameField
            .Append GuardianNumField
        End With
        
        ' Append table enrollee to db
        db.TableDefs.Append EnrolleeTable
        
        ' Create table "staff"
        Set StaffTable = db.CreateTableDef("staff")
        
        ' Create fields for table staff
        With StaffTable
            Set StaffIdField = .CreateField("staff_id", dbLong)
            StaffIdField.Attributes = dbAutoIncrField
            Set UsernameField = .CreateField("username", dbText)
            Set PasswordField = .CreateField("password", dbText)
            Set IsAdminField = .CreateField("is_admin", dbBoolean)
            Set DateCreatedField = .CreateField("date_created", dbDate)
        End With
        
        ' Append fields to table staff
        With StaffTable.Fields
            .Append StaffIdField
            .Append UsernameField
            .Append PasswordField
            .Append IsAdminField
            .Append DateCreatedField
        End With
        
        ' Append table staff to db
        db.TableDefs.Append StaffTable
        
        Debug.Print "Succesfully created new database."
        
        ' Create first admin account
        Set rs = db.OpenRecordset("staff")
        username = "admin"
        password = "adminpass"
        With rs
            .AddNew
            !username = username
            !password = password
            !is_admin = True
            !date_created = Date
            .Update
        End With
        rs.Close
        Set rs = Nothing
        
        Debug.Print "Succesfully created admin user '" & username & "'."
        Debug.Print "user: " & username & vbNewLine & "password: " & password
        
    End If
    
End Sub

Public Sub AddEnrollee(En As Enrollee)
    On Error GoTo ErrorHandler
    Set rs = db.OpenRecordset("enrollee")
    ' Populate recordset with the Enrollee object "En"
    ' Checkout Enrollee.cls
    With rs
        .AddNew
        !last_name = En.Lname
        !first_name = En.Fname
        !middle_name = En.Mname
        !total_fee = En.TotalFee
        !with_uniform = En.WithUniform
        !payment_type = En.PaymentType
        !payment = En.payment
        !grade_level = En.Grade
        !Sex = En.Sex
        !Age = En.Age
        !Birthdate = En.Birthdate
        !Birthplace = En.Birthplace
        !mother_tongue = En.Mt
        !Address = En.Address
        !father_name = En.Fathername
        !father_no = En.Fnum
        !mother_name = En.MotherName
        !mother_no = En.Mnum
        !guardian_name = En.GuardianName
        !guardian_no = En.Gnum
        !date_enrolled = En.Submission
        .Update
    End With
    ' Clean up
    rs.Close
    Set rs = Nothing
    MsgBox "Submission recorded successfuly.", vbInformation, "Success"
    Exit Sub
ErrorHandler:
    MsgBox "Please fill out all of the info.", vbExclamation, "Error"
End Sub

Public Sub CreateUser(username As String, password As String, adminPerm As Boolean)
    On Error GoTo ErrorHandler
    ' Query user input to database
    Dim qdf As QueryDef
    Set qdf = db.CreateQueryDef("", "SELECT * FROM staff WHERE username=[_uname]")
    qdf.Parameters("_uname") = username

    Set rs = qdf.OpenRecordset

    If rs.BOF And rs.EOF Then ' If account not exist in database
        rs.Close
        Set rs = Nothing
        Set rs = db.OpenRecordset("staff")
            With rs
                .AddNew
                !username = username
                !password = password
                !is_admin = adminPerm
                !date_created = Date
                .Update
            End With
            MsgBox "Succesfully created user '" & username & "'.", vbInformation, "Success"
            Unload CreateAccForm
    Else
        MsgBox "username already exists.", vbCritical, "Error"
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Description
    MsgBox "Please fill out all of the info.", vbExclamation, "Error"
End Sub

Public Function GetEnrollee(Optional page As Integer = 1, Optional search As String = "") As Collection
    Dim qdf As QueryDef
    Dim query As String
    Dim enrollees As New Collection
    Dim En As Enrollee
    Dim total As Integer
    Dim result As New Collection
    
    Debug.Print "Get_Enrollee()"
    
    query = "SELECT * FROM enrollee"
    
    ' If a user uses the search bar
    If search <> "" Then
        ' this is SQL.
        query = "SELECT * FROM enrollee " & _
        "WHERE enrollee_id LIKE '*" & search & "*' " & _
        "OR grade_level LIKE '*" & search & "*' " & _
        "OR last_name LIKE '*" & search & "*' " & _
        "OR first_name LIKE '*" & search & "*' " & _
        "OR middle_name LIKE '*" & search & "*'" & _
        "OR sex LIKE '*" & search & "*' " & _
        "OR age LIKE '*" & search & "*' " & _
        "OR birthdate LIKE '*" & search & "*' " & _
        "OR birthplace LIKE '*" & search & "*' " & _
        "OR date_enrolled LIKE '*" & search & "*' " & _
        "OR address LIKE '*" & search & "*' " & _
        "OR father_name LIKE '*" & search & "*' " & _
        "OR mother_name LIKE '*" & search & "*' " & _
        "OR guardian_name LIKE '*" & search & "*' "
    End If
    
    Set qdf = db.CreateQueryDef("", query)
    Set rs = qdf.OpenRecordset
    
    ' Page
    startIndex = (page - 1) * 23    ' where the rs starts
    stopIndex = startIndex          ' where the rs ends
    
    ' Gets the total record count on the recordset
    total = 0
    If rs.RecordCount <> 0 Then
        rs.MoveLast
        total = rs.RecordCount
        rs.MoveFirst
    End If
    
    ' Move at the start of the rs
    If page > 1 Then
        rs.Move startIndex
    End If
    
    ' Start processing the recordset.
    ' Nag-loloop until naka process na sya ng 23 records ng enrollees.
    ' For every page, 23 records ang dinidisplay ng list.
    ' 1. Creates an object Enrollee (check out Enrollee.cls) for every loop
    ' 2. Saves the recordset to the properties of Enrollee object
    ' 3. Adds the Enrollee object to the enrollees Collection. (Try to search for "Collection vb6" on google)
    i = 1
    While Not rs.EOF And i <= 23
        Debug.Print CStr(rs!date_enrolled)
        Set En = New Enrollee ' 1.
        With En ' 2.
            ' Enrollee.property = rs!field_name
            .id = rs!enrollee_id
            .Enrolled = rs!is_enrolled
            .Grade = rs!grade_level
            .Section = IIf(IsNull(rs!Section), "N/A", rs!Section)
            .Lname = rs!last_name
            .Fname = rs!first_name
            .Mname = rs!middle_name
            .TotalFee = rs!total_fee
            .WithUniform = rs!with_uniform
            .PaymentType = rs!payment_type
            .payment = rs!payment
            .Sex = rs!Sex
            .Age = rs!Age
            .Birthdate = rs!Birthdate
            .Birthplace = rs!Birthplace
            .Mt = rs!mother_tongue
            .Address = rs!Address
            .Fathername = rs!father_name
            .Fnum = rs!father_no
            .MotherName = rs!mother_name
            .Mnum = rs!mother_no
            .GuardianName = rs!guardian_name
            .Gnum = rs!guardian_no
            .Submission = rs!date_enrolled
        End With
        enrollees.Add En ' 3.
        i = i + 1
        stopIndex = stopIndex + 1
        rs.MoveNext
    Wend
    
    pages = (total \ 23)
    If pages Mod 2 = 1 Or pages = 0 Then
        pages = pages + 1
    End If
    
    ' Add all the needed by the pagination the the result Collection
    With result
        .Add enrollees, "enrollees"
        .Add total, "recordCount"           ' total records sa table or search query
        .Add pages, "pages"                 ' kung ilang pages
        .Add startIndex + 1, "startIndex"   ' saang record nag uumpisa
        .Add stopIndex, "stopIndex"         ' saang record nagtatapos
    End With
    
    ' Return the result to where the function called
    Set GetEnrollee = result
End Function

' Stuff here is essentially the same with the GetEnrollee Function.
' The only difference is instead of handling Enrollees, this function
' handles Users.
Public Function GetUser(isAdmin As Boolean, Optional page As Integer = 1, Optional search As String = "") As Collection
    Dim qdf As QueryDef
    Dim users As New Collection
    Dim u As user
    Dim result As New Collection
    
    Set qdf = db.CreateQueryDef("", "SELECT * FROM staff WHERE NOT username=[_currentuser] AND is_admin=[_isadmin]")
    qdf.Parameters("_currentuser") = CurrentUser.username
    qdf.Parameters("_isadmin") = isAdmin
    
    Set rs = qdf.OpenRecordset
    
    If search <> "" Then
        rs.Filter = "staff_id LIKE '*" & search & "*' " & _
        "OR username LIKE '*" & search & "*' " & _
        "OR date_created LIKE '*" & search & "*' "
    End If
    
    ' Page
    startIndex = (page - 1) * 23 ' where the rs starts
    stopIndex = startIndex ' where the rs ends
    
    total = 0
    If rs.RecordCount <> 0 Then
        rs.MoveLast
        total = rs.RecordCount
        rs.MoveFirst
    End If
    
    If page > 1 Then
        rs.Move startIndex
    End If
    
    i = 1
    While Not rs.EOF And i <= 23
        Set u = New user
        With u
           u.id = rs!staff_id
           u.username = rs!username
           u.password = rs!password
           u.isAdmin = rs!is_admin
           u.dateCreated = rs!date_created
        End With
        users.Add u
        i = i + 1
        stopIndex = stopIndex + 1
        rs.MoveNext
    Wend
    
    temp = (total / 23)
    pages = CInt(temp)
    If temp - pages > 0 Or pages = 0 Then
        pages = pages + 1
    End If
    
    With result
        .Add users, "users"
        .Add total, "recordCount"
        .Add pages, "pages"
        .Add startIndex + 1, "startIndex"
        .Add stopIndex, "stopIndex"
    End With
    
    rs.Close
    Set rs = Nothing
    
    Set GetUser = result
End Function

Public Function UpdateUser(id As Integer, NewUser As user)
    On Error GoTo ErrorHandler
    Dim qdf As QueryDef
    
    Set qdf = db.CreateQueryDef("", "SELECT * FROM staff WHERE username=[_uname]")
    qdf.Parameters("_uname") = NewUser.username

    Set rs = qdf.OpenRecordset

    If rs.BOF And rs.EOF Then ' If account not exist in database
        rs.Close
        Set rs = Nothing
        Set qdf = db.CreateQueryDef("", "SELECT * FROM staff WHERE staff_id=[_id]")
        qdf.Parameters("_id") = id
        Set rs = qdf.OpenRecordset
        With rs
            .Edit
            !username = NewUser.username
            !password = NewUser.password
            !is_admin = NewUser.isAdmin
            .Update
        End With
        rs.Close
        Set rs = Nothing
        MsgBox "User successfully updated.", vbInformation, "Success"
        Unload UserSelectForm
    Else
        MsgBox "username already exists.", vbCritical, "Error"
    End If
    
    UserSelectForm.PasswordConfirmed = False
    Exit Function
ErrorHandler:
    Debug.Print Err.Description
    MsgBox "Please fill out all of the info.", vbExclamation, "Error"
    PasswordConfirmed = False
End Function

Public Function UpdateCurrentUser(id As Integer, NewUser As user)
    Dim qdf As QueryDef
    
    Set qdf = db.CreateQueryDef("", "SELECT * FROM staff WHERE staff_id=[_id]")
    qdf.Parameters("_id") = id
    Set rs = qdf.OpenRecordset
    With rs
        .Edit
        !username = NewUser.username
        !password = NewUser.password
        !is_admin = NewUser.isAdmin
        .Update
    End With
    rs.Close
    Set rs = Nothing
    MsgBox "User successfully updated.", vbInformation, "Success"
End Function

Public Function DeleteUser(id As Integer)
    Dim qdf As QueryDef
    Set qdf = db.CreateQueryDef("", "SELECT * FROM staff WHERE staff_id=[_id]")
    qdf.Parameters("_id") = id
    
    Set rs = qdf.OpenRecordset
    
    rs.Delete
    
    rs.Close
    Set rs = Nothing
    MsgBox "User successfully deleted.", vbInformation, "Success"
End Function

Public Function UpdateEnrollee(id As Integer, NewEnrollee As Enrollee)
    On Error GoTo ErrorHandler
    Dim qdf As QueryDef
    Set qdf = db.CreateQueryDef("", "SELECT * FROM enrollee WHERE enrollee_id=[_id]")
    qdf.Parameters("_id") = id
    
    Set rs = qdf.OpenRecordset

    With rs
        .Edit
        !last_name = NewEnrollee.Lname
        !first_name = NewEnrollee.Fname
        !middle_name = NewEnrollee.Mname
        !grade_level = NewEnrollee.Grade
        !Sex = NewEnrollee.Sex
        !Age = NewEnrollee.Age
        !Birthdate = NewEnrollee.Birthdate
        !Birthplace = NewEnrollee.Birthplace
        !mother_tongue = NewEnrollee.Mt
        !Address = NewEnrollee.Address
        !father_name = NewEnrollee.Fathername
        !father_no = NewEnrollee.Fnum
        !mother_name = NewEnrollee.MotherName
        !mother_no = NewEnrollee.Mnum
        !guardian_name = NewEnrollee.GuardianName
        !guardian_no = NewEnrollee.Gnum
        .Update
    End With
    
    rs.Close
    Set rs = Nothing
    MsgBox "Submission recorded successfuly.", vbInformation, "Success"
    Exit Function
ErrorHandler:
    MsgBox "Please fill out all of the info.", vbExclamation, "Error"
End Function

Public Function DeleteEnrollee(id As Integer)
    Dim qdf As QueryDef
    Set qdf = db.CreateQueryDef("", "SELECT * FROM enrollee WHERE enrollee_id=[_id]")
    qdf.Parameters("_id") = id
    
    Set rs = qdf.OpenRecordset
    
    rs.Delete
    
    rs.Close
    Set rs = Nothing
End Function

Public Function UpdateEnrolleeStatus(status As Boolean, id As Integer)
    Dim En As Enrollee
    Dim qdf As QueryDef
    
        Set qdf = db.CreateQueryDef("", "SELECT * FROM enrollee WHERE enrollee_id=[_id]")
        qdf.Parameters("_id") = id
        
        Set rs = qdf.OpenRecordset
    
        With rs
            .Edit
            !is_enrolled = status
            .Update
        End With
        
        rs.Close
        Set rs = Nothing
End Function

Public Function AssignEnrolleeSection(Section As String, id As Integer)
    Dim En As Enrollee
    Dim qdf As QueryDef
    
    Set qdf = db.CreateQueryDef("", "SELECT * FROM enrollee WHERE enrollee_id=[_id]")
    qdf.Parameters("_id") = id
    
    Set rs = qdf.OpenRecordset

    With rs
        .Edit
        !Section = Section
        .Update
    End With
    
    rs.Close
    Set rs = Nothing
    
    AssignEnrolleeSection = 0
End Function

Public Function AddEnrolleeFeePayment(payment As Currency, id As Integer) As Integer
    On Err GoTo ErrorHandler
    Dim En As Enrollee
    Dim qdf As QueryDef
    Dim curFee As Currency
    Dim curPayment As Currency
    
    Set qdf = db.CreateQueryDef("", "SELECT * FROM enrollee WHERE enrollee_id=[_id]")
    qdf.Parameters("_id") = id
    
    Set rs = qdf.OpenRecordset
    
    curPayment = rs!payment
    curFee = rs!total_fee - curPayment
    
    With rs
        .Edit
        !payment = !payment + payment
        .Update
    End With
    
    rs.Close
    Set rs = Nothing
    
    AddEnrolleeFeePayment = 0
    Exit Function
ErrorHandler:
    AddEnrolleFeePayment = -1
End Function
