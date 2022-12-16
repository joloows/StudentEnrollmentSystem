Attribute VB_Name = "DatabaseModule"
Public db As Database
Public rs As Recordset

Public Sub InitDatabase()
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
        Dim SexField, AddressField As Field
        Dim IsEnrolledField, DateEnrolledField As Field
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
        With rs
            .AddNew
            !username = "jolo"
            !password = "admin"
            !is_admin = True
            !date_created = Date
            .Update
        End With
        rs.Close
        Set rs = Nothing
        Debug.Print "Succesfully created admin user 'jolo'."
    End If
    
End Sub

Public Sub AddEnrollee(En As Enrollee)
    Set rs = db.OpenRecordset("enrollee")
        ' Populate recordset
        With rs
            .AddNew
            !last_name = En.Lname
            !first_name = En.Fname
            !middle_name = En.Mname
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
End Sub

Public Function CreateUser(username As String, password As String, adminPerm As Boolean) As Integer
    ' Query user input to database
    Dim qdf As QueryDef
    Set qdf = db.CreateQueryDef("", "SELECT * FROM staff WHERE username=[_uname]")
    qdf.Parameters("_uname") = username

    Set rs = qdf.OpenRecordset

    If rs.BOF And rs.EOF Then ' If account not exist in database
        rs.Close
        Set rs = Nothing
        
        ' Append data to staff table
        Set rs = db.OpenRecordset("staff")
            With rs
                .AddNew
                !username = username
                !password = password
                !is_admin = adminPerm
                !date_created = Date
                .Update
            End With
            rs.Close
            Set rs = Nothing
            
            CreateUser = 1
            Exit Function
    Else
        CreateUser = 0
    End If
End Function

Public Function Get_Enrollee(Optional page As Integer = 1, Optional search As String = "") As Collection
    Dim qdf As QueryDef
    Dim query As String
    Dim enrollees As New Collection
    Dim En As Enrollee
    Dim total As Integer
    Dim result As New Collection
    
    query = "SELECT * FROM enrollee"
    
    ' Search
    If search <> "" Then
        rs.Close
        Set rs = Nothing
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
        "OR mother_tongue LIKE '*" & search & "*' " & _
        "OR address LIKE '*" & search & "*' " & _
        "OR father_name LIKE '*" & search & "*' " & _
        "OR father_no LIKE '*" & search & "*' " & _
        "OR mother_no LIKE '*" & search & "*' " & _
        "OR guardian_no LIKE '*" & search & "*' " & _
        "OR guardian_name LIKE '*" & search & "*' "
    End If
    
    Set qdf = db.CreateQueryDef("", query)
    Set rs = qdf.OpenRecordset
    
    ' Page
    start_index = (page - 1) * 23 ' where the rs starts
    stop_index = start_index ' where the rs ends
    
    total = 0
    If rs.RecordCount <> 0 Then
        rs.MoveLast
        total = rs.RecordCount
        rs.MoveFirst
    End If
    
    If page > 1 Then
        rs.Move start_index
    End If
    
    i = 1
        While Not rs.EOF And i <= 23
            Set En = New Enrollee
            With En
                .id = rs!enrollee_id
                .Grade = rs!grade_level
                .Lname = rs!last_name
                .Fname = rs!first_name
                .Mname = rs!middle_name
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
            enrollees.Add En
            i = i + 1
            stop_index = stop_index + 1
            rs.MoveNext
        Wend
        
    With result
        .Add enrollees, "enrollees"
        .Add total, "record_count"
        .Add CInt(total / 23) + 1, "pages" ' ceil dividing to get total pages
        .Add start_index + 1, "start_index"
        .Add stop_index, "stop_index"
    End With
    
    Set Get_Enrollee = result
End Function

