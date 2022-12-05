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
        
        ' Create temporary admin account for debugging
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
        rs.Close
        Set rs = Nothing
End Sub
