Attribute VB_Name = "DatabaseModule"
Dim db As Database
Dim rs As Recordset

Public Sub InitDatabase()
    Dim dbPath As String
    
    Dim EnrolleeTable, StaffTable As TableDef
    
    Dim EnrolleeIdField, GradeLevelField, SectionField As Field
    Dim LastNameField, FirstNameField, MiddleNameField, ExNameField As Field
    Dim SexField, AddressField As Field
    Dim IsEnrolledField, DateEnrolledField As Field
    Dim FatherNameField, MotherNameField, GuardianNameField As Field
    
    Dim StaffIdField, UsernameField, PasswordField, IsAdminField, DateCreatedField As Field
    
    dbPath = App.Path & "\database.accdb"
    
    ' for debugging -- resets database if it already exists
    If Dir(dbPath) <> "" Then
        Kill dbPath
    End If
    
    ' Creates the database
    Set db = CreateDatabase(dbPath, dbLangGeneral, dbEncrypted)
    
    ' Creates table "enrollee"
    Set EnrolleeTable = db.CreateTableDef("enrollee")
    
    ' Creates fields for table enrollee
    Set EnrolleeIdField = EnrolleeTable.CreateField("enrollee_id", dbLong)
    EnrolleeIdField.Attributes = dbAutoIncrField
    Set GradeLevelField = EnrolleeTable.CreateField("grade_level", dbInteger)
    Set SectionField = EnrolleeTable.CreateField("section", dbText)
    Set LastNameField = EnrolleeTable.CreateField("last_name", dbText)
    Set FirstNameField = EnrolleeTable.CreateField("first_name", dbText)
    Set MiddleNameField = EnrolleeTable.CreateField("middle_name", dbText)
    Set ExNameField = EnrolleeTable.CreateField("extension_name", dbText)
    Set SexField = EnrolleeTable.CreateField("sex", dbText)
    Set AddressField = EnrolleeTable.CreateField("address", dbText)
    Set IsEnrolledField = EnrolleeTable.CreateField("is_enrolled", dbBoolean)
    Set DateEnrolledField = EnrolleeTable.CreateField("date_enrolled", dbDate)
    Set FatherNameField = EnrolleeTable.CreateField("father_name", dbText)
    Set MotherNameField = EnrolleeTable.CreateField("mother_name", dbText)
    Set GuardianNameField = EnrolleeTable.CreateField("guardian_name", dbText)
    
    ' Append fields to table enrollee
    EnrolleeTable.Fields.Append EnrolleeIdField
    EnrolleeTable.Fields.Append GradeLevelField
    EnrolleeTable.Fields.Append SectionField
    EnrolleeTable.Fields.Append LastNameField
    EnrolleeTable.Fields.Append FirstNameField
    EnrolleeTable.Fields.Append MiddleNameField
    EnrolleeTable.Fields.Append ExNameField
    EnrolleeTable.Fields.Append SexField
    EnrolleeTable.Fields.Append AddressField
    EnrolleeTable.Fields.Append IsEnrolledField
    EnrolleeTable.Fields.Append DateEnrolledField
    EnrolleeTable.Fields.Append FatherNameField
    EnrolleeTable.Fields.Append MotherNameField
    EnrolleeTable.Fields.Append GuardianNameField
    
    ' Append table enrollee to db
    db.TableDefs.Append EnrolleeTable
    
    ' Create table "staff"
    Set StaffTable = db.CreateTableDef("staff")
    
    ' Create fields for table staff
    Set StaffIdField = StaffTable.CreateField("staff_id", dbLong)
    StaffIdField.Attributes = dbAutoIncrField
    Set UsernameField = StaffTable.CreateField("username", dbText)
    Set PasswordField = StaffTable.CreateField("password", dbText)
    Set IsAdminField = StaffTable.CreateField("is_admin", dbBoolean)
    Set DateCreatedField = StaffTable.CreateField("date_created", dbDate)
    
    ' Append fields to table staff
    StaffTable.Fields.Append StaffIdField
    StaffTable.Fields.Append UsernameField
    StaffTable.Fields.Append PasswordField
    StaffTable.Fields.Append IsAdminField
    StaffTable.Fields.Append DateCreatedField
    
    ' Append table staff to db
    db.TableDefs.Append StaffTable
    
    MsgBox "Succesfully created new database.", vbInformation
End Sub
