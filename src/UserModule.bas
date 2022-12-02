Attribute VB_Name = "UserModule"
Public CurrentUser As New User

Public Function LoginUser(username As String, password As String)
    ' Using a QueryDef with parameters to avoid SQL injection
    Dim qdf As QueryDef
    Set qdf = DatabaseModule.db.CreateQueryDef("", "SELECT * FROM staff WHERE username=[_uname] AND password=[_pw]")
    qdf.Parameters("_uname") = username
    qdf.Parameters("_pw") = password

    Set L = qdf.OpenRecordset

    If L.BOF And L.EOF Then ' If account not exist

        MsgBox "Invalid username or password. Please try again.", vbOKOnly, "Invalid Entry!"
        LoginForm.txtPassword.SetFocus

        ' clean up
        L.Close
        Set L = Nothing
        
    Else ' Login Success
        With L
            .MoveFirst
            LId = !staff_id
            LUsername = !username
            LPassword = !password
            LIsAdmin = !is_admin
        End With
        
        ' clean up
        L.Close
        Set L = Nothing
        
        With CurrentUser
            .id = LId
            .username = LUsername
            .password = LPassword
            .isAdmin = LIsAdmin
        End With
        
        ' Show staff form
        StaffForm.Show
        LoginForm.Hide
        StudentForm.Hide
    End If
End Function


