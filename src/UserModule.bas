Attribute VB_Name = "UserModule"
Public CurrentUser As New User

Public Sub LoginUser(username As String, password As String)
    ' Using a QueryDef with parameters to avoid SQL injection
    Dim qdf As QueryDef
    Set qdf = DatabaseModule.db.CreateQueryDef("", "SELECT * FROM staff WHERE username=[_uname] AND password=[_pw]")
    qdf.Parameters("_uname") = username
    qdf.Parameters("_pw") = password

    Set L = qdf.OpenRecordset

    If L.BOF And L.EOF Then ' If account not exist in database

        MsgBox "Invalid username or password. Please try again.", vbOKOnly, "Invalid Entry!"
        LoginForm.txtPassword.SetFocus

        ' clean up
        L.Close
        Set L = Nothing
        
    Else ' If account exist
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
            .isAuthenticated = True
            .id = LId
            .username = LUsername
            .password = LPassword
            .isAdmin = LIsAdmin
        End With
        
        ' Clear login info from previous
        LoginForm.txtUsername.Text = ""
        LoginForm.txtPassword.Text = ""
            
        ' Show staff form
        StaffForm.Show
        Unload LoginForm
        Unload StudentForm
        Set LoginForm = Nothing
        Set StudentForm = Nothing
    End If
End Sub

Public Sub LogoutUser()
        StudentForm.Show
        Unload StaffForm
        Set StaffForm = Nothing
        
        ' Depopulate CurrentUser properties
        With CurrentUser
            .isAuthenticated = False
            .id = 0
            .username = ""
            .password = ""
            .isAdmin = False
        End With
End Sub

