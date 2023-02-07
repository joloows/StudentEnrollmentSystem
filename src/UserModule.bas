Attribute VB_Name = "UserModule"
Public CurrentUser As New User

Public Sub LoginUser(username As String, password As String)
    Dim rs As Recordset
    Dim qdf As QueryDef
    
    ' Query user input to database
    Set qdf = db.CreateQueryDef("", "SELECT * FROM staff WHERE username=[_uname] AND password=[_pw]")
    qdf.Parameters("_uname") = username
    qdf.Parameters("_pw") = password

    Set rs = qdf.OpenRecordset

    If rs.BOF And rs.EOF Then ' If account not exist in database

        MsgBox "Invalid username or password. Please try again.", vbOKOnly, "Invalid Entry!"
        LoginForm.txtPassword.SetFocus

        ' clean up
        rs.Close
        Set rs = Nothing
        
    Else ' If account exist
        With rs
            .MoveFirst
            LId = !staff_id
            LUsername = !username
            LPassword = !password
            LIsAdmin = !is_admin
        End With
        
        ' clean up
        rs.Close
        Set rs = Nothing
        
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
        Unload LoginForm
        Unload StudentForm
        StaffForm.Show

    End If
End Sub

Public Sub LogoutUser()
        StudentForm.Show
        Unload StaffForm
        
        ' Depopulate CurrentUser properties
        With CurrentUser
            .isAuthenticated = False
            .id = 0
            .username = ""
            .password = ""
            .isAdmin = False
        End With
End Sub

