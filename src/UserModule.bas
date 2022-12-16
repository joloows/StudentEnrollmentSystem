Attribute VB_Name = "UserModule"
Public CurrentUser As New User

Public Function LoginUser(username As String, password As String) As Integer
    Dim rs As Recordset
    
    ' Query user input to database
    Dim qdf As QueryDef
    Set qdf = db.CreateQueryDef("", "SELECT * FROM staff WHERE username=[_uname] AND password=[_pw]")
    qdf.Parameters("_uname") = username
    qdf.Parameters("_pw") = password

    Set rs = qdf.OpenRecordset

    If rs.BOF And rs.EOF Then ' If account not exist in database

        LoginUser = 0
        
        ' clean up
        rs.Close
        Set rs = Nothing
        
        Exit Function
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
    End If
    LoginUser = 1
End Function

Public Sub LogoutUser()
        ' Depopulate CurrentUser properties
        With CurrentUser
            .isAuthenticated = False
            .id = 0
            .username = ""
            .password = ""
            .isAdmin = False
        End With
End Sub

