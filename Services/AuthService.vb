Imports System.Text.RegularExpressions

''' <summary>Login, logout, lockout and registration logic.</summary>
Public Module AuthService

    Private Const MaxFailedAttempts As Integer = 5
    Private Const LockoutSeconds As Integer = 60

    Private failedAttempts As Integer = 0
    Private lockedUntil As DateTime = DateTime.MinValue

    ' ---------- Lockout ----------

    Public Function IsLockedOut() As Boolean
        Return DateTime.Now < lockedUntil
    End Function

    Public Function LockoutSecondsRemaining() As Integer
        If Not IsLockedOut() Then Return 0
        Return CInt(Math.Ceiling((lockedUntil - DateTime.Now).TotalSeconds))
    End Function

    Private Sub RegisterFailure()
        failedAttempts += 1
        If failedAttempts >= MaxFailedAttempts Then
            lockedUntil = DateTime.Now.AddSeconds(LockoutSeconds)
            failedAttempts = 0
        End If
    End Sub

    Private Sub ClearFailures()
        failedAttempts = 0
        lockedUntil = DateTime.MinValue
    End Sub

    ' ---------- Administrator ----------

    Public Function AdminExists() As Boolean
        Dim count As Integer = Convert.ToInt32(DatabaseManager.ExecuteScalar(
            "SELECT COUNT(*) FROM Administrators WHERE IsActive = 1"))
        Return count > 0
    End Function

    ''' <summary>Creates an administrator. Returns "" on success, else an error message.</summary>
    Public Function CreateAdministrator(username As String, password As String) As String
        If String.IsNullOrWhiteSpace(username) OrElse username.Trim().Length < 3 Then
            Return "Username must be at least 3 characters."
        End If
        If password Is Nothing OrElse password.Length < 8 Then
            Return "Password must be at least 8 characters."
        End If

        Dim exists As Integer = Convert.ToInt32(DatabaseManager.ExecuteScalar(
            "SELECT COUNT(*) FROM Administrators WHERE Username = @u",
            DatabaseManager.P("@u", username.Trim())))
        If exists > 0 Then Return "That username already exists."

        Dim salt As String = PasswordHasher.GenerateSalt()
        Dim hash As String = PasswordHasher.HashSecret(password, salt)

        DatabaseManager.ExecuteNonQuery(
            "INSERT INTO Administrators (Username, PasswordHash, PasswordSalt, IsActive, CreatedAt) " &
            "VALUES (@u, @h, @s, 1, @c)",
            DatabaseManager.P("@u", username.Trim()),
            DatabaseManager.P("@h", hash),
            DatabaseManager.P("@s", salt),
            DatabaseManager.P("@c", DatabaseManager.NowText()))
        Return ""
    End Function

    Public Function LoginAdministrator(username As String, password As String,
                                       ByRef message As String) As Boolean
        message = ""
        If IsLockedOut() Then
            message = "Too many failed attempts. Try again in " &
                      LockoutSecondsRemaining() & " seconds."
            Return False
        End If
        If String.IsNullOrWhiteSpace(username) OrElse String.IsNullOrEmpty(password) Then
            message = "Enter your username and password."
            Return False
        End If

        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT AdminId, Username, PasswordHash, PasswordSalt FROM Administrators " &
            "WHERE Username = @u AND IsActive = 1",
            DatabaseManager.P("@u", username.Trim()))

        If dt.Rows.Count = 0 Then
            ' Do equal work so response time does not reveal valid usernames.
            PasswordHasher.HashSecret(password, PasswordHasher.GenerateSalt())
            RegisterFailure()
            LogAdminEvent("AdminLoginFailed", 0, username.Trim(),
                          "Failed administrator login attempt.")
            message = "Invalid username or password."
            Return False
        End If

        Dim row As DataRow = dt.Rows(0)
        Dim adminId As Integer = Convert.ToInt32(row("AdminId"))
        Dim adminName As String = row("Username").ToString()

        If Not PasswordHasher.Verify(password, row("PasswordSalt").ToString(),
                                     row("PasswordHash").ToString()) Then
            RegisterFailure()
            LogAdminEvent("AdminLoginFailed", adminId, adminName,
                          "Wrong password entered.")
            message = "Invalid username or password."
            Return False
        End If

        ClearFailures()
        SessionManager.Start("Administrator", adminId, adminName, 0)
        LogAdminEvent("AdminLogin", adminId, adminName, "Administrator logged in.")
        Return True
    End Function

    ' ---------- Employee ----------

    ''' <summary>Registers an employee. Returns "" on success, else an error message.</summary>
    Public Function RegisterEmployee(fullName As String, accessCode As String) As String
        If String.IsNullOrWhiteSpace(fullName) Then Return "Employee name is required."
        If accessCode Is Nothing OrElse Not Regex.IsMatch(accessCode, "^\d{4}$") Then
            Return "Access code must be exactly 4 digits."
        End If

        ' Codes are salted, so uniqueness is checked by testing each existing hash.
        Dim existing As DataTable = DatabaseManager.GetTable(
            "SELECT AccessCodeHash, AccessCodeSalt FROM Employees")
        For Each r As DataRow In existing.Rows
            If PasswordHasher.Verify(accessCode, r("AccessCodeSalt").ToString(),
                                     r("AccessCodeHash").ToString()) Then
                Return "That access code is already in use. Choose another."
            End If
        Next

        Dim salt As String = PasswordHasher.GenerateSalt()
        Dim hash As String = PasswordHasher.HashSecret(accessCode, salt)

        DatabaseManager.ExecuteNonQuery(
            "INSERT INTO Employees (FullName, AccessCodeHash, AccessCodeSalt, RegisteredAt, IsActive) " &
            "VALUES (@n, @h, @s, @r, 1)",
            DatabaseManager.P("@n", fullName.Trim()),
            DatabaseManager.P("@h", hash),
            DatabaseManager.P("@s", salt),
            DatabaseManager.P("@r", DatabaseManager.NowText()))
        Return ""
    End Function

    Public Function LoginEmployee(accessCode As String,
                                  ByRef message As String) As Boolean
        message = ""
        If IsLockedOut() Then
            message = "Too many failed attempts. Try again in " &
                      LockoutSecondsRemaining() & " seconds."
            Return False
        End If

        ' After the day is closed, employees wait until 6:00 AM.
        EndOfDayService.AutoCloseIfDue(True)
        Dim closedReason As String = EndOfDayService.EmployeeLoginBlockReason()
        If closedReason <> "" Then
            message = closedReason
            Return False
        End If

        If accessCode Is Nothing OrElse Not Regex.IsMatch(accessCode, "^\d{4}$") Then
            message = "Enter your 4-digit access code."
            Return False
        End If

        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT EmployeeId, FullName, AccessCodeHash, AccessCodeSalt, IsActive FROM Employees")

        For Each r As DataRow In dt.Rows
            If PasswordHasher.Verify(accessCode, r("AccessCodeSalt").ToString(),
                                     r("AccessCodeHash").ToString()) Then

                Dim empId As Integer = Convert.ToInt32(r("EmployeeId"))
                Dim empName As String = r("FullName").ToString()

                If Convert.ToInt32(r("IsActive")) <> 1 Then
                    ActivityLogger.LogFor(empId, empName, "LoginBlocked",
                        "Login attempt by a deactivated employee.", Nothing, False)
                    message = "This account has been deactivated. Contact the administrator."
                    Return False
                End If

                Dim sessionId As Long = Convert.ToInt64(DatabaseManager.ExecuteScalar(
                    "INSERT INTO EmployeeSessions (EmployeeId, LoginAt) " &
                    "VALUES (@e, @t) RETURNING SessionId",
                    DatabaseManager.P("@e", empId),
                    DatabaseManager.P("@t", DatabaseManager.NowText())))

                ClearFailures()
                SessionManager.Start("Employee", empId, empName, sessionId)
                ActivityLogger.Log("Login", empName & " logged in.")
                Return True
            End If
        Next

        RegisterFailure()
        message = "Invalid access code."
        Return False
    End Function

    ' ---------- Logout ----------

    Public Sub Logout(Optional reason As String = "Manual")
        If SessionManager.IsEmployee Then
            ActivityLogger.Log("Logout", SessionManager.CurrentUserName &
                               " logged out (" & reason & ").")
            DatabaseManager.ExecuteNonQuery(
                "UPDATE EmployeeSessions SET LogoutAt = @t, LogoutReason = @r " &
                "WHERE SessionId = @id AND LogoutAt IS NULL",
                DatabaseManager.P("@t", DatabaseManager.NowText()),
                DatabaseManager.P("@r", reason),
                DatabaseManager.P("@id", SessionManager.CurrentSessionId))
        ElseIf SessionManager.IsAdmin Then
            LogAdminEvent("AdminLogout", SessionManager.CurrentUserId,
                          SessionManager.CurrentUserName,
                          "Administrator logged out (" & reason & ").")
        End If
        SessionManager.Clear()
    End Sub

    ' ---------- Admin audit helper ----------

    Private Sub LogAdminEvent(eventType As String, userId As Integer,
                              userName As String, details As String)
        Try
            DatabaseManager.ExecuteNonQuery(
                "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, " &
                "PerformedByRole, Details, EventAt) VALUES (@e, @id, @n, 'Administrator', @d, @t)",
                DatabaseManager.P("@e", eventType),
                DatabaseManager.P("@id", If(userId = 0, Nothing, CObj(userId))),
                DatabaseManager.P("@n", userName),
                DatabaseManager.P("@d", details),
                DatabaseManager.P("@t", DatabaseManager.NowText()))
        Catch
        End Try
    End Sub

End Module