Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Data.Sqlite

''' <summary>Administrator-only employee management.</summary>
Public Module EmployeeService

    Private Function EscapeLike(text As String) As String
        Return text.Replace("\", "\\").Replace("%", "\%").Replace("_", "\_")
    End Function

    Private Sub Run(conn As SqliteConnection, tx As SqliteTransaction,
                    sql As String, ParamArray params() As SqliteParameter)
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText = sql
            For Each prm As SqliteParameter In params
                cmd.Parameters.Add(prm)
            Next
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Function LookupName(conn As SqliteConnection, tx As SqliteTransaction,
                                id As Integer) As String
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText = "SELECT FullName FROM Employees WHERE EmployeeId = @id"
            cmd.Parameters.Add(DatabaseManager.P("@id", id))
            Dim o As Object = cmd.ExecuteScalar()
            If o Is Nothing OrElse IsDBNull(o) Then Return ""
            Return o.ToString()
        End Using
    End Function

    ''' <summary>Runs the work in one transaction. Returns "" or an error message.</summary>
    Private Function InTransaction(work As Action(Of SqliteConnection, SqliteTransaction)) As String
        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        work(conn, tx)
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
            Return ""
        Catch ex As InvalidOperationException
            Return ex.Message
        Catch ex As Exception
            Return "The change could not be saved. (" & ex.Message & ")"
        End Try
    End Function

    Private Function CodeInUse(accessCode As String, exceptId As Integer) As Boolean
        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT EmployeeId, AccessCodeHash, AccessCodeSalt FROM Employees")
        For Each r As DataRow In dt.Rows
            If Convert.ToInt32(r("EmployeeId")) <> exceptId AndAlso
               PasswordHasher.Verify(accessCode, r("AccessCodeSalt").ToString(),
                                     r("AccessCodeHash").ToString()) Then
                Return True
            End If
        Next
        Return False
    End Function

    ' ---------- Lists and statistics ----------

    Public Function GetEmployees(searchText As String) As DataTable
        SessionManager.RequireAdmin()
        Dim q As String = If(searchText, "").Trim()
        Return DatabaseManager.GetTable(
            "SELECT EmployeeId, FullName, '****' AS AccessCode, " &
            "substr(RegisteredAt, 1, 10) AS RegisteredDate, " &
            "CASE WHEN IsActive = 1 THEN 'Active' ELSE 'Inactive' END AS StatusText " &
            "FROM Employees WHERE (@q = '' OR FullName LIKE @like ESCAPE '\' " &
            "OR CAST(EmployeeId AS TEXT) = @q) ORDER BY EmployeeId",
            DatabaseManager.P("@q", q),
            DatabaseManager.P("@like", "%" & EscapeLike(q) & "%"))
    End Function

    ''' <summary>Returns {total, active, inactive}.</summary>
    Public Function GetStats() As Integer()
        SessionManager.RequireAdmin()
        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT COUNT(*) AS Total, " &
            "COALESCE(SUM(CASE WHEN IsActive = 1 THEN 1 ELSE 0 END), 0) AS Active FROM Employees")
        Dim total As Integer = Convert.ToInt32(dt.Rows(0)("Total"))
        Dim active As Integer = Convert.ToInt32(dt.Rows(0)("Active"))
        Return New Integer() {total, active, total - active}
    End Function

    ' ---------- Changes ----------

    Public Function RegisterEmployee(fullName As String, accessCode As String) As String
        SessionManager.RequireAdmin()
        Dim name As String = If(fullName, "").Trim()
        If name = "" Then Return "The employee name is required."
        If name.Length > 80 Then Return "The name is too long (80 characters maximum)."
        If accessCode Is Nothing OrElse Not Regex.IsMatch(accessCode, "^\d{4}$") Then
            Return "The access code must be exactly 4 digits."
        End If
        If CodeInUse(accessCode, 0) Then Return "That access code is already in use. Choose another."

        Dim salt As String = PasswordHasher.GenerateSalt()
        Dim hash As String = PasswordHasher.HashSecret(accessCode, salt)
        Dim stamp As String = DatabaseManager.NowText()

        Return InTransaction(
            Sub(conn, tx)
                Run(conn, tx,
                    "INSERT INTO Employees (FullName, AccessCodeHash, AccessCodeSalt, " &
                    "RegisteredAt, IsActive) VALUES (@n, @h, @s, @r, 1)",
                    DatabaseManager.P("@n", name),
                    DatabaseManager.P("@h", hash),
                    DatabaseManager.P("@s", salt),
                    DatabaseManager.P("@r", stamp))
                AdminToolsService.WriteAudit(conn, tx, "EmployeeRegistered",
                                             "Registered employee '" & name & "'.")
            End Sub)
    End Function

    Public Function ResetAccessCode(employeeId As Integer, newCode As String) As String
        SessionManager.RequireAdmin()
        If newCode Is Nothing OrElse Not Regex.IsMatch(newCode, "^\d{4}$") Then
            Return "The access code must be exactly 4 digits."
        End If
        If CodeInUse(newCode, employeeId) Then Return "That access code is already in use. Choose another."

        Dim salt As String = PasswordHasher.GenerateSalt()
        Dim hash As String = PasswordHasher.HashSecret(newCode, salt)

        Return InTransaction(
            Sub(conn, tx)
                Dim empName As String = LookupName(conn, tx, employeeId)
                If empName = "" Then Throw New InvalidOperationException("That employee no longer exists.")
                Run(conn, tx,
                    "UPDATE Employees SET AccessCodeHash = @h, AccessCodeSalt = @s " &
                    "WHERE EmployeeId = @id",
                    DatabaseManager.P("@h", hash),
                    DatabaseManager.P("@s", salt),
                    DatabaseManager.P("@id", employeeId))
                AdminToolsService.WriteAudit(conn, tx, "EmployeeCodeReset",
                    "Reset the access code of employee '" & empName & "' (ID " & employeeId & ").")
            End Sub)
    End Function

    Public Function SetActive(employeeId As Integer, active As Boolean) As String
        SessionManager.RequireAdmin()
        Return InTransaction(
            Sub(conn, tx)
                Dim empName As String = LookupName(conn, tx, employeeId)
                If empName = "" Then Throw New InvalidOperationException("That employee no longer exists.")
                Run(conn, tx, "UPDATE Employees SET IsActive = @a WHERE EmployeeId = @id",
                    DatabaseManager.P("@a", If(active, 1, 0)),
                    DatabaseManager.P("@id", employeeId))
                AdminToolsService.WriteAudit(conn, tx,
                    If(active, "EmployeeActivated", "EmployeeDeactivated"),
                    If(active, "Activated", "Deactivated") & " employee '" & empName &
                    "' (ID " & employeeId & ").")
            End Sub)
    End Function

    ' ---------- Performance and revenue ----------

    ''' <summary>Readable performance report for one employee ("" if not found).</summary>
    Public Function GetPerformanceSummary(employeeId As Integer) As String
        SessionManager.RequireAdmin()
        Dim info As DataTable = DatabaseManager.GetTable(
            "SELECT FullName, RegisteredAt, IsActive FROM Employees WHERE EmployeeId = @id",
            DatabaseManager.P("@id", employeeId))
        If info.Rows.Count = 0 Then Return ""

        Dim today As String = DatabaseManager.TodayText()
        Dim s1 As Func(Of String, Double) =
            Function(q As String) Convert.ToDouble(DatabaseManager.ExecuteScalar(
                q, DatabaseManager.P("@id", employeeId)))
        Dim s2 As Func(Of String, Double) =
            Function(q As String) Convert.ToDouble(DatabaseManager.ExecuteScalar(
                q, DatabaseManager.P("@id", employeeId), DatabaseManager.P("@d", today)))

        Dim repairsAll As Double = s1("SELECT COUNT(*) FROM Repairs WHERE EmployeeId = @id")
        Dim repairsToday As Double = s2("SELECT COUNT(*) FROM Repairs WHERE EmployeeId = @id AND substr(CreatedAt, 1, 10) = @d")
        Dim payCount As Double = s1("SELECT COUNT(*) FROM Payments WHERE EmployeeId = @id")
        Dim payCountToday As Double = s2("SELECT COUNT(*) FROM Payments WHERE EmployeeId = @id AND substr(PaidAt, 1, 10) = @d")
        Dim paidAll As Double = s1("SELECT CAST(COALESCE(SUM(Amount), 0) AS REAL) FROM Payments WHERE EmployeeId = @id")
        Dim paidToday As Double = s2("SELECT CAST(COALESCE(SUM(Amount), 0) AS REAL) FROM Payments WHERE EmployeeId = @id AND substr(PaidAt, 1, 10) = @d")
        Dim refunds As Double = s1("SELECT CAST(COALESCE(SUM(Amount), 0) AS REAL) FROM Refunds WHERE EmployeeId = @id")
        Dim failed As Double = s1("SELECT COUNT(*) FROM ActivityLogs WHERE EmployeeId = @id AND WasSuccessful = 0")
        Dim logins As Double = s1("SELECT COUNT(*) FROM EmployeeSessions WHERE EmployeeId = @id")

        Dim terminations As Double = 0
        Try
            terminations = s1("SELECT COUNT(*) FROM RepairTerminations WHERE TerminatedById = @id")
        Catch
        End Try

        Dim lastLogin As String = Convert.ToString(DatabaseManager.ExecuteScalar(
            "SELECT MAX(LoginAt) FROM EmployeeSessions WHERE EmployeeId = @id",
            DatabaseManager.P("@id", employeeId)))
        If lastLogin = "" Then lastLogin = "never"

        Dim r As DataRow = info.Rows(0)
        Dim sb As New StringBuilder()
        sb.AppendLine(r("FullName").ToString() & "  (ID " & employeeId & ")")
        sb.AppendLine("Status: " & If(Convert.ToInt32(r("IsActive")) = 1, "Active", "Inactive"))
        sb.AppendLine("Registered: " & r("RegisteredAt").ToString())
        sb.AppendLine()
        sb.AppendLine("TODAY")
        sb.AppendLine("  Repairs registered: " & CInt(repairsToday))
        sb.AppendLine("  Payments received: " & CInt(payCountToday) & "  (" & UIHelpers.Money(paidToday) & ")")
        sb.AppendLine()
        sb.AppendLine("ALL TIME")
        sb.AppendLine("  Repairs registered: " & CInt(repairsAll))
        sb.AppendLine("  Payments received: " & CInt(payCount) & "  (" & UIHelpers.Money(paidAll) & ")")
        sb.AppendLine("  Refunds given: " & UIHelpers.Money(refunds))
        sb.AppendLine("  Net revenue: " & UIHelpers.Money(paidAll - refunds))
        sb.AppendLine("  Repairs terminated: " & CInt(terminations))
        sb.AppendLine("  Failed or rejected actions: " & CInt(failed))
        sb.AppendLine("  Logins: " & CInt(logins) & "  (last: " & lastLogin & ")")
        sb.Append(FeedbackService.GetEmployeeText(employeeId))
        Return sb.ToString()
    End Function

End Module