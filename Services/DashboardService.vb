Imports Microsoft.Data.Sqlite

Public Class EmployeeStats
    Public Property RegisteredToday As Integer
    Public Property UpdatedToday As Integer
    Public Property CompletedToday As Integer
    Public Property Pending As Integer
    Public Property PaymentsToday As Integer
    Public Property CollectedToday As Double
    Public Property RefundedToday As Double
    Public ReadOnly Property NetToday As Double
        Get
            Return CollectedToday - RefundedToday
        End Get
    End Property
End Class

Public Class AdminStats
    Public Property RevenueToday As Double
    Public Property PaymentsToday As Integer
    Public Property RegisteredToday As Integer
    Public Property Pending As Integer
    Public Property ActiveEmployees As Integer
End Class

''' <summary>Numbers and lists for the dashboards. Employees only ever see their own.</summary>
Public Module DashboardService

    Private Function P(name As String, value As Object) As SqliteParameter
        Return DatabaseManager.P(name, value)
    End Function

    Private Function Num(sql As String, ParamArray params() As SqliteParameter) As Double
        Dim o As Object = DatabaseManager.ExecuteScalar(sql, params)
        If o Is Nothing OrElse IsDBNull(o) Then Return 0
        Return Convert.ToDouble(o)
    End Function

    Public Function GetEmployeeStats() As EmployeeStats
        SessionManager.RequireLogin()
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("This dashboard is for employees.")
        End If

        Dim id As Integer = SessionManager.CurrentUserId
        Dim today As String = DatabaseManager.TodayText()
        Dim s As New EmployeeStats()

        s.RegisteredToday = CInt(Num(
            "SELECT COUNT(*) FROM Repairs WHERE EmployeeId = @id AND substr(CreatedAt, 1, 10) = @d",
            P("@id", id), P("@d", today)))

        s.UpdatedToday = CInt(Num(
            "SELECT COUNT(DISTINCT RepairCode) FROM ActivityLogs WHERE EmployeeId = @id " &
            "AND ActivityDate = @d AND WasSuccessful = 1 AND " &
            "(ActionType IN ('StatusUpdate','UpdateRepair','Terminate') OR " &
            "(ActionType = 'Payment' AND Description NOT LIKE '%initial payment%'))",
            P("@id", id), P("@d", today)))

        s.CompletedToday = CInt(Num(
            "SELECT COUNT(DISTINCT RepairCode) FROM ActivityLogs WHERE EmployeeId = @id " &
            "AND ActivityDate = @d AND WasSuccessful = 1 AND ActionType = 'StatusUpdate' AND " &
            "(Description LIKE '%to Completed.' OR Description LIKE '%to Collected.')",
            P("@id", id), P("@d", today)))

        s.Pending = CInt(Num(
            "SELECT COUNT(*) FROM Repairs WHERE EmployeeId = @id " &
            "AND Status IN ('Received','In Progress')", P("@id", id)))

        s.PaymentsToday = CInt(Num(
            "SELECT COUNT(*) FROM Payments WHERE EmployeeId = @id AND substr(PaidAt, 1, 10) = @d",
            P("@id", id), P("@d", today)))

        s.CollectedToday = Num(
            "SELECT COALESCE(SUM(Amount), 0) FROM Payments WHERE EmployeeId = @id " &
            "AND substr(PaidAt, 1, 10) = @d", P("@id", id), P("@d", today))

        s.RefundedToday = Num(
            "SELECT COALESCE(SUM(Amount), 0) FROM Refunds WHERE EmployeeId = @id " &
            "AND substr(RefundedAt, 1, 10) = @d", P("@id", id), P("@d", today))
        Return s
    End Function

    Public Function GetAdminStats() As AdminStats
        SessionManager.RequireAdmin()
        Dim today As String = DatabaseManager.TodayText()
        Dim a As New AdminStats()

        Dim paid As Double = Num(
            "SELECT COALESCE(SUM(Amount), 0) FROM Payments WHERE substr(PaidAt, 1, 10) = @d",
            P("@d", today))
        Dim refunded As Double = Num(
            "SELECT COALESCE(SUM(Amount), 0) FROM Refunds WHERE substr(RefundedAt, 1, 10) = @d",
            P("@d", today))
        a.RevenueToday = paid - refunded

        a.PaymentsToday = CInt(Num(
            "SELECT COUNT(*) FROM Payments WHERE substr(PaidAt, 1, 10) = @d", P("@d", today)))
        a.RegisteredToday = CInt(Num(
            "SELECT COUNT(*) FROM Repairs WHERE substr(CreatedAt, 1, 10) = @d", P("@d", today)))
        a.Pending = CInt(Num(
            "SELECT COUNT(*) FROM Repairs WHERE Status IN ('Received','In Progress')"))
        a.ActiveEmployees = CInt(Num("SELECT COUNT(*) FROM Employees WHERE IsActive = 1"))
        Return a
    End Function

    ''' <summary>Employees see their own entries; administrators see everyone's.</summary>
    Public Function GetRecentActivity(limit As Integer) As DataTable
        SessionManager.RequireLogin()
        Dim filter As String = ""
        Dim ps As New List(Of SqliteParameter)
        If SessionManager.IsEmployee Then
            filter = "WHERE a.EmployeeId = @id AND a.HiddenFromEmployee = 0 "
            ps.Add(P("@id", SessionManager.CurrentUserId))
        End If
        Return DatabaseManager.GetTable(
            "SELECT substr(a.Timestamp, 6, 11) AS WhenText, a.EmployeeName, a.Description, " &
            "CASE WHEN a.WasSuccessful = 1 THEN '' ELSE 'Failed' END AS Result " &
            "FROM ActivityLogs a " & filter & "ORDER BY a.ActivityId DESC LIMIT " & limit,
            ps.ToArray())
    End Function

    Public Function GetRecentPayments(limit As Integer) As DataTable
        SessionManager.RequireLogin()
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Payment records on the dashboard are for employees.")
        End If
        Return DatabaseManager.GetTable(
            "SELECT substr(p.PaidAt, 6, 11) AS WhenText, p.RepairCode, " &
            "CAST(p.Amount AS REAL) AS Amount, p.PaymentMethod, " &
            "COALESCE(rc.ReceiptNumber, '') AS ReceiptNumber " &
            "FROM Payments p LEFT JOIN Receipts rc ON rc.PaymentId = p.PaymentId " &
            "WHERE p.EmployeeId = @id ORDER BY p.PaymentId DESC LIMIT " & limit,
            P("@id", SessionManager.CurrentUserId))
    End Function

    ''' <summary>Administrator table: every active employee's figures for today.</summary>
    Public Function GetEmployeeRevenueToday() As DataTable
        SessionManager.RequireAdmin()
        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT e.FullName AS Employee, " &
            "(SELECT COUNT(*) FROM Repairs r WHERE r.EmployeeId = e.EmployeeId " &
            " AND substr(r.CreatedAt, 1, 10) = @d) AS Registered, " &
            "(SELECT COUNT(*) FROM Payments p WHERE p.EmployeeId = e.EmployeeId " &
            " AND substr(p.PaidAt, 1, 10) = @d) AS Payments, " &
            "CAST((SELECT COALESCE(SUM(p.Amount), 0) FROM Payments p WHERE p.EmployeeId = e.EmployeeId " &
            " AND substr(p.PaidAt, 1, 10) = @d) AS REAL) AS Collected, " &
            "CAST((SELECT COALESCE(SUM(f.Amount), 0) FROM Refunds f WHERE f.EmployeeId = e.EmployeeId " &
            " AND substr(f.RefundedAt, 1, 10) = @d) AS REAL) AS Refunded " &
            "FROM Employees e WHERE e.IsActive = 1 ORDER BY Collected DESC, e.FullName",
            P("@d", DatabaseManager.TodayText()))
        dt.Columns.Add("Net", GetType(Double), "Collected - Refunded")
        Return dt
    End Function

End Module