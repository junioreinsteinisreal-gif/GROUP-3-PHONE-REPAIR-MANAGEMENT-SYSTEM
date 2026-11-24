Imports System.Globalization
Imports System.Text
Imports Microsoft.Data.Sqlite

Public Class FeedbackStats
    Public Property Reviews As Integer
    Public Property AvgRating As Double
    Public Property GoodReviews As Integer
    Public Property VerifiedGood As Integer
    Public Property TipsTotal As Double
    Public Property TipsToday As Double
End Class

''' <summary>Reading customer reviews and tips. Reviews are created only by customers on the
''' review website; employees can only read their own.</summary>
Public Module FeedbackService

    Public Const MaxComment As Integer = 300

    Private Function P(name As String, value As Object) As SqliteParameter
        Return DatabaseManager.P(name, value)
    End Function

    ' ---------- Statistics ----------

    ''' <summary>employeeId 0 = everyone.</summary>
    Private Function StatsFor(employeeId As Integer) As FeedbackStats
        Dim s As New FeedbackStats()
        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT COUNT(*) AS N, CAST(COALESCE(AVG(Rating), 0) AS REAL) AS AvgR, " &
            "COALESCE(SUM(CASE WHEN Rating >= 4 THEN 1 ELSE 0 END), 0) AS Good, " &
            "COALESCE(SUM(CASE WHEN Rating >= 4 AND Verified = 1 THEN 1 ELSE 0 END), 0) AS GoodV " &
            "FROM Feedback WHERE (@e = 0 OR EmployeeId = @e)", P("@e", employeeId))
        s.Reviews = Convert.ToInt32(dt.Rows(0)("N"))
        s.AvgRating = Convert.ToDouble(dt.Rows(0)("AvgR"))
        s.GoodReviews = Convert.ToInt32(dt.Rows(0)("Good"))
        s.VerifiedGood = Convert.ToInt32(dt.Rows(0)("GoodV"))
        s.TipsTotal = Convert.ToDouble(DatabaseManager.ExecuteScalar(
            "SELECT CAST(COALESCE(SUM(Amount), 0) AS REAL) FROM Tips WHERE (@e = 0 OR EmployeeId = @e)",
            P("@e", employeeId)))
        s.TipsToday = Convert.ToDouble(DatabaseManager.ExecuteScalar(
            "SELECT CAST(COALESCE(SUM(Amount), 0) AS REAL) FROM Tips WHERE (@e = 0 OR EmployeeId = @e) " &
            "AND substr(ReceivedAt, 1, 10) = @d", P("@e", employeeId), P("@d", DatabaseManager.TodayText())))
        Return s
    End Function

    Public Function GetMyStats() As FeedbackStats
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("These figures are for employees.")
        End If
        Return StatsFor(SessionManager.CurrentUserId)
    End Function

    Public Function GetShopStats() As FeedbackStats
        SessionManager.RequireAdmin()
        Return StatsFor(0)
    End Function

    ''' <summary>Text block added to the employee performance report (administrators only).</summary>
    Public Function GetEmployeeText(employeeId As Integer) As String
        SessionManager.RequireAdmin()
        Dim s As FeedbackStats = StatsFor(employeeId)
        Dim sb As New StringBuilder()
        sb.AppendLine()
        sb.AppendLine("CUSTOMER FEEDBACK")
        sb.AppendLine("  Reviews: " & s.Reviews &
                      If(s.Reviews > 0, "  (average " & s.AvgRating.ToString("0.0", CultureInfo.InvariantCulture) & " / 5)", ""))
        sb.AppendLine("  Good reviews (4-5 stars): " & s.GoodReviews & "  (" & s.VerifiedGood & " verified)")
        sb.AppendLine("  Tips received: " & UIHelpers.Money(s.TipsTotal) & "  (today " & UIHelpers.Money(s.TipsToday) & ")")
        Return sb.ToString()
    End Function

    ' ---------- Lists ----------

    Private Function EscapeLike(text As String) As String
        Return text.Replace("\", "\\").Replace("%", "\%").Replace("_", "\_")
    End Function

    ''' <summary>Employees get only their own feedback; administrators get everything.</summary>
    Public Function GetFeedbackList(searchText As String) As DataTable
        SessionManager.RequireLogin()
        Dim owner As Integer = If(SessionManager.IsEmployee, SessionManager.CurrentUserId, 0)
        Dim q As String = If(searchText, "").Trim()
        Return DatabaseManager.GetTable(
            "SELECT f.RecordedAt, f.RepairCode, f.CustomerName, e.FullName AS Employee, f.Rating, " &
            "COALESCE(f.Comment, '') AS Comment, " &
            "CAST(COALESCE((SELECT SUM(t.Amount) FROM Tips t WHERE t.FeedbackId = f.FeedbackId), 0) AS REAL) AS Tip, " &
            "CASE WHEN f.Verified = 1 THEN 'Yes' ELSE 'No' END AS Verified " &
            "FROM Feedback f JOIN Employees e ON e.EmployeeId = f.EmployeeId " &
            "WHERE (@o = 0 OR f.EmployeeId = @o) AND (@q = '' OR f.RepairCode LIKE @like ESCAPE '\' " &
            "OR f.CustomerName LIKE @like ESCAPE '\' OR e.FullName LIKE @like ESCAPE '\' " &
            "OR f.Comment LIKE @like ESCAPE '\') ORDER BY f.FeedbackId DESC LIMIT 2000",
            P("@o", owner), P("@q", q), P("@like", "%" & EscapeLike(q) & "%"))
    End Function

    Public Function GetEmployeeSummary() As DataTable
        SessionManager.RequireAdmin()
        Return DatabaseManager.GetTable(
            "SELECT e.FullName AS Employee, " &
            "(SELECT COUNT(*) FROM Feedback f WHERE f.EmployeeId = e.EmployeeId) AS Reviews, " &
            "CAST((SELECT COALESCE(AVG(f.Rating), 0) FROM Feedback f WHERE f.EmployeeId = e.EmployeeId) AS REAL) AS AvgRating, " &
            "(SELECT COUNT(*) FROM Feedback f WHERE f.EmployeeId = e.EmployeeId AND f.Rating >= 4) AS GoodReviews, " &
            "(SELECT COUNT(*) FROM Feedback f WHERE f.EmployeeId = e.EmployeeId AND f.Rating >= 4 AND f.Verified = 1) AS VerifiedGood, " &
            "CAST((SELECT COALESCE(SUM(t.Amount), 0) FROM Tips t WHERE t.EmployeeId = e.EmployeeId) AS REAL) AS Tips " &
            "FROM Employees e ORDER BY Tips DESC, GoodReviews DESC, e.FullName")
    End Function

End Module