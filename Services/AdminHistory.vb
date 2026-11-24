Imports Microsoft.Data.Sqlite

''' <summary>Administrator-only reading of the audit trail and deleted records.</summary>
Public Module AdminHistoryService

    Public Const MaxRows As Integer = 5000

    Private Function EscapeLike(text As String) As String
        Return text.Replace("\", "\\").Replace("%", "\%").Replace("_", "\_")
    End Function

    Public Function GetEventTypes() As List(Of String)
        SessionManager.RequireAdmin()
        Dim list As New List(Of String)
        For Each r As DataRow In DatabaseManager.GetTable(
            "SELECT DISTINCT EventType FROM AdministratorHistory ORDER BY EventType").Rows
            list.Add(r("EventType").ToString())
        Next
        Return list
    End Function

    Public Function GetEvents(searchText As String, eventType As String,
                              fromDate As Date?, toDate As Date?) As DataTable
        SessionManager.RequireAdmin()

        Dim conds As New List(Of String)
        Dim ps As New List(Of SqliteParameter)

        Dim q As String = If(searchText, "").Trim()
        If q <> "" Then
            conds.Add("(PerformedByName LIKE @q ESCAPE '\' OR RepairCode LIKE @q ESCAPE '\' " &
                      "OR Details LIKE @q ESCAPE '\' OR OriginalData LIKE @q ESCAPE '\')")
            ps.Add(DatabaseManager.P("@q", "%" & EscapeLike(q) & "%"))
        End If
        If eventType <> "" Then
            conds.Add("EventType = @t")
            ps.Add(DatabaseManager.P("@t", eventType))
        End If
        If fromDate.HasValue Then
            conds.Add("substr(EventAt, 1, 10) >= @from")
            ps.Add(DatabaseManager.P("@from", fromDate.Value.ToString("yyyy-MM-dd")))
        End If
        If toDate.HasValue Then
            conds.Add("substr(EventAt, 1, 10) <= @to")
            ps.Add(DatabaseManager.P("@to", toDate.Value.ToString("yyyy-MM-dd")))
        End If

        Dim sql As String =
            "SELECT HistoryId, EventAt, EventType, PerformedByName, PerformedByRole, " &
            "COALESCE(RepairCode, '') AS RepairCode, Details, " &
            "COALESCE(OriginalData, '') AS OriginalData FROM AdministratorHistory" &
            If(conds.Count > 0, " WHERE " & String.Join(" AND ", conds), "") &
            " ORDER BY HistoryId DESC LIMIT " & MaxRows
        Return DatabaseManager.GetTable(sql, ps.ToArray())
    End Function

    Public Function GetDeletedRepairs(searchText As String) As DataTable
        SessionManager.RequireAdmin()
        Dim q As String = If(searchText, "").Trim()
        Return DatabaseManager.GetTable(
            "SELECT d.DeletedAt, d.RepairCode, d.CustomerName, d.CustomerPhone, " &
            "d.Brand || ' ' || d.Model AS Device, d.ProblemDescription, " &
            "CAST(d.EstimatedCost AS REAL) AS EstimatedCost, d.Status, " &
            "CAST(d.TotalPaidAtDeletion AS REAL) AS TotalPaid, " &
            "COALESCE(o.FullName, '') AS RegisteredBy, d.DeletedByName, " &
            "COALESCE(d.DeletionReason, '') AS Reason " &
            "FROM DeletedRepairs d LEFT JOIN Employees o ON o.EmployeeId = d.OriginalEmployeeId " &
            "WHERE (@q = '' OR d.RepairCode LIKE @like ESCAPE '\' " &
            "OR d.CustomerName LIKE @like ESCAPE '\' OR d.CustomerPhone LIKE @like ESCAPE '\' " &
            "OR d.DeletedByName LIKE @like ESCAPE '\' OR d.DeletionReason LIKE @like ESCAPE '\') " &
            "ORDER BY d.DeletedId DESC LIMIT " & MaxRows,
            DatabaseManager.P("@q", q),
            DatabaseManager.P("@like", "%" & EscapeLike(q) & "%"))
    End Function

End Module