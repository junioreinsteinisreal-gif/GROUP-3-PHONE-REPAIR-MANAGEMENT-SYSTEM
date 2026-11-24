Imports Microsoft.Data.Sqlite

''' <summary>Collection reminders: overdue, due tomorrow, awaiting collection.</summary>
Public Module ReminderService

    Public ReadOnly Categories As String() = {"Overdue", "Due tomorrow", "Awaiting collection"}

    Private Const Cols As String =
        "r.RepairId, r.RepairCode, r.CustomerName, r.CustomerPhone, " &
        "r.Brand || ' ' || r.Model AS Device, r.Status, " &
        "r.ExpectedCollectionDate AS DueDate, " &
        "CAST(r.EstimatedCost - COALESCE((SELECT SUM(p.Amount) FROM Payments p " &
        "WHERE p.RepairCode = r.RepairCode), 0) AS REAL) AS Balance, " &
        "e.FullName AS Employee "

    Private Const FromJoin As String =
        "FROM Repairs r JOIN Employees e ON e.EmployeeId = r.EmployeeId "

    Private Const NotFinished As String =
        "r.Status NOT IN ('Collected','Unrepairable') "

    ''' <summary>
    ''' Repairs needing attention. Collected repairs are never included in the reminder
    ''' categories. mineOnly limits it to the current employee's own repairs.
    ''' includeRecent adds repairs collected in the last 7 days.
    ''' </summary>
    Public Function GetReminders(category As String, mineOnly As Boolean,
                                 includeRecent As Boolean) As DataTable
        SessionManager.RequireLogin()
        If SessionManager.IsEmployee Then mineOnly = True

        Dim ps As New List(Of SqliteParameter)
        ps.Add(DatabaseManager.P("@today", DatabaseManager.TodayText()))
        ps.Add(DatabaseManager.P("@tomorrow", Date.Today.AddDays(1).ToString("yyyy-MM-dd")))

        Dim owner As String = ""
        If mineOnly Then
            owner = "AND r.EmployeeId = @me "
            ps.Add(DatabaseManager.P("@me", SessionManager.CurrentUserId))
        End If

        Dim sql As String =
            "SELECT t.Category, t.RepairCode, t.CustomerName, t.CustomerPhone, t.Device, " &
            "t.Status, t.DueDate, t.Balance, t.Employee FROM (" &
            "SELECT 'Overdue' AS Category, 1 AS Sort, " & Cols & FromJoin &
            "WHERE " & NotFinished & "AND r.ExpectedCollectionDate < @today " & owner &
            "UNION ALL SELECT 'Due tomorrow', 2, " & Cols & FromJoin &
            "WHERE " & NotFinished & "AND r.ExpectedCollectionDate = @tomorrow " & owner &
            "UNION ALL SELECT 'Awaiting collection', 3, " & Cols & FromJoin &
            "WHERE r.Status = 'Completed' " & owner

        If includeRecent Then
            sql &= "UNION ALL SELECT 'Recently completed', 4, " & Cols & FromJoin &
                   "WHERE r.Status = 'Collected' AND substr(r.CollectedAt, 1, 10) >= @weekago " & owner
            ps.Add(DatabaseManager.P("@weekago", Date.Today.AddDays(-7).ToString("yyyy-MM-dd")))
        End If

        sql &= ") t "
        If category <> "" AndAlso category <> "All" Then
            sql &= "WHERE t.Category = @cat "
            ps.Add(DatabaseManager.P("@cat", category))
        End If
        sql &= "ORDER BY t.Sort, t.DueDate, t.RepairId"

        Return DatabaseManager.GetTable(sql, ps.ToArray())
    End Function

    ''' <summary>Whole-shop counts: {overdue, due tomorrow, awaiting collection}.</summary>
    Public Function GetCounts() As Integer()
        Dim dt As DataTable = GetReminders("All", False, False)
        Dim overdue As Integer = 0
        Dim tomorrow As Integer = 0
        Dim awaiting As Integer = 0
        For Each r As DataRow In dt.Rows
            Select Case r("Category").ToString()
                Case "Overdue"
                    overdue += 1
                Case "Due tomorrow"
                    tomorrow += 1
                Case "Awaiting collection"
                    awaiting += 1
            End Select
        Next
        Return New Integer() {overdue, tomorrow, awaiting}
    End Function

End Module