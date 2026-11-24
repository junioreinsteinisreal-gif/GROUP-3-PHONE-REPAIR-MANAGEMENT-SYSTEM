Imports Microsoft.Data.Sqlite

''' <summary>All the filters that can be combined on the history screen.</summary>
Public Class ActivityFilter
    Public Property SpecificDate As Date?
    Public Property DayName As String = ""
    Public Property FromDate As Date?
    Public Property ToDate As Date?
    Public Property EmployeeName As String = ""
    Public Property EmployeeId As Integer = 0
    Public Property ActionType As String = ""
    Public Property RepairCode As String = ""
    Public Property NewestFirst As Boolean = True
End Class

''' <summary>Reading activity history and removing entries from an employee's own view.</summary>
Public Module ActivityService

    Public Const MaxRows As Integer = 5000
    Public ReadOnly DayNames As String() =
        {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}

    Private Function EscapeLike(text As String) As String
        Return text.Replace("\", "\\").Replace("%", "\%").Replace("_", "\_")
    End Function

    ''' <summary>
    ''' Employees only ever get their own visible entries. Administrators get everything,
    ''' including entries employees removed from their own view.
    ''' </summary>
    Public Function GetActivities(f As ActivityFilter) As DataTable
        SessionManager.RequireLogin()
        Dim conds As New List(Of String)
        Dim ps As New List(Of SqliteParameter)

        If SessionManager.IsEmployee Then
            conds.Add("a.EmployeeId = @self AND a.HiddenFromEmployee = 0")
            ps.Add(DatabaseManager.P("@self", SessionManager.CurrentUserId))
        Else
            If f.EmployeeId > 0 Then
                conds.Add("a.EmployeeId = @eid")
                ps.Add(DatabaseManager.P("@eid", f.EmployeeId))
            End If
            If f.EmployeeName <> "" Then
                conds.Add("a.EmployeeName LIKE @en ESCAPE '\'")
                ps.Add(DatabaseManager.P("@en", "%" & EscapeLike(f.EmployeeName) & "%"))
            End If
        End If

        If f.SpecificDate.HasValue Then
            conds.Add("a.ActivityDate = @d")
            ps.Add(DatabaseManager.P("@d", f.SpecificDate.Value.ToString("yyyy-MM-dd")))
        End If

        Dim idx As Integer = Array.IndexOf(DayNames, f.DayName)
        If idx >= 0 Then
            ' SQLite weekday: 0 = Sunday ... 6 = Saturday
            conds.Add("CAST(strftime('%w', a.ActivityDate) AS INTEGER) = @dow")
            ps.Add(DatabaseManager.P("@dow", (idx + 1) Mod 7))
        End If

        If f.FromDate.HasValue Then
            conds.Add("a.ActivityDate >= @from")
            ps.Add(DatabaseManager.P("@from", f.FromDate.Value.ToString("yyyy-MM-dd")))
        End If
        If f.ToDate.HasValue Then
            conds.Add("a.ActivityDate <= @to")
            ps.Add(DatabaseManager.P("@to", f.ToDate.Value.ToString("yyyy-MM-dd")))
        End If
        If f.ActionType <> "" Then
            conds.Add("a.ActionType = @type")
            ps.Add(DatabaseManager.P("@type", f.ActionType))
        End If
        If f.RepairCode <> "" Then
            conds.Add("a.RepairCode LIKE @rc ESCAPE '\'")
            ps.Add(DatabaseManager.P("@rc", "%" & EscapeLike(f.RepairCode) & "%"))
        End If

        Dim dir As String = If(f.NewestFirst, "DESC", "ASC")
        Dim sql As String =
            "SELECT a.ActivityId, a.EmployeeId, a.EmployeeName, a.ActionType, " &
            "COALESCE(a.RepairCode, '') AS RepairCode, a.Description, a.ActivityDate, " &
            "CASE strftime('%w', a.ActivityDate) WHEN '0' THEN 'Sunday' WHEN '1' THEN 'Monday' " &
            "WHEN '2' THEN 'Tuesday' WHEN '3' THEN 'Wednesday' WHEN '4' THEN 'Thursday' " &
            "WHEN '5' THEN 'Friday' ELSE 'Saturday' END AS DayName, " &
            "a.ActivityTime, " &
            "CASE WHEN a.WasSuccessful = 1 THEN 'Success' ELSE 'Failed' END AS Result, " &
            "CASE WHEN a.HiddenFromEmployee = 1 THEN 'Yes' ELSE '' END AS Removed " &
            "FROM ActivityLogs a" &
            If(conds.Count > 0, " WHERE " & String.Join(" AND ", conds), "") &
            " ORDER BY a.Timestamp " & dir & ", a.ActivityId " & dir &
            " LIMIT " & MaxRows

        Return DatabaseManager.GetTable(sql, ps.ToArray())
    End Function

    ''' <summary>Distinct action types the current user may see (for the filter list).</summary>
    Public Function GetActionTypes() As List(Of String)
        SessionManager.RequireLogin()
        Dim dt As DataTable
        If SessionManager.IsEmployee Then
            dt = DatabaseManager.GetTable(
                "SELECT DISTINCT ActionType FROM ActivityLogs WHERE EmployeeId = @id " &
                "AND HiddenFromEmployee = 0 ORDER BY ActionType",
                DatabaseManager.P("@id", SessionManager.CurrentUserId))
        Else
            dt = DatabaseManager.GetTable(
                "SELECT DISTINCT ActionType FROM ActivityLogs ORDER BY ActionType")
        End If
        Dim list As New List(Of String)
        For Each r As DataRow In dt.Rows
            list.Add(r("ActionType").ToString())
        Next
        Return list
    End Function

    ''' <summary>
    ''' Hides an entry from the employee's own history. The original row is kept and a full
    ''' copy goes to AdministratorHistory in the SAME transaction. Returns "" or an error.
    ''' </summary>
    Public Function RemoveFromMyHistory(activityId As Long) As String
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can remove entries from their own history.")
        End If

        Dim stamp As String = DatabaseManager.NowText()
        Dim action As String = ""

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Dim found As Boolean = False
                    Dim code As String = ""
                    Dim desc As String = ""
                    Dim origStamp As String = ""
                    Dim ok As Integer = 1

                    Using cmd As SqliteCommand = conn.CreateCommand()
                        cmd.Transaction = tx
                        cmd.CommandText =
                            "SELECT ActionType, RepairCode, Description, Timestamp, WasSuccessful " &
                            "FROM ActivityLogs WHERE ActivityId = @id AND EmployeeId = @emp " &
                            "AND HiddenFromEmployee = 0"
                        cmd.Parameters.Add(DatabaseManager.P("@id", activityId))
                        cmd.Parameters.Add(DatabaseManager.P("@emp", SessionManager.CurrentUserId))
                        Using rd As SqliteDataReader = cmd.ExecuteReader()
                            If rd.Read() Then
                                found = True
                                action = rd.GetString(0)
                                code = If(rd.IsDBNull(1), "", rd.GetString(1))
                                desc = rd.GetString(2)
                                origStamp = rd.GetString(3)
                                ok = rd.GetInt32(4)
                            End If
                        End Using
                    End Using

                    If Not found Then Return "That entry is no longer available."

                    Using cmd As SqliteCommand = conn.CreateCommand()
                        cmd.Transaction = tx
                        cmd.CommandText = "UPDATE ActivityLogs SET HiddenFromEmployee = 1 WHERE ActivityId = @id"
                        cmd.Parameters.Add(DatabaseManager.P("@id", activityId))
                        cmd.ExecuteNonQuery()
                    End Using

                    Dim original As String =
                        "ActivityId=" & activityId & " | Employee=" & SessionManager.CurrentUserName &
                        " (ID " & SessionManager.CurrentUserId & ") | Action=" & action &
                        " | Repair=" & code & " | Description=" & desc &
                        " | Timestamp=" & origStamp & " | Success=" & If(ok = 1, "Yes", "No")

                    Using cmd As SqliteCommand = conn.CreateCommand()
                        cmd.Transaction = tx
                        cmd.CommandText =
                            "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, " &
                            "PerformedByRole, RepairCode, OriginalData, Details, EventAt) " &
                            "VALUES ('ActivityRemoved', @id, @n, 'Employee', @code, @orig, @det, @t)"
                        cmd.Parameters.Add(DatabaseManager.P("@id", SessionManager.CurrentUserId))
                        cmd.Parameters.Add(DatabaseManager.P("@n", SessionManager.CurrentUserName))
                        cmd.Parameters.Add(DatabaseManager.P("@code", If(code = "", Nothing, code)))
                        cmd.Parameters.Add(DatabaseManager.P("@orig", original))
                        cmd.Parameters.Add(DatabaseManager.P("@det",
                            SessionManager.CurrentUserName & " removed activity #" & activityId &
                            " (" & action & ") from their own history."))
                        cmd.Parameters.Add(DatabaseManager.P("@t", stamp))
                        cmd.ExecuteNonQuery()
                    End Using

                    tx.Commit()
                End Using
            End Using
        Catch ex As Exception
            Return "The entry could not be removed. (" & ex.Message & ")"
        End Try

        ' Logged AFTER the commit (SQLite allows one writer at a time).
        ActivityLogger.Log("RemoveHistory",
                           "Removed activity #" & activityId & " (" & action & ") from own history.")
        Return ""
    End Function

End Module