''' <summary>Saves employee actions permanently to ActivityLogs.</summary>
Public Module ActivityLogger

    ''' <summary>Logs an action for the currently logged-in employee.</summary>
    Public Sub Log(actionType As String, description As String,
                   Optional repairCode As String = Nothing,
                   Optional wasSuccessful As Boolean = True)
        If Not SessionManager.IsEmployee Then Return
        LogFor(SessionManager.CurrentUserId, SessionManager.CurrentUserName,
               actionType, description, repairCode, wasSuccessful)
    End Sub

    ''' <summary>Logs an action for a specific employee (used for login events).</summary>
    Public Sub LogFor(employeeId As Integer, employeeName As String,
                      actionType As String, description As String,
                      Optional repairCode As String = Nothing,
                      Optional wasSuccessful As Boolean = True)
        Try
            Dim stamp As DateTime = DateTime.Now
            DatabaseManager.ExecuteNonQuery(
                "INSERT INTO ActivityLogs (EmployeeId, EmployeeName, ActionType, " &
                "RepairCode, Description, ActivityDate, ActivityTime, Timestamp, " &
                "WasSuccessful) VALUES (@id, @name, @type, @code, @desc, @d, @t, @ts, @ok)",
                DatabaseManager.P("@id", employeeId),
                DatabaseManager.P("@name", employeeName),
                DatabaseManager.P("@type", actionType),
                DatabaseManager.P("@code", repairCode),
                DatabaseManager.P("@desc", description),
                DatabaseManager.P("@d", stamp.ToString("yyyy-MM-dd")),
                DatabaseManager.P("@t", stamp.ToString("HH:mm:ss")),
                DatabaseManager.P("@ts", stamp.ToString("yyyy-MM-dd HH:mm:ss")),
                DatabaseManager.P("@ok", If(wasSuccessful, 1, 0)))
        Catch
            ' Logging must never crash the application.
        End Try
    End Sub

End Module