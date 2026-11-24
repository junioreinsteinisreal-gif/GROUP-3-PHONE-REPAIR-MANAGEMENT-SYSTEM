Imports System.Globalization
Imports System.IO
Imports ClosedXML.Excel
Imports DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing
Imports Microsoft.Data.Sqlite

''' <summary>The headline numbers for one business day.</summary>
Public Class DayFigures
    Public Property Collected As Double
    Public Property PaymentCount As Integer
    Public Property Refunded As Double
    Public Property RefundCount As Integer
    Public Property Registered As Integer
    Public Property Completed As Integer
    Public Property Outstanding As Double

    Public ReadOnly Property NetRevenue As Double
        Get
            Return Collected - Refunded
        End Get
    End Property
End Class

''' <summary>End-of-day figures, closing (manual and automatic), the login lock and the Excel report.</summary>
Public Module EndOfDayService

    ' Change CloseTime only for testing, then put it back to 23:30.
    Public ReadOnly CloseTime As TimeSpan = New TimeSpan(23, 30, 0)
    Public ReadOnly ReopenTime As TimeSpan = New TimeSpan(6, 0, 0)

    Private lastAutoAttempt As DateTime = DateTime.MinValue

    Private Function P(name As String, value As Object) As SqliteParameter
        Return DatabaseManager.P(name, value)
    End Function

    Private Function Num(sql As String, ParamArray params() As SqliteParameter) As Double
        Dim o As Object = DatabaseManager.ExecuteScalar(sql, params)
        If o Is Nothing OrElse IsDBNull(o) Then Return 0
        Return Convert.ToDouble(o)
    End Function


    Private Function Run(conn As SqliteConnection, tx As SqliteTransaction,
                         sql As String, ParamArray params() As SqliteParameter) As Integer
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText = sql
            For Each prm As SqliteParameter In params
                cmd.Parameters.Add(prm)
            Next
            Return cmd.ExecuteNonQuery()
        End Using
    End Function

    Private Function DateText(d As Date) As String
        Return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
    End Function

    Public Function ReportsFolder() As String
        Dim folder As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "PhoneRepairReports")
        Directory.CreateDirectory(folder)
        Return folder
    End Function

    ' ---------- Core queries (no session checks; the public wrappers check) ----------

    Private Function FiguresCore(ds As String, owner As Integer) As DayFigures
        Dim byDay As Func(Of String, Double) =
            Function(sql As String) Num(sql, P("@d", ds), P("@o", owner))

        Dim f As New DayFigures()
        f.Collected = byDay(
            "SELECT CAST(COALESCE(SUM(Amount), 0) AS REAL) FROM Payments " &
            "WHERE substr(PaidAt, 1, 10) = @d AND (@o = 0 OR EmployeeId = @o)")
        f.PaymentCount = CInt(byDay(
            "SELECT COUNT(*) FROM Payments " &
            "WHERE substr(PaidAt, 1, 10) = @d AND (@o = 0 OR EmployeeId = @o)"))
        f.Refunded = byDay(
            "SELECT CAST(COALESCE(SUM(Amount), 0) AS REAL) FROM Refunds " &
            "WHERE substr(RefundedAt, 1, 10) = @d AND (@o = 0 OR EmployeeId = @o)")
        f.RefundCount = CInt(byDay(
            "SELECT COUNT(*) FROM Refunds " &
            "WHERE substr(RefundedAt, 1, 10) = @d AND (@o = 0 OR EmployeeId = @o)"))
        f.Registered = CInt(byDay(
            "SELECT COUNT(*) FROM Repairs " &
            "WHERE substr(CreatedAt, 1, 10) = @d AND (@o = 0 OR EmployeeId = @o)"))
        f.Completed = CInt(byDay(
            "SELECT COUNT(DISTINCT RepairCode) FROM ActivityLogs WHERE ActivityDate = @d " &
            "AND WasSuccessful = 1 AND ActionType = 'StatusUpdate' AND " &
            "(Description LIKE '%to Completed.' OR Description LIKE '%to Collected.') " &
            "AND (@o = 0 OR EmployeeId = @o)"))

        f.Outstanding = Math.Round(Num(
            "SELECT CAST(COALESCE(SUM(r.EstimatedCost - COALESCE((SELECT SUM(p.Amount) " &
            "FROM Payments p WHERE p.RepairCode = r.RepairCode), 0)), 0) AS REAL) " &
            "FROM Repairs r WHERE r.Status NOT IN ('Collected','Unrepairable') " &
            "AND (@o = 0 OR r.EmployeeId = @o)", P("@o", owner)), 2)
        Return f
    End Function

    Private Function BreakdownCore(ds As String) As DataTable
        Return DatabaseManager.GetTable(
            "SELECT t.EmployeeId, t.Employee, t.Payments, t.Collected, t.Refunded, " &
            "CAST(t.Collected - t.Refunded AS REAL) AS Net FROM (" &
            "SELECT e.EmployeeId, e.FullName AS Employee, " &
            "(SELECT COUNT(*) FROM Payments p WHERE p.EmployeeId = e.EmployeeId " &
            " AND substr(p.PaidAt, 1, 10) = @d) AS Payments, " &
            "CAST((SELECT COALESCE(SUM(p.Amount), 0) FROM Payments p WHERE p.EmployeeId = e.EmployeeId " &
            " AND substr(p.PaidAt, 1, 10) = @d) AS REAL) AS Collected, " &
            "CAST((SELECT COALESCE(SUM(f.Amount), 0) FROM Refunds f WHERE f.EmployeeId = e.EmployeeId " &
            " AND substr(f.RefundedAt, 1, 10) = @d) AS REAL) AS Refunded " &
            "FROM Employees e) t WHERE t.Payments > 0 OR t.Refunded > 0 " &
            "ORDER BY Net DESC, t.Employee", P("@d", ds))
    End Function

    Private Function PaymentsCore(ds As String, owner As Integer) As DataTable
        Return DatabaseManager.GetTable(
            "SELECT p.PaidAt, COALESCE(rc.ReceiptNumber, '') AS ReceiptNumber, p.RepairCode, " &
            "COALESCE(r.CustomerName, '(removed)') AS CustomerName, " &
            "CAST(p.Amount AS REAL) AS Amount, p.PaymentMethod, e.FullName AS Employee " &
            "FROM Payments p JOIN Employees e ON e.EmployeeId = p.EmployeeId " &
            "LEFT JOIN Repairs r ON r.RepairCode = p.RepairCode " &
            "LEFT JOIN Receipts rc ON rc.PaymentId = p.PaymentId " &
            "WHERE substr(p.PaidAt, 1, 10) = @d AND (@o = 0 OR p.EmployeeId = @o) " &
            "ORDER BY p.PaymentId", P("@d", ds), P("@o", owner))
    End Function

    Private Function RefundsCore(ds As String) As DataTable
        Return DatabaseManager.GetTable(
            "SELECT f.RefundedAt, f.RepairCode, CAST(f.Amount AS REAL) AS Amount, " &
            "f.RefundMethod, f.EmployeeName FROM Refunds f " &
            "WHERE substr(f.RefundedAt, 1, 10) = @d ORDER BY f.RefundId", P("@d", ds))
    End Function

    Private Function ClosingCore(ds As String) As DataTable
        Return DatabaseManager.GetTable(
            "SELECT ClosingId, BusinessDate, ClosedByName, ClosedAt, " &
            "COALESCE(ReportPath, '') AS ReportPath FROM DailyClosings WHERE BusinessDate = @d",
            P("@d", ds))
    End Function

    Private Function IsClosed(d As Date) As Boolean
        Return Num("SELECT COUNT(*) FROM DailyClosings WHERE BusinessDate = @d",
                   P("@d", DateText(d))) > 0
    End Function

    Private Function HasActivity(d As Date) As Boolean
        Dim ds As String = DateText(d)
        Return Num("SELECT (SELECT COUNT(*) FROM Payments WHERE substr(PaidAt, 1, 10) = @d) + " &
                   "(SELECT COUNT(*) FROM Refunds WHERE substr(RefundedAt, 1, 10) = @d) + " &
                   "(SELECT COUNT(*) FROM Repairs WHERE substr(CreatedAt, 1, 10) = @d)",
                   P("@d", ds)) > 0
    End Function

    ' ---------- Public queries ----------

    ''' <summary>Employees always get their own figures; administrators get the shop (ownerId 0).</summary>
    Public Function GetDayFigures(d As Date, ownerId As Integer) As DayFigures
        SessionManager.RequireLogin()
        Dim owner As Integer = If(SessionManager.IsEmployee, SessionManager.CurrentUserId, ownerId)
        Return FiguresCore(DateText(d), owner)
    End Function

    Public Function GetEmployeeBreakdown(d As Date) As DataTable
        SessionManager.RequireAdmin()
        Return BreakdownCore(DateText(d))
    End Function

    Public Function GetPaymentsForDate(d As Date, ownerId As Integer) As DataTable
        SessionManager.RequireLogin()
        Dim owner As Integer = If(SessionManager.IsEmployee, SessionManager.CurrentUserId, ownerId)
        Return PaymentsCore(DateText(d), owner)
    End Function

    Public Function GetClosing(d As Date) As DataTable
        SessionManager.RequireAdmin()
        Return ClosingCore(DateText(d))
    End Function

    ' ---------- Login lock ----------

    ''' <summary>
    ''' "" if employees may log in. Otherwise the reason. The day closed today locks employees
    ''' until 6:00 AM tomorrow; a day closed yesterday locks them until 6:00 AM today.
    ''' </summary>
    Public Function EmployeeLoginBlockReason() As String
        Try
            Dim nowT As DateTime = DateTime.Now
            If IsClosed(nowT.Date) Then
                Return "The working day has been closed. Employees can log in again from 6:00 AM tomorrow."
            End If
            If nowT.TimeOfDay < ReopenTime AndAlso IsClosed(nowT.Date.AddDays(-1)) Then
                Return "The working day has been closed. Employees can log in again from 6:00 AM."
            End If
        Catch
        End Try
        Return ""
    End Function

    ' ---------- Closing ----------

    ''' <summary>Administrator closes a day. Returns "" on success, otherwise an error message.</summary>
    Public Function CloseDay(d As Date, ByRef reportPath As String,
                             ByRef reportError As String) As String
        reportPath = ""
        reportError = ""
        SessionManager.RequireAdmin()
        If d.Date > Date.Today Then Return "A day in the future cannot be closed."
        If IsClosed(d.Date) Then Return "This day has already been closed."
        Return CloseCore(d.Date, CObj(SessionManager.CurrentUserId),
                         SessionManager.CurrentUserName, "Administrator", reportPath, reportError)
    End Function

    ''' <summary>
    ''' Closes today at 11:30 PM, and catches up yesterday if it had activity but was never
    ''' closed. Needs no login. Safe to call often. Returns True if it closed a day now.
    ''' </summary>
    Public Function AutoCloseIfDue(Optional force As Boolean = False) As Boolean
        If Not force AndAlso (DateTime.Now - lastAutoAttempt).TotalSeconds < 30 Then Return False
        lastAutoAttempt = DateTime.Now

        Dim closedNow As Boolean = False
        Try
            Dim nowT As DateTime = DateTime.Now
            Dim yesterday As Date = nowT.Date.AddDays(-1)
            If Not IsClosed(yesterday) AndAlso HasActivity(yesterday) Then
                If AutoClose(yesterday) Then closedNow = True
            End If
            If nowT.TimeOfDay >= CloseTime AndAlso Not IsClosed(nowT.Date) Then
                If AutoClose(nowT.Date) Then closedNow = True
            End If
        Catch
        End Try
        Return closedNow
    End Function

    Private Function AutoClose(d As Date) As Boolean
        Dim rp As String = ""
        Dim re As String = ""
        Return CloseCore(d, Nothing, "System (automatic)", "System", rp, re) = ""
    End Function

    Private Function CloseCore(bizDate As Date, closedById As Object, closedByName As String,
                               role As String, ByRef reportPath As String,
                               ByRef reportError As String) As String
        reportPath = ""
        reportError = ""
        Dim ds As String = DateText(bizDate)
        Dim f As DayFigures = FiguresCore(ds, 0)
        Dim emp As DataTable = BreakdownCore(ds)
        Dim stamp As String = DatabaseManager.NowText()
        Dim isAuto As Boolean = (role = "System")
        Dim closingId As Long = 0
        Dim err As String = ""

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        Dim already As Integer
                        Using cmd As SqliteCommand = conn.CreateCommand()
                            cmd.Transaction = tx
                            cmd.CommandText = "SELECT COUNT(*) FROM DailyClosings WHERE BusinessDate = @d"
                            cmd.Parameters.Add(P("@d", ds))
                            already = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using

                        If already > 0 Then
                            err = "This day has already been closed."
                            tx.Rollback()
                        Else
                            ' 1. The closing record
                            Using cmd As SqliteCommand = conn.CreateCommand()
                                cmd.Transaction = tx
                                cmd.CommandText =
                                    "INSERT INTO DailyClosings (BusinessDate, TotalRevenue, PaymentCount, " &
                                    "RepairsRegistered, RepairsCompleted, OutstandingBalance, ClosedByName, " &
                                    "ClosedByRole, ClosedById, ClosedAt) VALUES (@d, @rev, @pc, @reg, @done, " &
                                    "@out, @by, @role, @byId, @at) RETURNING ClosingId"
                                cmd.Parameters.Add(P("@d", ds))
                                cmd.Parameters.Add(P("@rev", f.NetRevenue))
                                cmd.Parameters.Add(P("@pc", f.PaymentCount))
                                cmd.Parameters.Add(P("@reg", f.Registered))
                                cmd.Parameters.Add(P("@done", f.Completed))
                                cmd.Parameters.Add(P("@out", f.Outstanding))
                                cmd.Parameters.Add(P("@by", closedByName))
                                cmd.Parameters.Add(P("@role", role))
                                cmd.Parameters.Add(P("@byId", closedById))
                                cmd.Parameters.Add(P("@at", stamp))
                                closingId = Convert.ToInt64(cmd.ExecuteScalar())
                            End Using

                            ' 2. Each employee's revenue for the day
                            For Each r As DataRow In emp.Rows
                                Run(conn, tx,
                                    "INSERT INTO EmployeeDailyRevenue (ClosingId, EmployeeId, EmployeeName, " &
                                    "Revenue, PaymentCount) VALUES (@c, @e, @n, @rev, @pc)",
                                    P("@c", closingId),
                                    P("@e", Convert.ToInt32(r("EmployeeId"))),
                                    P("@n", r("Employee").ToString()),
                                    P("@rev", Convert.ToDouble(r("Net"))),
                                    P("@pc", Convert.ToInt32(r("Payments"))))
                            Next

                            ' 3. End employee sessions still open
                            Run(conn, tx,
                                "UPDATE EmployeeSessions SET LogoutAt = @t, " &
                                "LogoutReason = 'Day closed' " &
                                "WHERE LogoutAt IS NULL AND substr(LoginAt, 1, 10) <= @d",
                                P("@t", stamp), P("@d", ds))

                            ' 4. Audit entry
                            Dim details As String =
                                If(isAuto, "Automatic close of the working day ", "Closed the working day ") &
                                ds & ". Net revenue " & UIHelpers.Money(f.NetRevenue) & " from " &
                                f.PaymentCount & " payment(s); " & f.Registered & " repair(s) registered, " &
                                f.Completed & " completed; outstanding balances " &
                                UIHelpers.Money(f.Outstanding) & "."
                            Run(conn, tx,
                                "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, " &
                                "PerformedByRole, Details, EventAt) VALUES ('DayClosed', @id, @n, @r, @d, @t)",
                                P("@id", closedById), P("@n", closedByName), P("@r", role),
                                P("@d", details), P("@t", stamp))
                            tx.Commit()
                        End If
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Return "The day could not be closed, so nothing was saved. (" & ex.Message & ")"
        End Try

        If err <> "" Then Return err

        ' The Excel report is written AFTER the commit, so a report problem can never undo the close.
        Try
            reportPath = BuildReportCore(bizDate, closedByName)
            DatabaseManager.ExecuteNonQuery(
                "UPDATE DailyClosings SET ReportPath = @p WHERE ClosingId = @id",
                P("@p", reportPath), P("@id", closingId))
        Catch ex As Exception
            reportPath = ""
            reportError = ex.Message
        End Try
        Return ""
    End Function

    ' ---------- Reopening (administrator, today only, before the automatic close) ----------

    Public Function ReopenDay(d As Date) As String
        SessionManager.RequireAdmin()
        If d.Date <> Date.Today Then Return "Only today's closing can be reopened."
        If DateTime.Now.TimeOfDay >= CloseTime Then
            Return "The day cannot be reopened after 11:30 PM because it closes automatically."
        End If

        Dim ds As String = DateText(d.Date)
        Dim stamp As String = DatabaseManager.NowText()

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        Dim found As Boolean = False
                        Dim closingId As Long = 0
                        Dim summary As String = ""

                        Using cmd As SqliteCommand = conn.CreateCommand()
                            cmd.Transaction = tx
                            cmd.CommandText =
                                "SELECT ClosingId, CAST(TotalRevenue AS REAL), PaymentCount, " &
                                "ClosedByName, ClosedAt FROM DailyClosings WHERE BusinessDate = @d"
                            cmd.Parameters.Add(P("@d", ds))
                            Using rd As SqliteDataReader = cmd.ExecuteReader()
                                If rd.Read() Then
                                    found = True
                                    closingId = rd.GetInt64(0)
                                    summary = "closed by " & rd.GetString(3) & " at " & rd.GetString(4) &
                                              ": net revenue " & UIHelpers.Money(rd.GetDouble(1)) &
                                              " from " & rd.GetInt32(2) & " payment(s)"
                                End If
                            End Using
                        End Using

                        If Not found Then
                            tx.Rollback()
                            Return "This day is not closed."
                        End If

                        Dim emps As New List(Of String)
                        Using cmd As SqliteCommand = conn.CreateCommand()
                            cmd.Transaction = tx
                            cmd.CommandText =
                                "SELECT EmployeeName, CAST(Revenue AS REAL) FROM EmployeeDailyRevenue " &
                                "WHERE ClosingId = @c"
                            cmd.Parameters.Add(P("@c", closingId))
                            Using rd As SqliteDataReader = cmd.ExecuteReader()
                                While rd.Read()
                                    emps.Add(rd.GetString(0) & " " & UIHelpers.Money(rd.GetDouble(1)))
                                End While
                            End Using
                        End Using

                        Run(conn, tx,
                            "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, " &
                            "PerformedByRole, OriginalData, Details, EventAt) " &
                            "VALUES ('DayReopened', @id, @n, 'Administrator', @orig, @d, @t)",
                            P("@id", SessionManager.CurrentUserId),
                            P("@n", SessionManager.CurrentUserName),
                            P("@orig", "Employee revenue at closing: " &
                                       If(emps.Count = 0, "(none)", String.Join("; ", emps))),
                            P("@d", "Reopened the working day " & ds & ". The closing had been " & summary & "."),
                            P("@t", stamp))

                        Run(conn, tx, "DELETE FROM EmployeeDailyRevenue WHERE ClosingId = @c", P("@c", closingId))
                        Run(conn, tx, "DELETE FROM DailyClosings WHERE ClosingId = @c", P("@c", closingId))
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Return "The day could not be reopened. (" & ex.Message & ")"
        End Try
        Return ""
    End Function

    ' ---------- Excel ----------

    ''' <summary>Adds a DataTable as a worksheet. Empty tables are skipped (except Summary).</summary>
    Public Sub AddSheet(wb As XLWorkbook, dt As DataTable, sheetName As String)
        If dt.Rows.Count = 0 AndAlso sheetName <> "Summary" Then Return
        dt.TableName = sheetName
        Dim ws = wb.Worksheets.Add(dt, sheetName)
        ws.Columns().AdjustToContents()
    End Sub

    Public Function BuildReport(d As Date) As String
        SessionManager.RequireAdmin()
        Return BuildReportCore(d.Date, SessionManager.CurrentUserName)
    End Function

    Private Function BuildReportCore(bizDate As Date, createdBy As String) As String
        Dim ds As String = DateText(bizDate)
        Dim f As DayFigures = FiguresCore(ds, 0)
        Dim closing As DataTable = ClosingCore(ds)

        Dim summary As New DataTable("Summary")
        summary.Columns.Add("Item", GetType(String))
        summary.Columns.Add("Value", GetType(String))
        summary.Rows.Add("Business", AdminToolsService.GetSetting("BusinessName"))
        summary.Rows.Add("Business date", bizDate.ToString("dddd, dd MMMM yyyy", CultureInfo.InvariantCulture))
        If closing.Rows.Count > 0 Then
            summary.Rows.Add("Status", "Closed")
            summary.Rows.Add("Closed by", closing.Rows(0)("ClosedByName").ToString())
            summary.Rows.Add("Closed at", closing.Rows(0)("ClosedAt").ToString())
        Else
            summary.Rows.Add("Status", "Not closed yet (preview)")
        End If
        summary.Rows.Add("Report created", DatabaseManager.NowText() & " by " & createdBy)
        summary.Rows.Add("", "")
        summary.Rows.Add("Net revenue (after refunds)", UIHelpers.Money(f.NetRevenue))
        summary.Rows.Add("Payments collected", UIHelpers.Money(f.Collected))
        summary.Rows.Add("Number of payments", f.PaymentCount.ToString())
        summary.Rows.Add("Refunds given", UIHelpers.Money(f.Refunded))
        summary.Rows.Add("Number of refunds", f.RefundCount.ToString())
        summary.Rows.Add("Repairs registered", f.Registered.ToString())
        summary.Rows.Add("Repairs completed", f.Completed.ToString())
        summary.Rows.Add("Outstanding balances (as of report time)", UIHelpers.Money(f.Outstanding))

        ' Build a fresh table without the ID column (removing a key column is not allowed).
        Dim empRaw As DataTable = BreakdownCore(ds)
        Dim emp As New DataTable("Employees")
        emp.Columns.Add("Employee", GetType(String))
        emp.Columns.Add("Payments", GetType(Integer))
        emp.Columns.Add("Collected", GetType(Double))
        emp.Columns.Add("Refunded", GetType(Double))
        emp.Columns.Add("Net", GetType(Double))
        For Each r As DataRow In empRaw.Rows
            emp.Rows.Add(r("Employee").ToString(),
                         Convert.ToInt32(r("Payments")),
                         Convert.ToDouble(r("Collected")),
                         Convert.ToDouble(r("Refunded")),
                         Convert.ToDouble(r("Net")))
        Next

        Dim pays As DataTable = PaymentsCore(ds, 0)
        Dim refunds As DataTable = RefundsCore(ds)

        Dim filePath As String = Path.Combine(ReportsFolder(), "EndOfDay_" & ds & ".xlsx")
        Using wb As New XLWorkbook()
            AddSheet(wb, summary, "Summary")
            AddSheet(wb, emp, "Employees")
            AddSheet(wb, pays, "Payments")
            AddSheet(wb, refunds, "Refunds")
            wb.SaveAs(filePath)
        End Using
        Return filePath
    End Function

End Module