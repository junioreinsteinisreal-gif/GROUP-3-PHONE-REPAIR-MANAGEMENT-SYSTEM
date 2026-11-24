Imports System.Globalization
Imports Microsoft.Data.Sqlite

''' <summary>Terminating repairs that are beyond repair (with refund), and listing them.</summary>
Public Module TerminationService

    Public Const MinReasonLength As Integer = 10
    Public Const MaxReasonLength As Integer = 500

    Private schemaReady As Boolean = False

    Private Sub EnsureSchema()
        If schemaReady Then Return
        DatabaseManager.ExecuteNonQuery("
CREATE TABLE IF NOT EXISTS RepairTerminations (
    TerminationId       INTEGER PRIMARY KEY AUTOINCREMENT,
    RepairCode          TEXT NOT NULL UNIQUE,
    CustomerName        TEXT NOT NULL,
    CustomerPhone       TEXT NOT NULL,
    Brand               TEXT NOT NULL,
    Model               TEXT NOT NULL,
    ProblemDescription  TEXT NOT NULL,
    EstimatedCost       REAL NOT NULL,
    AmountPaid          REAL NOT NULL DEFAULT 0,
    PreviousStatus      TEXT NOT NULL,
    Reason              TEXT NOT NULL,
    TerminatedById      INTEGER NOT NULL,
    TerminatedByName    TEXT NOT NULL,
    TerminatedAt        TEXT NOT NULL,
    FOREIGN KEY (TerminatedById) REFERENCES Employees(EmployeeId)
);
CREATE INDEX IF NOT EXISTS idx_terminations_date ON RepairTerminations(TerminatedAt);")
        schemaReady = True
    End Sub

    ' ---------- Terminate (employees) ----------

    ''' <summary>Returns "" on success, otherwise an error message.</summary>
    Public Function TerminateRepair(repairCode As String, reason As String,
                                    refundAmount As Decimal, refundMethod As String) As String
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can terminate repairs.")
        End If

        Dim code As String = If(repairCode, "").Trim()
        Dim why As String = If(reason, "").Trim()
        If code = "" Then Return "Select a repair first."
        If why.Length < MinReasonLength Then
            Return "Type a reason of at least " & MinReasonLength & " characters."
        End If
        If why.Length > MaxReasonLength Then
            Return "The reason is too long (" & MaxReasonLength & " characters maximum)."
        End If
        If refundAmount < 0D Then Return "The refund cannot be negative."
        If refundAmount <> Math.Round(refundAmount, 2) Then
            Return "The refund can have at most 2 decimal places."
        End If
        If refundAmount > 0D AndAlso
           Array.IndexOf(RepairService.PaymentMethods, refundMethod) < 0 Then
            Return "Choose a valid refund method."
        End If

        EnsureSchema()

        Dim err As String = ""
        Dim done As Boolean = False
        Dim stamp As String = DatabaseManager.NowText()

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        done = ApplyTermination(conn, tx, code, why, refundAmount,
                                                refundMethod, stamp, err)
                        If done Then tx.Commit() Else tx.Rollback()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            ActivityLogger.Log("Terminate",
                               "Failed to terminate repair " & code & ": " & ex.Message,
                               code, False)
            Return "The repair could not be terminated, so nothing was changed. (" & ex.Message & ")"
        End Try

        If Not done Then
            ActivityLogger.Log("Terminate",
                               "Termination of repair " & code & " was rejected: " & err,
                               code, False)
            Return err
        End If

        ' Logging happens AFTER the commit (SQLite allows one writer at a time).
        ActivityLogger.Log("Terminate",
                           "Terminated repair " & code & " as beyond repair. Reason: " & why, code)
        If refundAmount > 0D Then
            ActivityLogger.Log("Refund",
                               "Refunded " & UIHelpers.Money(CDbl(refundAmount)) & " (" &
                               refundMethod & ") to the customer on " & code & ".", code)
        End If
        Return ""
    End Function

    Private Function ApplyTermination(conn As SqliteConnection, tx As SqliteTransaction,
                                      code As String, why As String, refund As Decimal,
                                      method As String, stamp As String,
                                      ByRef err As String) As Boolean
        Dim customer As String
        Dim phone As String
        Dim brand As String
        Dim model As String
        Dim problem As String
        Dim status As String
        Dim cost As Double
        Dim paid As Double
        Dim refundedBefore As Double
        Dim ownerId As Integer

        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText =
                "SELECT r.CustomerName, r.CustomerPhone, r.Brand, r.Model, r.ProblemDescription, " &
                "r.Status, CAST(r.EstimatedCost AS REAL), " &
                "CAST(COALESCE((SELECT SUM(p.Amount) FROM Payments p " &
                "WHERE p.RepairCode = r.RepairCode), 0) AS REAL), " &
                "CAST(COALESCE((SELECT SUM(f.Amount) FROM Refunds f " &
                "WHERE f.RepairCode = r.RepairCode), 0) AS REAL), r.EmployeeId " &
                "FROM Repairs r WHERE r.RepairCode = @c"
            cmd.Parameters.Add(DatabaseManager.P("@c", code))
            Using rd As SqliteDataReader = cmd.ExecuteReader()
                If Not rd.Read() Then
                    err = "That repair no longer exists."
                    Return False
                End If
                customer = rd.GetString(0)
                phone = rd.GetString(1)
                brand = rd.GetString(2)
                model = rd.GetString(3)
                problem = rd.GetString(4)
                status = rd.GetString(5)
                cost = rd.GetDouble(6)
                paid = rd.GetDouble(7)
                refundedBefore = rd.GetDouble(8)
                ownerId = rd.GetInt32(9)
            End Using
        End Using

        If ownerId <> SessionManager.CurrentUserId Then
            err = "You can only terminate repairs that you registered yourself."
            Return False
        End If
        If status = "Unrepairable" Then
            err = "This repair has already been terminated."
            Return False
        End If
        If status = "Collected" Then
            err = "This repair has already been collected and cannot be terminated."
            Return False
        End If
        If status <> "Received" AndAlso status <> "In Progress" Then
            err = "Only repairs that are Received or In Progress can be terminated."
            Return False
        End If

        Dim refundable As Decimal = Convert.ToDecimal(Math.Round(paid - refundedBefore, 2))
        If refund > refundable Then
            err = "The refund (" & UIHelpers.Money(CDbl(refund)) &
                  ") is more than the customer has paid (" &
                  UIHelpers.Money(CDbl(refundable)) & ")."
            Return False
        End If

        ' ----- 1. Close the repair -----
        Dim rows As Integer = Run(conn, tx,
            "UPDATE Repairs SET Status = 'Unrepairable', UpdatedAt = @u " &
            "WHERE RepairCode = @c AND Status IN ('Received','In Progress')",
            DatabaseManager.P("@u", stamp),
            DatabaseManager.P("@c", code))
        If rows <> 1 Then
            err = "The repair changed while you were working. Search for it again."
            Return False
        End If

        ' ----- 2. Record the refund -----
        If refund > 0D Then
            Run(conn, tx,
                "INSERT INTO Refunds (RepairCode, Amount, RefundMethod, RefundedAt, " &
                "EmployeeId, EmployeeName, Reason) " &
                "VALUES (@code, @amt, @method, @at, @empId, @empName, 'Repair terminated')",
                DatabaseManager.P("@code", code),
                DatabaseManager.P("@amt", CDbl(refund)),
                DatabaseManager.P("@method", method),
                DatabaseManager.P("@at", stamp),
                DatabaseManager.P("@empId", SessionManager.CurrentUserId),
                DatabaseManager.P("@empName", SessionManager.CurrentUserName))
        End If

        ' ----- 3. Save the termination record -----
        Run(conn, tx,
            "INSERT INTO RepairTerminations (RepairCode, CustomerName, CustomerPhone, Brand, " &
            "Model, ProblemDescription, EstimatedCost, AmountPaid, PreviousStatus, Reason, " &
            "TerminatedById, TerminatedByName, TerminatedAt) " &
            "VALUES (@code, @name, @phone, @brand, @model, @problem, @cost, @paid, @prev, " &
            "@reason, @empId, @empName, @at)",
            DatabaseManager.P("@code", code),
            DatabaseManager.P("@name", customer),
            DatabaseManager.P("@phone", phone),
            DatabaseManager.P("@brand", brand),
            DatabaseManager.P("@model", model),
            DatabaseManager.P("@problem", problem),
            DatabaseManager.P("@cost", cost),
            DatabaseManager.P("@paid", paid),
            DatabaseManager.P("@prev", status),
            DatabaseManager.P("@reason", why),
            DatabaseManager.P("@empId", SessionManager.CurrentUserId),
            DatabaseManager.P("@empName", SessionManager.CurrentUserName),
            DatabaseManager.P("@at", stamp))

        ' ----- 4. Administrator audit entry (same transaction) -----
        Dim snapshot As String =
            "Repair " & code & " | Customer: " & customer & " (" & phone & ")" &
            " | Device: " & brand & " " & model & " | Problem: " & problem &
            " | Estimated: " & UIHelpers.Money(cost) & " | Paid: " & UIHelpers.Money(paid) &
            " | Refunded: " & UIHelpers.Money(CDbl(refund)) & " | Status before: " & status

        Run(conn, tx,
            "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, " &
            "PerformedByRole, RepairCode, OriginalData, Details, EventAt) " &
            "VALUES ('RepairTerminated', @id, @name, 'Employee', @code, @orig, @details, @at)",
            DatabaseManager.P("@id", SessionManager.CurrentUserId),
            DatabaseManager.P("@name", SessionManager.CurrentUserName),
            DatabaseManager.P("@code", code),
            DatabaseManager.P("@orig", snapshot),
            DatabaseManager.P("@details",
                SessionManager.CurrentUserName & " terminated repair " & code &
                " as beyond repair. Refund given: " & UIHelpers.Money(CDbl(refund)) &
                ". Reason: " & why),
            DatabaseManager.P("@at", stamp))

        Return True
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

    ' ---------- Summary for the update screen ----------

    Public Function GetTerminationSummary(repairCode As String) As String
        SessionManager.RequireLogin()
        EnsureSchema()

        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT TerminatedByName, TerminatedAt, Reason FROM RepairTerminations " &
            "WHERE RepairCode = @c",
            DatabaseManager.P("@c", repairCode))
        If dt.Rows.Count = 0 Then Return "This repair was terminated. No further changes are allowed."

        Dim r As DataRow = dt.Rows(0)
        Dim whenText As String = r("TerminatedAt").ToString()
        Dim parsed As DateTime
        If DateTime.TryParseExact(whenText, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture,
                                  DateTimeStyles.None, parsed) Then
            whenText = parsed.ToString("dd MMM yyyy HH:mm")
        End If

        Dim refunded As Double = Convert.ToDouble(DatabaseManager.ExecuteScalar(
            "SELECT COALESCE(SUM(Amount), 0) FROM Refunds WHERE RepairCode = @c",
            DatabaseManager.P("@c", repairCode)))

        Dim text As String = "Terminated by " & r("TerminatedByName").ToString() & " on " &
                             whenText & ". Reason: " & r("Reason").ToString()
        If refunded > 0 Then text &= "  Refund given: " & UIHelpers.Money(refunded) & "."
        Return text
    End Function

    ' ---------- List (administrators) ----------

    Private Function EscapeLike(text As String) As String
        Return text.Replace("\", "\\").Replace("%", "\%").Replace("_", "\_")
    End Function

    Public Function GetTerminations(searchText As String) As DataTable
        SessionManager.RequireAdmin()
        EnsureSchema()

        Dim q As String = If(searchText, "").Trim()
        Dim likeText As String = "%" & EscapeLike(q) & "%"

        Return DatabaseManager.GetTable(
            "SELECT t.TerminatedAt, t.RepairCode, t.CustomerName, t.CustomerPhone, " &
            "t.Brand || ' ' || t.Model AS Device, CAST(t.EstimatedCost AS REAL) AS EstimatedCost, " &
            "CAST(t.AmountPaid AS REAL) AS AmountPaid, " &
            "CAST(COALESCE((SELECT SUM(f.Amount) FROM Refunds f " &
            "WHERE f.RepairCode = t.RepairCode), 0) AS REAL) AS Refunded, " &
            "t.PreviousStatus, t.TerminatedByName, t.Reason FROM RepairTerminations t " &
            "WHERE (@q = '' " &
            "OR t.RepairCode LIKE @like ESCAPE '\' " &
            "OR t.CustomerName LIKE @like ESCAPE '\' " &
            "OR t.CustomerPhone LIKE @like ESCAPE '\' " &
            "OR t.TerminatedByName LIKE @like ESCAPE '\' " &
            "OR t.Reason LIKE @like ESCAPE '\') " &
            "ORDER BY t.TerminationId DESC",
            DatabaseManager.P("@q", q),
            DatabaseManager.P("@like", likeText))
    End Function

End Module