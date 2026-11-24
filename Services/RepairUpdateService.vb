Imports System.Globalization
Imports Microsoft.Data.Sqlite

''' <summary>Looks up a repair and applies status, payment and date updates.</summary>
Public Module RepairUpdateService

    Private Const SelectRepairSql As String =
        "SELECT r.RepairCode, r.CustomerName, r.CustomerPhone, r.Brand, r.Model, " &
        "r.ProblemDescription, r.Status, CAST(r.EstimatedCost AS REAL), " &
        "CAST(COALESCE((SELECT SUM(p.Amount) FROM Payments p " &
        "WHERE p.RepairCode = r.RepairCode), 0) AS REAL), " &
        "r.DateReceived, r.ExpectedCollectionDate, r.CollectedAt, e.FullName, r.EmployeeId " &
        "FROM Repairs r JOIN Employees e ON e.EmployeeId = r.EmployeeId " &
        "WHERE r.RepairCode = @c"

    ' ---------- Lookup ----------

    ''' <summary>Returns the repair, or Nothing if the code does not exist (or is not yours).</summary>
    Public Function FindRepair(repairCode As String) As RepairInfo
        SessionManager.RequireLogin()
        Using conn As SqliteConnection = DatabaseManager.GetConnection()
            Dim info As RepairInfo = ReadRepair(conn, Nothing, repairCode)
            ' Employees can only open repairs they registered.
            If info IsNot Nothing AndAlso SessionManager.IsEmployee AndAlso
               info.OwnerId <> SessionManager.CurrentUserId Then Return Nothing
            Return info
        End Using
    End Function

    Private Function ReadRepair(conn As SqliteConnection, tx As SqliteTransaction,
                                repairCode As String) As RepairInfo
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText = SelectRepairSql
            cmd.Parameters.Add(DatabaseManager.P("@c", repairCode))
            Using rd As SqliteDataReader = cmd.ExecuteReader()
                If Not rd.Read() Then Return Nothing

                Dim info As New RepairInfo()
                info.RepairCode = rd.GetString(0)
                info.CustomerName = rd.GetString(1)
                info.CustomerPhone = rd.GetString(2)
                info.Brand = rd.GetString(3)
                info.Model = rd.GetString(4)
                info.Problem = rd.GetString(5)
                info.Status = rd.GetString(6)
                info.EstimatedCost = Convert.ToDecimal(Math.Round(rd.GetDouble(7), 2))
                info.TotalPaid = Convert.ToDecimal(Math.Round(rd.GetDouble(8), 2))
                info.DateReceived = Date.ParseExact(rd.GetString(9), "yyyy-MM-dd",
                                                    CultureInfo.InvariantCulture)
                info.ExpectedCollection = Date.ParseExact(rd.GetString(10), "yyyy-MM-dd",
                                                          CultureInfo.InvariantCulture)
                info.CollectedAt = If(rd.IsDBNull(11), "", rd.GetString(11))
                info.EmployeeName = rd.GetString(12)
                info.OwnerId = rd.GetInt32(13)
                Return info
            End Using
        End Using
    End Function

    ' ---------- Update ----------

    ''' <summary>
    ''' Applies a status change, an optional payment and a collection date in ONE
    ''' transaction. Returns "" on success, otherwise an error message.
    ''' </summary>
    Public Function UpdateRepair(repairCode As String, newStatus As String,
                                 paymentAmount As Decimal, paymentMethod As String,
                                 newExpectedDate As Date) As String

        ' Business rule: only employees update repairs.
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can update repairs.")
        End If

        Dim code As String = If(repairCode, "").Trim()
        If code = "" Then Return "Select a repair first."
        If Array.IndexOf(RepairService.RepairStatuses, newStatus) < 0 Then
            Return "Choose a valid status."
        End If
        If newStatus = "Unrepairable" Then
            Return "To mark a phone as beyond repair, use the Terminate Repair button so a reason is recorded."
        End If
        If paymentAmount < 0 Then Return "The payment cannot be negative."
        If paymentAmount > 1000000D Then Return "The payment is too large."
        If paymentAmount <> Math.Round(paymentAmount, 2) Then
            Return "The payment can have at most 2 decimal places."
        End If
        If paymentAmount > 0 AndAlso
           Array.IndexOf(RepairService.PaymentMethods, paymentMethod) < 0 Then
            Return "Choose a valid payment method."
        End If

        Dim err As String = ""
        Dim oldStatus As String = ""
        Dim oldExpected As Date
        Dim applied As Boolean = False
        Dim stamp As String = DatabaseManager.NowText()

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        applied = ApplyUpdate(conn, tx, code, newStatus, paymentAmount,
                                              paymentMethod, newExpectedDate.Date, stamp,
                                              err, oldStatus, oldExpected)
                        If applied Then tx.Commit() Else tx.Rollback()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            ActivityLogger.Log("UpdateRepair",
                               "Failed to update repair " & code & ": " & ex.Message,
                               code, False)
            Return "The update could not be saved, so nothing was changed. (" & ex.Message & ")"
        End Try

        ' Rejected by a business rule: log it so administrators can investigate.
        If Not applied Then
            ActivityLogger.Log("UpdateRepair",
                               "Update of repair " & code & " was rejected: " & err,
                               code, False)
            Return err
        End If

        ' Logging and customer messages happen AFTER the commit (SQLite allows one writer at a time).
        If newStatus <> oldStatus Then
            ActivityLogger.Log("StatusUpdate",
                               "Changed repair " & code & " from " & oldStatus &
                               " to " & newStatus & ".", code)
            NotificationService.OnStatusChanged(code, newStatus)
        End If
        If paymentAmount > 0 Then
            ActivityLogger.Log("Payment",
                               "Received " & UIHelpers.Money(CDbl(paymentAmount)) &
                               " payment (" & paymentMethod & ") on " & code & ".", code)
        End If
        If newExpectedDate.Date <> oldExpected Then
            ActivityLogger.Log("UpdateRepair",
                               "Changed expected collection date of " & code & " to " &
                               newExpectedDate.ToString("dd MMM yyyy") & ".", code)
        End If
        Return ""
    End Function

    ''' <summary>
    ''' Re-reads the repair inside the transaction, checks the rules, then writes.
    ''' Returns False (with err set) if a rule is broken. Nothing is written then.
    ''' </summary>
    Private Function ApplyUpdate(conn As SqliteConnection, tx As SqliteTransaction,
                                 code As String, newStatus As String, payment As Decimal,
                                 method As String, newExpected As Date, stamp As String,
                                 ByRef err As String, ByRef oldStatus As String,
                                 ByRef oldExpected As Date) As Boolean

        Dim repair As RepairInfo = ReadRepair(conn, tx, code)
        If repair Is Nothing Then
            err = "That repair no longer exists."
            Return False
        End If
        If repair.OwnerId <> SessionManager.CurrentUserId Then
            err = "You can only update repairs that you registered yourself."
            Return False
        End If
        oldStatus = repair.Status
        oldExpected = repair.ExpectedCollection

        If repair.Status = "Collected" Then
            err = "This repair has already been collected and can no longer be changed."
            Return False
        End If

        If repair.Status = "Unrepairable" Then
            err = "This repair was terminated and can no longer be changed."
            Return False
        End If

        If payment > repair.Balance Then
            err = "The payment (" & UIHelpers.Money(CDbl(payment)) &
                  ") is more than the outstanding balance (" &
                  UIHelpers.Money(CDbl(repair.Balance)) & ")."
            Return False
        End If

        If newStatus = "Collected" AndAlso (repair.Balance - payment) > 0.005D Then
            err = "A repair cannot be marked Collected until the balance is fully paid. " &
                  "Balance after this payment: " &
                  UIHelpers.Money(CDbl(repair.Balance - payment)) & "."
            Return False
        End If

        If newExpected < repair.DateReceived Then
            err = "The collection date cannot be before the date received (" &
                  repair.DateReceived.ToString("dd MMM yyyy") & ")."
            Return False
        End If

        If newStatus = repair.Status AndAlso payment = 0D AndAlso
           newExpected = repair.ExpectedCollection Then
            err = "Nothing to update. Change the status, add a payment or change the collection date."
            Return False
        End If

        ' ----- Write the repair -----
        Dim collectedValue As Object = If(newStatus = "Collected", CObj(stamp), Nothing)
        Dim rows As Integer = Run(conn, tx,
            "UPDATE Repairs SET Status = @s, ExpectedCollectionDate = @d, " &
            "CollectedAt = @c, UpdatedAt = @u WHERE RepairCode = @code",
            DatabaseManager.P("@s", newStatus),
            DatabaseManager.P("@d", newExpected.ToString("yyyy-MM-dd")),
            DatabaseManager.P("@c", collectedValue),
            DatabaseManager.P("@u", stamp),
            DatabaseManager.P("@code", code))
        If rows <> 1 Then Throw New InvalidOperationException("The repair could not be updated.")

        ' ----- Save the payment as its own transaction row and create its receipt -----
        If payment > 0 Then
            Run(conn, tx,
                "INSERT INTO Payments (RepairCode, Amount, PaymentMethod, PaidAt, " &
                "EmployeeId, Note) VALUES (@code, @amt, @method, @paid, @emp, " &
                "'Payment on update')",
                DatabaseManager.P("@code", code),
                DatabaseManager.P("@amt", CDbl(payment)),
                DatabaseManager.P("@method", method),
                DatabaseManager.P("@paid", stamp),
                DatabaseManager.P("@emp", SessionManager.CurrentUserId))

            Dim paymentId As Long
            Using cmd As SqliteCommand = conn.CreateCommand()
                cmd.Transaction = tx
                cmd.CommandText = "SELECT last_insert_rowid()"
                paymentId = Convert.ToInt64(cmd.ExecuteScalar())
            End Using

            ReceiptService.CreateReceipt(conn, tx, paymentId, stamp)
        End If

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

End Module