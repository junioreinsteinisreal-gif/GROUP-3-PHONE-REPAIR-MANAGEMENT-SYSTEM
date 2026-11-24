Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Data.Sqlite

''' <summary>Business logic for repairs: registering, listing and details.</summary>
Public Module RepairService

    Public ReadOnly PaymentMethods As String() =
        {"Cash", "Mobile Money", "Card", "Bank Transfer"}

    Public ReadOnly RepairStatuses As String() =
        {"Received", "In Progress", "Completed", "Collected", "Unrepairable"}

    ' ---------- Register ----------

    ''' <summary>
    ''' Registers a repair (and its initial payment) in one transaction.
    ''' Returns "" on success, otherwise an error message. repairCode is set on success.
    ''' </summary>
    Public Function RegisterRepair(customerName As String, customerPhone As String,
                                   brand As String, model As String, problem As String,
                                   estimatedCost As Decimal, initialPayment As Decimal,
                                   paymentMethod As String, dateReceived As Date,
                                   expectedCollection As Date,
                                   ByRef repairCode As String) As String
        repairCode = ""

        ' Business rule: only employees register repairs.
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can register repairs.")
        End If

        Dim name As String = If(customerName, "").Trim()
        Dim phone As String = If(customerPhone, "").Trim()
        Dim brandT As String = If(brand, "").Trim()
        Dim modelT As String = If(model, "").Trim()
        Dim problemT As String = If(problem, "").Trim()

        Dim err As String = Validate(name, phone, brandT, modelT, problemT,
                                     estimatedCost, initialPayment, paymentMethod,
                                     dateReceived, expectedCollection)
        If err <> "" Then Return err

        Dim newCode As String = ""
        Dim stamp As String = DatabaseManager.NowText()

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        ' 1. Take the next repair number
                        Dim counter As Long
                        Using cmd As SqliteCommand = conn.CreateCommand()
                            cmd.Transaction = tx
                            cmd.CommandText =
                                "UPDATE Counters SET LastValue = LastValue + 1 " &
                                "WHERE Name = 'RepairCode' RETURNING LastValue"
                            counter = Convert.ToInt64(cmd.ExecuteScalar())
                        End Using
                        newCode = "PR-" & counter.ToString("D3")

                        ' 2. Save the repair
                        Run(conn, tx,
                            "INSERT INTO Repairs (RepairCode, CustomerName, CustomerPhone, Brand, " &
                            "Model, ProblemDescription, EstimatedCost, Status, DateReceived, " &
                            "ExpectedCollectionDate, EmployeeId, CreatedAt) " &
                            "VALUES (@code, @name, @phone, @brand, @model, @problem, @cost, " &
                            "'Received', @received, @expected, @emp, @created)",
                            DatabaseManager.P("@code", newCode),
                            DatabaseManager.P("@name", name),
                            DatabaseManager.P("@phone", phone),
                            DatabaseManager.P("@brand", brandT),
                            DatabaseManager.P("@model", modelT),
                            DatabaseManager.P("@problem", problemT),
                            DatabaseManager.P("@cost", CDbl(estimatedCost)),
                            DatabaseManager.P("@received", dateReceived.ToString("yyyy-MM-dd")),
                            DatabaseManager.P("@expected", expectedCollection.ToString("yyyy-MM-dd")),
                            DatabaseManager.P("@emp", SessionManager.CurrentUserId),
                            DatabaseManager.P("@created", stamp))

                        ' 3. Save the initial payment as its own transaction row
                        ' 3. Save the initial payment and automatically create its receipt.
                        If initialPayment > 0 Then

                            ' Save payment.
                            Run(conn, tx,
        "INSERT INTO Payments (RepairCode, Amount, PaymentMethod, PaidAt, " &
        "EmployeeId, Note) VALUES (@code, @amt, @method, @paid, @emp, " &
        "'Initial payment')",
        DatabaseManager.P("@code", newCode),
        DatabaseManager.P("@amt", CDbl(initialPayment)),
        DatabaseManager.P("@method", paymentMethod),
        DatabaseManager.P("@paid", stamp),
        DatabaseManager.P("@emp", SessionManager.CurrentUserId))

                            ' Retrieve the ID of the payment just inserted.
                            Dim paymentId As Long

                            Using cmd As SqliteCommand = conn.CreateCommand()

                                cmd.Transaction = tx
                                cmd.CommandText = "SELECT last_insert_rowid()"

                                paymentId = Convert.ToInt64(cmd.ExecuteScalar())

                            End Using

                            ' Automatically create the receipt.
                            ReceiptService.CreateReceipt(
        conn, tx, paymentId, stamp)

                        End If
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            ' Failures are logged too, so administrators can investigate.
            ActivityLogger.Log("RegisterRepair",
                               "Failed to register a repair: " & ex.Message, Nothing, False)
            Return "The repair could not be saved, so nothing was recorded. (" & ex.Message & ")"
        End Try

        ' Logging happens AFTER the commit (SQLite allows one writer at a time).
        repairCode = newCode
        ActivityLogger.Log("RegisterRepair",
                           "Registered repair " & newCode & " for " & name &
                           " (" & brandT & " " & modelT & ").", newCode)
        If initialPayment > 0 Then
            ActivityLogger.Log("Payment",
                               "Received " & UIHelpers.Money(CDbl(initialPayment)) &
                               " initial payment (" & paymentMethod & ") on " & newCode & ".",
                               newCode)
        End If
        Return ""
    End Function

    Private Function Validate(name As String, phone As String, brand As String,
                              model As String, problem As String,
                              cost As Decimal, paid As Decimal, method As String,
                              received As Date, expected As Date) As String
        If name = "" Then Return "Customer name is required."
        If name.Length > 80 Then Return "Customer name is too long (80 characters maximum)."

        If phone = "" Then Return "Customer phone number is required."
        Dim digitCount As Integer = Regex.Matches(phone, "\d").Count
        If Not Regex.IsMatch(phone, "^\+?[0-9 \-]+$") OrElse digitCount < 7 OrElse digitCount > 15 Then
            Return "Enter a valid phone number (7 to 15 digits)."
        End If

        If brand = "" Then Return "Phone brand is required."
        If brand.Length > 40 Then Return "Phone brand is too long (40 characters maximum)."
        If model = "" Then Return "Phone model is required."
        If model.Length > 60 Then Return "Phone model is too long (60 characters maximum)."
        If problem = "" Then Return "Problem description is required."
        If problem.Length > 500 Then Return "Problem description is too long (500 characters maximum)."

        If cost < 0 OrElse cost > 1000000D Then Return "Estimated cost must be between 0 and 1,000,000."
        If cost <> Math.Round(cost, 2) Then Return "Estimated cost can have at most 2 decimal places."
        If paid < 0 Then Return "Initial payment cannot be negative."
        If paid <> Math.Round(paid, 2) Then Return "Initial payment can have at most 2 decimal places."
        If paid > cost Then Return "Initial payment cannot be more than the estimated cost."

        If Array.IndexOf(PaymentMethods, method) < 0 Then Return "Choose a valid payment method."

        If received.Date > Date.Today Then Return "Date received cannot be in the future."
        If expected.Date < received.Date Then
            Return "Expected collection date cannot be before the date received."
        End If
        Return ""
    End Function

    ''' <summary>Runs a command inside an open transaction.</summary>
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

    ' ---------- List ----------

    Private Function EscapeLike(text As String) As String
        Return text.Replace("\", "\\").Replace("%", "\%").Replace("_", "\_")
    End Function

    ''' <summary>Repairs with amount paid (from Payments) and balance.</summary>
    Public Function GetRepairs(searchText As String, status As String) As DataTable
        SessionManager.RequireLogin()

        Dim q As String = If(searchText, "").Trim()
        Dim likeText As String = "%" & EscapeLike(q) & "%"
        Dim st As String = If(String.IsNullOrEmpty(status), "All", status)

        Dim sql As String = "
SELECT t.RepairCode, t.CustomerName, t.CustomerPhone, t.Brand, t.Model, t.Status,
       t.EstimatedCost, t.Paid,
      CASE WHEN t.Status = 'Unrepairable' THEN 0.0 ELSE CAST(t.EstimatedCost - t.Paid AS REAL) END AS Balance,
       t.DateReceived, t.ExpectedCollectionDate, t.Employee
FROM (
    SELECT r.RepairId, r.RepairCode, r.CustomerName, r.CustomerPhone, r.Brand, r.Model,
           r.Status, CAST(r.EstimatedCost AS REAL) AS EstimatedCost,
           CAST(COALESCE((SELECT SUM(p.Amount) FROM Payments p
                          WHERE p.RepairCode = r.RepairCode), 0) AS REAL) AS Paid,
           r.DateReceived, r.ExpectedCollectionDate, e.FullName AS Employee
    FROM Repairs r
    JOIN Employees e ON e.EmployeeId = r.EmployeeId
    WHERE (@q = ''
           OR r.RepairCode    LIKE @like ESCAPE '\'
           OR r.CustomerName  LIKE @like ESCAPE '\'
           OR r.CustomerPhone LIKE @like ESCAPE '\'
           OR r.Brand         LIKE @like ESCAPE '\'
           OR r.Model         LIKE @like ESCAPE '\')
     AND (@status = 'All' OR r.Status = @status)
      AND (@owner = 0 OR r.EmployeeId = @owner)
) t
ORDER BY t.RepairId DESC"

        Return DatabaseManager.GetTable(sql,
            DatabaseManager.P("@q", q),
            DatabaseManager.P("@like", likeText),
            DatabaseManager.P("@status", st),
            DatabaseManager.P("@owner", If(SessionManager.IsEmployee, SessionManager.CurrentUserId, 0)))
    End Function

    ' ---------- Details ----------

    ''' <summary>Readable summary of one repair with its payment history ("" if not found).</summary>
    Public Function GetRepairDetails(repairCode As String) As String
        SessionManager.RequireLogin()

        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT r.*, e.FullName AS EmployeeName FROM Repairs r " &
            "JOIN Employees e ON e.EmployeeId = r.EmployeeId WHERE r.RepairCode = @c",
            DatabaseManager.P("@c", repairCode))
        If dt.Rows.Count = 0 Then Return ""
        If SessionManager.IsEmployee AndAlso Convert.ToInt32(dt.Rows(0)("EmployeeId")) <> SessionManager.CurrentUserId Then Return ""
        Dim r As DataRow = dt.Rows(0)

        Dim pay As DataTable = DatabaseManager.GetTable(
            "SELECT Amount, PaymentMethod, PaidAt FROM Payments " &
            "WHERE RepairCode = @c ORDER BY PaymentId",
            DatabaseManager.P("@c", repairCode))

        Dim cost As Double = Convert.ToDouble(r("EstimatedCost"))
        Dim total As Double = 0
        For Each p As DataRow In pay.Rows
            total += Convert.ToDouble(p("Amount"))
        Next

        Dim sb As New StringBuilder()
        sb.AppendLine("Repair code: " & r("RepairCode").ToString())
        sb.AppendLine("Status: " & r("Status").ToString())
        sb.AppendLine()
        sb.AppendLine("Customer: " & r("CustomerName").ToString())
        sb.AppendLine("Phone: " & r("CustomerPhone").ToString())
        sb.AppendLine("Device: " & r("Brand").ToString() & " " & r("Model").ToString())
        sb.AppendLine("Problem: " & r("ProblemDescription").ToString())
        sb.AppendLine()
        sb.AppendLine("Date received: " & r("DateReceived").ToString())
        sb.AppendLine("Expected collection: " & r("ExpectedCollectionDate").ToString())
        sb.AppendLine("Registered by: " & r("EmployeeName").ToString())
        sb.AppendLine()
        sb.AppendLine("Estimated cost: " & UIHelpers.Money(cost))
        sb.AppendLine("Total paid: " & UIHelpers.Money(total))
        If r("Status").ToString() = "Unrepairable" Then
            sb.AppendLine("Balance: " & UIHelpers.Money(0) & "  (terminated, nothing more is owed)")
        Else
            sb.AppendLine("Balance: " & UIHelpers.Money(cost - total))
        End If
        sb.AppendLine()
        sb.AppendLine("Payments:")
        If pay.Rows.Count = 0 Then
            sb.AppendLine("  (none yet)")
        Else
            For Each p As DataRow In pay.Rows
                sb.AppendLine("  " & p("PaidAt").ToString() & "  -  " &
                              UIHelpers.Money(Convert.ToDouble(p("Amount"))) &
                              "  (" & p("PaymentMethod").ToString() & ")")
            Next
        End If
        Return sb.ToString()
    End Function

End Module