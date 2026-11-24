Imports System.Globalization
Imports Microsoft.Data.Sqlite

''' <summary>Everything printed on one receipt.</summary>
Public Class ReceiptData
    Public Property ReceiptNumber As String = ""
    Public Property PaymentId As Long
    Public Property PaidAt As String = ""
    Public Property RepairCode As String = ""
    Public Property CustomerName As String = ""
    Public Property CustomerPhone As String = ""
    Public Property Brand As String = ""
    Public Property Model As String = ""
    Public Property EstimatedCost As Decimal
    Public Property PreviousPaid As Decimal
    Public Property Amount As Decimal
    Public Property Method As String = ""
    Public Property EmployeeName As String = ""
    Public Property PrintCount As Integer
    Public Property BusinessName As String = ""
    Public Property BusinessPhone As String = ""
    Public Property BusinessAddress As String = ""

    Public ReadOnly Property TotalPaid As Decimal
        Get
            Return PreviousPaid + Amount
        End Get
    End Property

    Public ReadOnly Property Balance As Decimal
        Get
            Return EstimatedCost - TotalPaid
        End Get
    End Property
End Class

''' <summary>Payment list, receipt numbers (one per payment) and print counting.</summary>
Public Module ReceiptService

    Private Function EscapeLike(text As String) As String
        Return text.Replace("\", "\\").Replace("%", "\%").Replace("_", "\_")
    End Function

    Private Function Setting(key As String) As String
        Return Convert.ToString(DatabaseManager.ExecuteScalar(
            "SELECT Value FROM Settings WHERE Key = @k", DatabaseManager.P("@k", key)))
    End Function

    ' ---------- Creating a receipt (called inside the payment's transaction) ----------

    ''' <summary>
    ''' Creates the receipt number for a payment inside the caller's open transaction.
    ''' UNIQUE(PaymentId) means a payment can never get two receipts.
    ''' </summary>
    Public Sub CreateReceipt(conn As SqliteConnection, tx As SqliteTransaction,
                             paymentId As Long, issuedAt As String)
        Dim counter As Long
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText =
                "UPDATE Counters SET LastValue = LastValue + 1 " &
                "WHERE Name = 'ReceiptNumber' RETURNING LastValue"
            counter = Convert.ToInt64(cmd.ExecuteScalar())
        End Using

        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText =
                "INSERT INTO Receipts (ReceiptNumber, PaymentId, IssuedAt, PrintCount) " &
                "VALUES (@n, @p, @t, 0)"
            cmd.Parameters.Add(DatabaseManager.P("@n", "RC-" & counter.ToString("D3")))
            cmd.Parameters.Add(DatabaseManager.P("@p", paymentId))
            cmd.Parameters.Add(DatabaseManager.P("@t", issuedAt))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' ---------- Payment list ----------

    Public Function GetPayments(searchText As String) As DataTable
        SessionManager.RequireLogin()
        Dim q As String = If(searchText, "").Trim()
        Dim likeText As String = "%" & EscapeLike(q) & "%"

        Return DatabaseManager.GetTable(
            "SELECT p.PaymentId, COALESCE(rc.ReceiptNumber, '') AS ReceiptNumber, p.PaidAt, " &
            "p.RepairCode, COALESCE(r.CustomerName, '(removed)') AS CustomerName, " &
            "CAST(p.Amount AS REAL) AS Amount, p.PaymentMethod, e.FullName AS Employee, " &
            "COALESCE(rc.PrintCount, 0) AS PrintCount " &
            "FROM Payments p " &
            "JOIN Employees e ON e.EmployeeId = p.EmployeeId " &
            "LEFT JOIN Repairs r ON r.RepairCode = p.RepairCode " &
            "LEFT JOIN Receipts rc ON rc.PaymentId = p.PaymentId " &
           "WHERE (@owner = 0 OR p.EmployeeId = @owner) AND (@q = '' OR p.RepairCode LIKE @like ESCAPE '\' " &
            "OR r.CustomerName LIKE @like ESCAPE '\' " &
            "OR r.CustomerPhone LIKE @like ESCAPE '\' " &
            "OR rc.ReceiptNumber LIKE @like ESCAPE '\') " &
            "ORDER BY p.PaymentId DESC LIMIT 300",
            DatabaseManager.P("@q", q), DatabaseManager.P("@like", likeText),
            DatabaseManager.P("@owner", If(SessionManager.IsEmployee, SessionManager.CurrentUserId, 0)))
    End Function

    ' ---------- Receipt ----------

    ''' <summary>Returns the receipt for a saved payment. Older payments get their number once.</summary>
    Public Function GetReceipt(paymentId As Long) As ReceiptData
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can print receipts.")
        End If

        Dim payOwner As Object = DatabaseManager.ExecuteScalar(
            "SELECT EmployeeId FROM Payments WHERE PaymentId = @p", DatabaseManager.P("@p", paymentId))
        If payOwner Is Nothing OrElse IsDBNull(payOwner) OrElse
           Convert.ToInt32(payOwner) <> SessionManager.CurrentUserId Then Return Nothing

        EnsureReceiptNumber(paymentId)

        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT p.RepairCode, CAST(p.Amount AS REAL) AS Amount, p.PaymentMethod, p.PaidAt, " &
            "e.FullName AS EmployeeName, r.CustomerName, r.CustomerPhone, r.Brand, r.Model, " &
            "CAST(r.EstimatedCost AS REAL) AS EstimatedCost, " &
            "CAST(COALESCE((SELECT SUM(x.Amount) FROM Payments x " &
            "WHERE x.RepairCode = p.RepairCode AND x.PaymentId < p.PaymentId), 0) AS REAL) AS PreviousPaid, " &
            "rc.PrintCount, rc.ReceiptNumber " &
            "FROM Payments p " &
            "JOIN Employees e ON e.EmployeeId = p.EmployeeId " &
            "JOIN Repairs r ON r.RepairCode = p.RepairCode " &
            "JOIN Receipts rc ON rc.PaymentId = p.PaymentId " &
            "WHERE p.PaymentId = @id",
            DatabaseManager.P("@id", paymentId))
        If dt.Rows.Count = 0 Then Return Nothing

        Dim r As DataRow = dt.Rows(0)
        Dim d As New ReceiptData()
        d.ReceiptNumber = r("ReceiptNumber").ToString()
        d.PaymentId = paymentId

        Dim paidText As String = r("PaidAt").ToString()
        Dim parsed As DateTime
        If DateTime.TryParseExact(paidText, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture,
                                  DateTimeStyles.None, parsed) Then
            paidText = parsed.ToString("dd MMM yyyy HH:mm")
        End If
        d.PaidAt = paidText

        d.RepairCode = r("RepairCode").ToString()
        d.CustomerName = r("CustomerName").ToString()
        d.CustomerPhone = r("CustomerPhone").ToString()
        d.Brand = r("Brand").ToString()
        d.Model = r("Model").ToString()
        d.EstimatedCost = Convert.ToDecimal(Math.Round(Convert.ToDouble(r("EstimatedCost")), 2))
        d.PreviousPaid = Convert.ToDecimal(Math.Round(Convert.ToDouble(r("PreviousPaid")), 2))
        d.Amount = Convert.ToDecimal(Math.Round(Convert.ToDouble(r("Amount")), 2))
        d.Method = r("PaymentMethod").ToString()
        d.EmployeeName = r("EmployeeName").ToString()
        d.PrintCount = Convert.ToInt32(r("PrintCount"))
        d.BusinessName = Setting("BusinessName")
        d.BusinessPhone = Setting("BusinessPhone")
        d.BusinessAddress = Setting("BusinessAddress")
        Return d
    End Function

    ''' <summary>Safety net for payments saved before receipts were created automatically.</summary>
    Private Sub EnsureReceiptNumber(paymentId As Long)
        Dim existing As Object = DatabaseManager.ExecuteScalar(
            "SELECT ReceiptNumber FROM Receipts WHERE PaymentId = @p",
            DatabaseManager.P("@p", paymentId))
        If existing IsNot Nothing AndAlso Not IsDBNull(existing) Then Return

        Using conn As SqliteConnection = DatabaseManager.GetConnection()
            Using tx As SqliteTransaction = conn.BeginTransaction()
                Try
                    CreateReceipt(conn, tx, paymentId, DatabaseManager.NowText())
                    tx.Commit()
                Catch
                    tx.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    ''' <summary>Called after a REAL print (not a preview). Only raises the count.</summary>
    Public Sub RecordPrint(data As ReceiptData)
        Try
            DatabaseManager.ExecuteNonQuery(
                "UPDATE Receipts SET PrintCount = PrintCount + 1, LastPrintedAt = @t " &
                "WHERE PaymentId = @p",
                DatabaseManager.P("@t", DatabaseManager.NowText()),
                DatabaseManager.P("@p", data.PaymentId))
            ActivityLogger.Log("PrintReceipt",
                               "Printed receipt " & data.ReceiptNumber & " (" &
                               UIHelpers.Money(CDbl(data.Amount)) & ") for " & data.RepairCode & ".",
                               data.RepairCode)
        Catch
        End Try
    End Sub

End Module