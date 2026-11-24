Imports Microsoft.Data.Sqlite

''' <summary>Deleting a repair while keeping a permanent copy for administrators.</summary>
Public Module RepairDeletionService

    Public Const MaxReasonLength As Integer = 500

    ''' <summary>Returns "" on success, otherwise an error message.</summary>
    Public Function DeleteRepair(repairCode As String, reason As String) As String
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can delete repairs.")
        End If

        Dim code As String = If(repairCode, "").Trim()
        Dim why As String = If(reason, "").Trim()
        If code = "" Then Return "Select a repair first."
        If why.Length > MaxReasonLength Then
            Return "The reason is too long (" & MaxReasonLength & " characters maximum)."
        End If

        Dim err As String = ""
        Dim done As Boolean = False
        Dim customer As String = ""
        Dim stamp As String = DatabaseManager.NowText()

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        done = ApplyDelete(conn, tx, code, why, stamp, err, customer)
                        If done Then tx.Commit() Else tx.Rollback()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            ActivityLogger.Log("DeleteRepair",
                               "Failed to delete repair " & code & ": " & ex.Message, code, False)
            Return "The repair could not be deleted, so nothing was changed. (" & ex.Message & ")"
        End Try

        If Not done Then
            ActivityLogger.Log("DeleteRepair",
                               "Deletion of repair " & code & " was refused: " & err, code, False)
            Return err
        End If

        ' Logged AFTER the commit (SQLite allows one writer at a time).
        ActivityLogger.Log("DeleteRepair",
                           "Deleted repair " & code & " (" & customer & ")." &
                           If(why <> "", " Reason: " & why, ""), code)
        Return ""
    End Function

    Private Function ApplyDelete(conn As SqliteConnection, tx As SqliteTransaction,
                                 code As String, why As String, stamp As String,
                                 ByRef err As String, ByRef customerOut As String) As Boolean
        Dim customer As String = ""
        Dim phone As String = ""
        Dim brand As String = ""
        Dim model As String = ""
        Dim problem As String = ""
        Dim status As String = ""
        Dim received As String = ""
        Dim expected As String = ""
        Dim cost As Double = 0
        Dim paid As Double = 0
        Dim ownerId As Integer = 0

        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText =
                "SELECT r.CustomerName, r.CustomerPhone, r.Brand, r.Model, r.ProblemDescription, " &
                "r.Status, CAST(r.EstimatedCost AS REAL), r.DateReceived, r.ExpectedCollectionDate, " &
                "r.EmployeeId, CAST(COALESCE((SELECT SUM(p.Amount) FROM Payments p " &
                "WHERE p.RepairCode = r.RepairCode), 0) AS REAL) " &
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
                received = rd.GetString(7)
                expected = rd.GetString(8)
                ownerId = rd.GetInt32(9)
                paid = rd.GetDouble(10)
            End Using
        End Using
        customerOut = customer

        ' ----- Business rules -----
        If ownerId <> SessionManager.CurrentUserId Then
            err = "You can only delete repairs that you registered yourself."
            Return False
        End If
        If status <> "Received" Then
            err = "Only repairs still in the Received status can be deleted. This one is " & status & "."
            Return False
        End If
        If paid > 0.005 Then
            err = "This repair has payments recorded (" & UIHelpers.Money(paid) &
                  "), so it cannot be deleted. If the phone cannot be repaired, use Terminate Repair."
            Return False
        End If

        ' ----- 1. Permanent copy for administrators -----
        Run(conn, tx,
            "INSERT INTO DeletedRepairs (RepairCode, CustomerName, CustomerPhone, Brand, Model, " &
            "ProblemDescription, EstimatedCost, Status, DateReceived, ExpectedCollectionDate, " &
            "TotalPaidAtDeletion, OriginalEmployeeId, DeletedByEmployeeId, DeletedByName, " &
            "DeletedAt, DeletionReason) VALUES (@code, @name, @phone, @brand, @model, @problem, " &
            "@cost, @status, @rec, @exp, @paid, @orig, @by, @byName, @at, @reason)",
            DatabaseManager.P("@code", code),
            DatabaseManager.P("@name", customer),
            DatabaseManager.P("@phone", phone),
            DatabaseManager.P("@brand", brand),
            DatabaseManager.P("@model", model),
            DatabaseManager.P("@problem", problem),
            DatabaseManager.P("@cost", cost),
            DatabaseManager.P("@status", status),
            DatabaseManager.P("@rec", received),
            DatabaseManager.P("@exp", expected),
            DatabaseManager.P("@paid", paid),
            DatabaseManager.P("@orig", ownerId),
            DatabaseManager.P("@by", SessionManager.CurrentUserId),
            DatabaseManager.P("@byName", SessionManager.CurrentUserName),
            DatabaseManager.P("@at", stamp),
            DatabaseManager.P("@reason", If(why = "", Nothing, why)))

        ' ----- 2. Administrator audit entry -----
        Dim snapshot As String =
            "Repair " & code & " | Customer: " & customer & " (" & phone & ")" &
            " | Device: " & brand & " " & model & " | Problem: " & problem &
            " | Estimated: " & UIHelpers.Money(cost) & " | Paid: " & UIHelpers.Money(paid) &
            " | Status: " & status & " | Received: " & received & " | Expected: " & expected

        Run(conn, tx,
            "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, " &
            "PerformedByRole, RepairCode, OriginalData, Details, EventAt) " &
            "VALUES ('RepairDeleted', @id, @n, 'Employee', @code, @orig, @det, @t)",
            DatabaseManager.P("@id", SessionManager.CurrentUserId),
            DatabaseManager.P("@n", SessionManager.CurrentUserName),
            DatabaseManager.P("@code", code),
            DatabaseManager.P("@orig", snapshot),
            DatabaseManager.P("@det",
                SessionManager.CurrentUserName & " deleted repair " & code &
                If(why <> "", ". Reason: " & why, ". No reason given.")),
            DatabaseManager.P("@t", stamp))

        ' ----- 3. Delete the repair itself -----
        Dim rows As Integer = Run(conn, tx,
            "DELETE FROM Repairs WHERE RepairCode = @c AND Status = 'Received' AND EmployeeId = @e",
            DatabaseManager.P("@c", code),
            DatabaseManager.P("@e", SessionManager.CurrentUserId))
        If rows <> 1 Then
            err = "The repair changed while you were working. Search for it again."
            Return False
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