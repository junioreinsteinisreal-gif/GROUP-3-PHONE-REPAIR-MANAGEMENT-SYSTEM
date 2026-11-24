Imports System.Globalization
Imports System.Net
Imports System.Net.Http
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Microsoft.Data.Sqlite

''' <summary>Automatic customer messages (SMS and email) when a repair is received, In Progress or Completed.</summary>
Public Module NotificationService

    Private ReadOnly sendLock As New Object()
    Private ReadOnly http As HttpClient = CreateClient()

    Private Function CreateClient() As HttpClient
        Dim c As New HttpClient()
        c.Timeout = TimeSpan.FromSeconds(20)
        Return c
    End Function

    Private Function P(name As String, value As Object) As SqliteParameter
        Return DatabaseManager.P(name, value)
    End Function

    Public Function GetSetting(key As String) As String
        Return AdminToolsService.GetSetting(key)
    End Function

    ' ---------- Small helpers ----------

    Public Function IsValidEmail(text As String) As Boolean
        Dim t As String = If(text, "").Trim()
        If t = "" OrElse t.Length > 80 Then Return False
        Try
            Dim addr As New MailAddress(t)
            Return addr.Address = t AndAlso t.Contains(".")
        Catch
            Return False
        End Try
    End Function

    ''' <summary>Saves the optional customer email right after a repair is registered.</summary>
    Public Sub SaveCustomerEmail(repairCode As String, email As String)
        Dim mail As String = If(email, "").Trim()
        If mail = "" OrElse Not IsValidEmail(mail) Then Return
        Try
            DatabaseManager.ExecuteNonQuery(
                "UPDATE Repairs SET CustomerEmail = @e WHERE RepairCode = @c",
                P("@e", mail), P("@c", repairCode))
        Catch
        End Try
    End Sub

    ''' <summary>International digits only (for example 0244123456 becomes 233244123456). "" if invalid.</summary>
    Public Function NormalizePhone(raw As String, countryCode As String) As String
        Dim digits As String = Regex.Replace(If(raw, ""), "\D", "")
        Dim cc As String = Regex.Replace(If(countryCode, ""), "\D", "")
        If digits.StartsWith("00") Then
            digits = digits.Substring(2)
        ElseIf digits.StartsWith("0") AndAlso cc <> "" Then
            digits = cc & digits.Substring(1)
        ElseIf cc <> "" AndAlso Not digits.StartsWith(cc) AndAlso digits.Length <= 9 Then
            digits = cc & digits
        End If
        If digits.Length < 10 OrElse digits.Length > 15 Then Return ""
        Return digits
    End Function

    Private Function FirstWord(text As String) As String
        Dim t As String = If(text, "").Trim()
        If t = "" Then Return ""
        Return t.Split(" "c)(0)
    End Function

    Private Function BuildMessage(kind As String, customer As String, device As String,
                                  code As String, balance As Double, employee As String,
                                  pin As String, biz As String) As String
        If kind = "Received" Then
            Return biz & ": Hi " & FirstWord(customer) & ", we have received your " & device &
                   " (ref " & code & "). We will message you as the repair progresses."
        End If
        If kind = "InProgress" Then
            Return biz & ": Hi " & FirstWord(customer) & ", we have started work on your " &
                   device & " (ref " & code & "). We will message you when it is ready."
        End If
        Dim bal As String = If(balance > 0.005,
                               "Balance due: GHS " & balance.ToString("N2", CultureInfo.InvariantCulture),
                               "No balance due")
        Return biz & ": Hi " & FirstWord(customer) & ", your " & device & " (ref " & code &
               ") is ready for collection. " & bal & ". Rate our service or tip " &
               FirstWord(employee) & ": " & ReviewServer.LinkFor(code, pin)
    End Function

    Private Sub Queue(repairCode As String, kind As String, channel As String,
                      recipient As String, subject As String, body As String,
                      status As String, lastError As String)
        DatabaseManager.ExecuteNonQuery(
            "INSERT OR IGNORE INTO Notifications (RepairCode, Kind, Channel, Recipient, Subject, " &
            "Body, Status, Attempts, LastError, CreatedAt) " &
            "VALUES (@c, @k, @ch, @r, @s, @b, @st, 0, @err, @t)",
            P("@c", repairCode), P("@k", kind), P("@ch", channel), P("@r", recipient),
            P("@s", subject), P("@b", body), P("@st", status), P("@err", lastError),
            P("@t", DatabaseManager.NowText()))
    End Sub

    ' ---------- Called after a repair is registered or its status changes ----------

    Public Sub OnStatusChanged(repairCode As String, newStatus As String)
        Dim kind As String = ""
        If newStatus = "Received" Then kind = "Received"
        If newStatus = "In Progress" Then kind = "InProgress"
        If newStatus = "Completed" Then kind = "Completed"
        If kind = "" Then Return

        Try
            Dim smsOn As Boolean = (GetSetting("SmsEnabled") = "1")
            Dim mailOn As Boolean = (GetSetting("EmailEnabled") = "1")
            If Not smsOn AndAlso Not mailOn Then
                Queue(repairCode, kind, "System", "(none)", Nothing, "No message was sent.", "Skipped",
                      "SMS and email are both switched off in Customer Messages > Settings.")
                Return
            End If

            Dim dt As DataTable = DatabaseManager.GetTable(
                "SELECT r.CustomerName, r.CustomerPhone, COALESCE(r.CustomerEmail, '') AS Email, " &
                "r.Brand || ' ' || r.Model AS Device, CAST(r.EstimatedCost AS REAL) AS Cost, " &
                "CAST(COALESCE((SELECT SUM(p.Amount) FROM Payments p WHERE p.RepairCode = r.RepairCode), 0) AS REAL) AS Paid, " &
                "e.FullName AS Employee FROM Repairs r JOIN Employees e ON e.EmployeeId = r.EmployeeId " &
                "WHERE r.RepairCode = @c", P("@c", repairCode))
            If dt.Rows.Count = 0 Then Return

            Dim r As DataRow = dt.Rows(0)
            Dim pin As String = If(kind = "Completed", ReviewServer.GetOrCreatePin(repairCode), "")
            Dim biz As String = GetSetting("BusinessName")
            Dim body As String = BuildMessage(kind, r("CustomerName").ToString(), r("Device").ToString(),
                                              repairCode, Convert.ToDouble(r("Cost")) - Convert.ToDouble(r("Paid")),
                                              r("Employee").ToString(), pin, biz)
            Dim subject As String = If(kind = "Received",
                                       "We received your phone for repair (" & repairCode & ")",
                                       "Update on your phone repair (" & repairCode & ")")

            If smsOn Then
                Dim rawPhone As String = r("CustomerPhone").ToString()
                Dim phone As String = NormalizePhone(rawPhone, GetSetting("CountryCode"))
                If phone = "" Then
                    Queue(repairCode, kind, "SMS", rawPhone, Nothing, body, "Skipped", "Invalid phone number")
                Else
                    Queue(repairCode, kind, "SMS", phone, Nothing, body, "Pending", Nothing)
                End If
            End If

            If mailOn Then
                Dim email As String = r("Email").ToString()
                If email = "" Then
                    Queue(repairCode, kind, "Email", "(none)", subject, body, "Skipped",
                          "This repair has no customer email address.")
                ElseIf Not IsValidEmail(email) Then
                    Queue(repairCode, kind, "Email", email, subject, body, "Skipped",
                          "The customer email address is not valid.")
                Else
                    Queue(repairCode, kind, "Email", email, subject, body, "Pending", Nothing)
                End If
            End If

            ProcessPendingAsync()
        Catch ex As Exception
            Try
                Queue(repairCode, kind, "System", "(none)", Nothing, "The message could not be prepared.",
                      "Skipped", Left(ex.Message, 250))
            Catch
            End Try
        End Try
    End Sub

    ' ---------- Sending ----------

    Public Function SendSms(rawPhone As String, text As String) As String
        Dim template As String = GetSetting("SmsUrl").Trim()
        If template = "" Then Return "The SMS gateway address is not set."
        Dim phone As String = NormalizePhone(rawPhone, GetSetting("CountryCode"))
        If phone = "" Then Return "Invalid phone number."

        Dim url As String = template.Replace("{to}", phone).
            Replace("{message}", Uri.EscapeDataString(text)).
            Replace("{sender}", Uri.EscapeDataString(GetSetting("SmsSender")))
        Try
            Using resp As HttpResponseMessage = http.GetAsync(url).GetAwaiter().GetResult()
                Dim reply As String = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                If resp.IsSuccessStatusCode Then Return ""
                Return "Gateway replied " & CInt(resp.StatusCode) & ": " & Left(reply, 150)
            End Using
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function SendEmail(toAddress As String, subject As String, body As String) As String
        Dim host As String = GetSetting("SmtpHost").Trim()
        Dim fromAddr As String = GetSetting("SmtpFrom").Trim()
        If host = "" OrElse fromAddr = "" Then Return "The email server and sender address are not set."
        Dim port As Integer
        If Not Integer.TryParse(GetSetting("SmtpPort"), port) Then port = 587

        Try
            Using msg As New MailMessage(New MailAddress(fromAddr, GetSetting("BusinessName")),
                                         New MailAddress(toAddress))
                msg.Subject = subject
                msg.Body = body
                msg.IsBodyHtml = False
                Using client As New SmtpClient(host, port)
                    client.EnableSsl = (GetSetting("SmtpSsl") = "1")
                    client.Timeout = 20000
                    Dim user As String = GetSetting("SmtpUser")
                    If user <> "" Then client.Credentials = New NetworkCredential(user, GetSetting("SmtpPassword"))
                    client.Send(msg)
                End Using
            End Using
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Sub ProcessPendingAsync()
        Task.Run(Sub() ProcessPending())
    End Sub

    ''' <summary>Sends queued messages. A message that fails 3 times is marked Failed.</summary>
    Public Function ProcessPending() As Integer
        If Not System.Threading.Monitor.TryEnter(sendLock) Then Return 0
        Dim sent As Integer = 0
        Try
            Dim dt As DataTable = DatabaseManager.GetTable(
                "SELECT NotificationId, Channel, Recipient, COALESCE(Subject, '') AS Subject, " &
                "Body, Attempts FROM Notifications WHERE Status = 'Pending' " &
                "ORDER BY NotificationId LIMIT 25")
            For Each r As DataRow In dt.Rows
                Dim id As Long = Convert.ToInt64(r("NotificationId"))
                Dim err As String = ""
                If r("Channel").ToString() = "SMS" Then
                    err = SendSms(r("Recipient").ToString(), r("Body").ToString())
                Else
                    err = SendEmail(r("Recipient").ToString(), r("Subject").ToString(), r("Body").ToString())
                End If

                If err = "" Then
                    DatabaseManager.ExecuteNonQuery(
                        "UPDATE Notifications SET Status = 'Sent', SentAt = @t, LastError = NULL, " &
                        "Attempts = Attempts + 1 WHERE NotificationId = @id",
                        P("@t", DatabaseManager.NowText()), P("@id", id))
                    sent += 1
                Else
                    Dim attempts As Integer = Convert.ToInt32(r("Attempts")) + 1
                    DatabaseManager.ExecuteNonQuery(
                        "UPDATE Notifications SET Status = @s, Attempts = @a, LastError = @e " &
                        "WHERE NotificationId = @id",
                        P("@s", If(attempts >= 3, "Failed", "Pending")), P("@a", attempts),
                        P("@e", Left(err, 300)), P("@id", id))
                End If
            Next
        Catch
        Finally
            System.Threading.Monitor.Exit(sendLock)
        End Try
        Return sent
    End Function

    ' ---------- Administrator tools ----------

    Public Function GetLog() As DataTable
        SessionManager.RequireAdmin()
        Return DatabaseManager.GetTable(
            "SELECT NotificationId, CreatedAt, RepairCode, Kind, Channel, Recipient, Status, " &
            "Attempts, COALESCE(LastError, '') AS LastError, COALESCE(SentAt, '') AS SentAt, Body " &
            "FROM Notifications ORDER BY NotificationId DESC LIMIT 500")
    End Function

    Public Function RetryFailed() As Integer
        SessionManager.RequireAdmin()
        Dim n As Integer = DatabaseManager.ExecuteNonQuery(
            "UPDATE Notifications SET Status = 'Pending', Attempts = 0 WHERE Status = 'Failed'")
        ProcessPendingAsync()
        Return n
    End Function

    Public Function SendTestSms(phone As String) As String
        SessionManager.RequireAdmin()
        Return SendSms(phone, GetSetting("BusinessName") & ": this is a test message from the repair system.")
    End Function

    Public Function SendTestEmail(address As String) As String
        SessionManager.RequireAdmin()
        If Not IsValidEmail(address) Then Return "Enter a valid email address."
        Return SendEmail(address, "Test message from " & GetSetting("BusinessName"),
                         "This is a test message from the repair system. Email sending works.")
    End Function

    Private Sub PutSetting(conn As SqliteConnection, tx As SqliteTransaction,
                           key As String, value As String)
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText = "INSERT OR REPLACE INTO Settings (Key, Value) VALUES (@k, @v)"
            cmd.Parameters.Add(P("@k", key))
            cmd.Parameters.Add(P("@v", value))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>Returns "" on success, otherwise an error message. An empty password keeps the saved one.</summary>
    Public Function SaveSettings(smsOn As Boolean, smsUrl As String, smsSender As String,
                                 countryCode As String, mailOn As Boolean, host As String,
                                 port As String, ssl As Boolean, user As String,
                                 password As String, fromAddr As String) As String
        SessionManager.RequireAdmin()
        Dim url As String = If(smsUrl, "").Trim()
        Dim cc As String = Regex.Replace(If(countryCode, ""), "\D", "")
        Dim portNum As Integer

        If cc.Length < 1 OrElse cc.Length > 4 Then Return "Enter the country calling code, for example 233."
        If smsOn Then
            If Not url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                Return "The SMS address must start with https://"
            End If
            If Not url.Contains("{to}") OrElse Not url.Contains("{message}") Then
                Return "The SMS address must contain {to} and {message}."
            End If
        End If
        If mailOn Then
            If If(host, "").Trim() = "" Then Return "Enter the email server (SMTP host)."
            If Not Integer.TryParse(If(port, "").Trim(), portNum) OrElse portNum < 1 OrElse portNum > 65535 Then
                Return "Enter a valid email port, for example 587."
            End If
            If Not IsValidEmail(fromAddr) Then Return "Enter a valid sender email address."
        End If

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        PutSetting(conn, tx, "SmsEnabled", If(smsOn, "1", "0"))
                        PutSetting(conn, tx, "SmsUrl", url)
                        PutSetting(conn, tx, "SmsSender", If(smsSender, "").Trim())
                        PutSetting(conn, tx, "CountryCode", cc)
                        PutSetting(conn, tx, "EmailEnabled", If(mailOn, "1", "0"))
                        PutSetting(conn, tx, "SmtpHost", If(host, "").Trim())
                        PutSetting(conn, tx, "SmtpPort", If(port, "587").Trim())
                        PutSetting(conn, tx, "SmtpSsl", If(ssl, "1", "0"))
                        PutSetting(conn, tx, "SmtpUser", If(user, "").Trim())
                        PutSetting(conn, tx, "SmtpFrom", If(fromAddr, "").Trim())
                        If If(password, "") <> "" Then PutSetting(conn, tx, "SmtpPassword", password)
                        AdminToolsService.WriteAudit(conn, tx, "SettingsChanged",
                            "Updated customer message settings (SMS " & If(smsOn, "on", "off") &
                            ", email " & If(mailOn, "on", "off") & ").")
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
            Return ""
        Catch ex As Exception
            Return "The settings could not be saved. (" & ex.Message & ")"
        End Try
    End Function

End Module