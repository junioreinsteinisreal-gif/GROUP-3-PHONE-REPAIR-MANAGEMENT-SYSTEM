Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.Data.Sqlite

Public Class HttpReq
    Public Property Method As String = ""
    Public Property Path As String = ""
    Public Property Query As New Dictionary(Of String, String)
    Public Property Body As String = ""
End Class

Public Class HttpResp
    Public Property Status As Integer = 200
    Public Property ContentType As String = "text/html; charset=utf-8"
    Public Property Body As String = ""
End Class

Public Class ReviewContext
    Public Property Business As String = ""
    Public Property Customer As String = ""
    Public Property CustomerFirst As String = ""
    Public Property Device As String = ""
    Public Property EmployeeId As Integer
    Public Property EmployeeName As String = ""
    Public Property EmployeeFirst As String = ""
End Class

''' <summary>A small web server inside the app that serves the customer review page.</summary>
Public Module ReviewServer

    Private listener As TcpListener
    Private cts As CancellationTokenSource
    Private ReadOnly hits As New Dictionary(Of String, List(Of DateTime))
    Private ReadOnly fails As New Dictionary(Of String, List(Of DateTime))
    Private ReadOnly lockObj As New Object()
    Public LastError As String = ""

    Public ReadOnly Property IsRunning As Boolean
        Get
            Return listener IsNot Nothing
        End Get
    End Property

    Private Function P(name As String, value As Object) As SqliteParameter
        Return DatabaseManager.P(name, value)
    End Function

    ' ---------- Settings ----------

    Private Function Setting(key As String, def As String) As String
        Dim v As String = AdminToolsService.GetSetting(key)
        Return If(String.IsNullOrEmpty(v), def, v)
    End Function

    Private Sub PutSetting(key As String, value As String)
        DatabaseManager.ExecuteNonQuery(
            "INSERT OR REPLACE INTO Settings (Key, Value) VALUES (@k, @v)", P("@k", key), P("@v", value))
    End Sub

    Public Function DemoMode() As Boolean
        Return Setting("ReviewDemoMode", "1") = "1"
    End Function

    Public Function ConfiguredPort() As Integer
        Dim n As Integer
        If Integer.TryParse(Setting("ReviewPort", "8080"), n) AndAlso n >= 1024 AndAlso n <= 65535 Then Return n
        Return 8080
    End Function

    Public Function LanIp() As String
        Try
            For Each ni As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
                If ni.OperationalStatus <> OperationalStatus.Up Then Continue For
                If ni.NetworkInterfaceType = NetworkInterfaceType.Loopback Then Continue For
                For Each ua As UnicastIPAddressInformation In ni.GetIPProperties().UnicastAddresses
                    If ua.Address.AddressFamily = AddressFamily.InterNetwork Then
                        Dim s As String = ua.Address.ToString()
                        If s.StartsWith("192.168.") OrElse s.StartsWith("10.") OrElse s.StartsWith("172.") Then Return s
                    End If
                Next
            Next
        Catch
        End Try
        Return "localhost"
    End Function

    Public Function BaseUrl() As String
        Dim custom As String = Setting("ReviewBaseUrl", "").Trim().TrimEnd("/"c)
        If custom <> "" Then Return custom
        Return "http://" & LanIp() & ":" & ConfiguredPort()
    End Function

    Public Function LinkFor(repairCode As String, pin As String) As String
        Return BaseUrl() & "/r/" & repairCode & "?k=" & pin
    End Function

    ''' <summary>Returns "" on success, otherwise an error message.</summary>
    Public Function SaveConfig(port As String, demo As Boolean, baseUrlText As String) As String
        SessionManager.RequireAdmin()
        Dim n As Integer
        If Not Integer.TryParse(If(port, "").Trim(), n) OrElse n < 1024 OrElse n > 65535 Then
            Return "Enter a port between 1024 and 65535, for example 8080."
        End If
        Dim b As String = If(baseUrlText, "").Trim()
        If b <> "" AndAlso Not b.StartsWith("http://", StringComparison.OrdinalIgnoreCase) AndAlso
           Not b.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
            Return "The public address must start with http:// or https://"
        End If
        Try
            PutSetting("ReviewPort", n.ToString())
            PutSetting("ReviewDemoMode", If(demo, "1", "0"))
            PutSetting("ReviewBaseUrl", b)
            DatabaseManager.ExecuteNonQuery(
                "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, PerformedByRole, " &
                "Details, EventAt) VALUES ('SettingsChanged', @id, @n, 'Administrator', @d, @t)",
                P("@id", SessionManager.CurrentUserId), P("@n", SessionManager.CurrentUserName),
                P("@d", "Updated the review website settings (port " & n & ", demo mode " & If(demo, "on", "off") & ")."),
                P("@t", DatabaseManager.NowText()))
        Catch ex As Exception
            Return "The settings could not be saved. (" & ex.Message & ")"
        End Try
        Return ""
    End Function

    Public Sub EnsureSchema()
        DatabaseManager.ExecuteNonQuery("
CREATE TABLE IF NOT EXISTS TipRequests (
    Ref         TEXT PRIMARY KEY,
    RepairCode  TEXT NOT NULL,
    FeedbackId  INTEGER,
    EmployeeId  INTEGER NOT NULL,
    Amount      REAL NOT NULL CHECK (Amount > 0),
    Network     TEXT NOT NULL,
    Phone       TEXT NOT NULL,
    Status      TEXT NOT NULL DEFAULT 'Pending',
    Demo        INTEGER NOT NULL DEFAULT 1,
    CreatedAt   TEXT NOT NULL,
    PaidAt      TEXT
);")
    End Sub

    ' ---------- Start and stop ----------

    ''' <summary>Returns "" on success, otherwise an error message.</summary>
    Public Function StartServer() As String
        StopServer()
        Try
            EnsureSchema()
            Dim lst As New TcpListener(IPAddress.Any, ConfiguredPort())
            lst.Start()
            listener = lst
            cts = New CancellationTokenSource()
            Dim token As CancellationToken = cts.Token
            Task.Run(Function() AcceptLoop(lst, token))
            LastError = ""
            Return ""
        Catch ex As Exception
            listener = Nothing
            LastError = ex.Message
            Return ex.Message
        End Try
    End Function

    Public Sub StopServer()
        Try
            If cts IsNot Nothing Then cts.Cancel()
            If listener IsNot Nothing Then listener.Stop()
        Catch
        End Try
        listener = Nothing
        cts = Nothing
    End Sub

    Private Async Function AcceptLoop(lst As TcpListener, token As CancellationToken) As Task
        While Not token.IsCancellationRequested
            Dim client As TcpClient = Nothing
            Try
                client = Await lst.AcceptTcpClientAsync(token)
            Catch
                Exit While
            End Try
            Dim c As TcpClient = client
            Task.Run(Function() HandleClient(c))
        End While
    End Function

    ' ---------- Reading and writing HTTP ----------

    Private Async Function HandleClient(client As TcpClient) As Task
        Using client
            Try
                Dim ip As String = CType(client.Client.RemoteEndPoint, IPEndPoint).Address.ToString()
                Using stream As NetworkStream = client.GetStream()
                    Using timeout As New CancellationTokenSource(TimeSpan.FromSeconds(10))
                        Dim req As HttpReq = Await ReadRequest(stream, timeout.Token)
                        If req Is Nothing Then Return
                        Dim resp As HttpResp = Route(req, ip)
                        Await WriteResponse(stream, resp, timeout.Token)
                    End Using
                End Using
            Catch
            End Try
        End Using
    End Function

    Private Function IndexOfHeaderEnd(buf() As Byte, count As Integer) As Integer
        For i As Integer = 0 To count - 4
            If buf(i) = 13 AndAlso buf(i + 1) = 10 AndAlso buf(i + 2) = 13 AndAlso buf(i + 3) = 10 Then Return i
        Next
        Return -1
    End Function

    Private Async Function ReadRequest(stream As NetworkStream, token As CancellationToken) As Task(Of HttpReq)
        Dim buf(16383) As Byte
        Dim total As Integer = 0
        Dim headerEnd As Integer = -1
        While headerEnd < 0 AndAlso total < buf.Length
            Dim n As Integer = Await stream.ReadAsync(buf, total, buf.Length - total, token)
            If n = 0 Then Return Nothing
            total += n
            headerEnd = IndexOfHeaderEnd(buf, total)
        End While
        If headerEnd < 0 Then Return Nothing

        Dim head As String = Encoding.ASCII.GetString(buf, 0, headerEnd)
        Dim lines() As String = head.Split({vbCrLf}, StringSplitOptions.None)
        Dim parts() As String = lines(0).Split(" "c)
        If parts.Length < 2 Then Return Nothing

        Dim req As New HttpReq()
        req.Method = parts(0).ToUpperInvariant()
        Dim target As String = parts(1)
        Dim qi As Integer = target.IndexOf("?"c)
        If qi >= 0 Then
            req.Path = target.Substring(0, qi)
            For Each pair As String In target.Substring(qi + 1).Split("&"c)
                Dim eq As Integer = pair.IndexOf("="c)
                If eq > 0 Then
                    req.Query(WebUtility.UrlDecode(pair.Substring(0, eq))) = WebUtility.UrlDecode(pair.Substring(eq + 1))
                End If
            Next
        Else
            req.Path = target
        End If

        Dim contentLength As Integer = 0
        For i As Integer = 1 To lines.Length - 1
            If lines(i).StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase) Then
                Integer.TryParse(lines(i).Substring(15).Trim(), contentLength)
            End If
        Next
        If contentLength < 0 OrElse contentLength > 8192 Then Return Nothing

        If contentLength > 0 Then
            Dim bodyStart As Integer = headerEnd + 4
            Dim have As Integer = Math.Min(total - bodyStart, contentLength)
            Dim body(contentLength - 1) As Byte
            If have > 0 Then Array.Copy(buf, bodyStart, body, 0, have)
            Dim got As Integer = have
            While got < contentLength
                Dim n As Integer = Await stream.ReadAsync(body, got, contentLength - got, token)
                If n = 0 Then Return Nothing
                got += n
            End While
            req.Body = Encoding.UTF8.GetString(body)
        End If
        Return req
    End Function

    Private Async Function WriteResponse(stream As NetworkStream, resp As HttpResp,
                                         token As CancellationToken) As Task
        Dim bodyBytes() As Byte = Encoding.UTF8.GetBytes(resp.Body)
        Dim statusText As String = If(resp.Status = 200, "OK", If(resp.Status = 404, "Not Found",
                                    If(resp.Status = 429, "Too Many Requests", "OK")))
        Dim head As String = "HTTP/1.1 " & resp.Status & " " & statusText & vbCrLf &
                             "Content-Type: " & resp.ContentType & vbCrLf &
                             "Content-Length: " & bodyBytes.Length & vbCrLf &
                             "Cache-Control: no-store" & vbCrLf &
                             "X-Content-Type-Options: nosniff" & vbCrLf &
                             "Connection: close" & vbCrLf & vbCrLf
        Dim headBytes() As Byte = Encoding.ASCII.GetBytes(head)
        Await stream.WriteAsync(headBytes, 0, headBytes.Length, token)
        Await stream.WriteAsync(bodyBytes, 0, bodyBytes.Length, token)
        Await stream.FlushAsync(token)
    End Function

    ' ---------- Routing ----------

    Private Function Html(status As Integer, body As String) As HttpResp
        Dim r As New HttpResp()
        r.Status = status
        r.Body = body
        Return r
    End Function

    Private Function JsonReply(obj As Object) As HttpResp
        Dim r As New HttpResp()
        r.ContentType = "application/json; charset=utf-8"
        r.Body = JsonSerializer.Serialize(obj, obj.GetType())
        Return r
    End Function

    Private Function Fail(message As String) As HttpResp
        Return JsonReply(New With {.ok = False, .message = message})
    End Function

    Private Function RateLimited(ip As String) As Boolean
        SyncLock lockObj
            If hits.Count > 500 Then hits.Clear()
            Dim list As List(Of DateTime) = Nothing
            If Not hits.TryGetValue(ip, list) Then
                list = New List(Of DateTime)
                hits(ip) = list
            End If
            Dim nowT As DateTime = DateTime.Now
            list.RemoveAll(Function(t) (nowT - t).TotalSeconds > 60)
            If list.Count >= 60 Then Return True
            list.Add(nowT)
            Return False
        End SyncLock
    End Function

    Private Function IsLocked(code As String) As Boolean
        SyncLock lockObj
            Dim list As List(Of DateTime) = Nothing
            If Not fails.TryGetValue(code, list) Then Return False
            Dim nowT As DateTime = DateTime.Now
            list.RemoveAll(Function(t) (nowT - t).TotalMinutes > 15)
            Return list.Count >= 10
        End SyncLock
    End Function

    Private Sub RegisterFail(code As String)
        SyncLock lockObj
            If fails.Count > 500 Then fails.Clear()
            Dim list As List(Of DateTime) = Nothing
            If Not fails.TryGetValue(code, list) Then
                list = New List(Of DateTime)
                fails(code) = list
            End If
            list.Add(DateTime.Now)
        End SyncLock
    End Sub

    Private Function Route(req As HttpReq, ip As String) As HttpResp
        If RateLimited(ip) Then Return Html(429, SimplePage("Please wait", "Too many requests. Please try again in a minute."))
        Try
            If req.Method = "GET" Then
                If req.Path = "/" Then Return Html(200, SimplePage("Reviews", "The customer review service is running."))
                Dim m As Match = Regex.Match(req.Path, "^/r/(PR-\d{3,8})$", RegexOptions.IgnoreCase)
                If m.Success Then
                    Dim k As String = ""
                    req.Query.TryGetValue("k", k)
                    Return ReviewPage(m.Groups(1).Value.ToUpperInvariant(), If(k, ""))
                End If
                Return Html(404, SimplePage("Not found", "This page does not exist."))
            ElseIf req.Method = "POST" Then
                Select Case req.Path
                    Case "/api/submit"
                        Return ApiSubmit(req.Body)
                    Case "/api/approve"
                        Return ApiApprove(req.Body)
                    Case "/api/cancel"
                        Return ApiCancel(req.Body)
                End Select
            End If
            Return Html(404, SimplePage("Not found", "This page does not exist."))
        Catch ex As Exception
            Return Html(200, SimplePage("Sorry", "Something went wrong. Please try again later."))
        End Try
    End Function

    ' ---------- Validation of a review link ----------

    Private Function FirstWord(text As String) As String
        Dim t As String = If(text, "").Trim()
        If t = "" Then Return ""
        Return t.Split(" "c)(0)
    End Function

    Private Function LoadContext(code As String, pin As String, ByRef err As String) As ReviewContext
        err = ""
        If IsLocked(code) Then
            err = "Too many wrong attempts. Please try again later."
            Return Nothing
        End If

        Dim stored As Object = DatabaseManager.ExecuteScalar(
            "SELECT Pin FROM FeedbackInvites WHERE RepairCode = @c", P("@c", code))
        If stored Is Nothing OrElse IsDBNull(stored) OrElse Not Regex.IsMatch(pin, "^\d{4}$") OrElse
           stored.ToString() <> pin Then
            RegisterFail(code)
            err = "This link is not valid. Please use the link from your message."
            Return Nothing
        End If

        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT r.CustomerName, r.Brand || ' ' || r.Model AS Device, r.Status, r.EmployeeId, " &
            "e.FullName AS EmpName FROM Repairs r JOIN Employees e ON e.EmployeeId = r.EmployeeId " &
            "WHERE r.RepairCode = @c", P("@c", code))
        If dt.Rows.Count = 0 Then
            err = "This repair could not be found."
            Return Nothing
        End If
        Dim row As DataRow = dt.Rows(0)
        Dim status As String = row("Status").ToString()
        If status <> "Completed" AndAlso status <> "Collected" Then
            err = "Your repair is not ready for a review yet."
            Return Nothing
        End If

        Dim existing As Integer = Convert.ToInt32(DatabaseManager.ExecuteScalar(
            "SELECT COUNT(*) FROM Feedback WHERE RepairCode = @c", P("@c", code)))
        If existing > 0 Then
            err = "Thank you - we already have your review."
            Return Nothing
        End If

        Dim ctx As New ReviewContext()
        ctx.Business = Setting("BusinessName", "Phone Repair")
        ctx.Customer = row("CustomerName").ToString()
        ctx.CustomerFirst = FirstWord(ctx.Customer)
        ctx.Device = row("Device").ToString()
        ctx.EmployeeId = Convert.ToInt32(row("EmployeeId"))
        ctx.EmployeeName = row("EmpName").ToString()
        ctx.EmployeeFirst = FirstWord(ctx.EmployeeName)
        Return ctx
    End Function

    ''' <summary>Creates (or reuses) the PIN that protects a repair's review link.</summary>
    Public Function GetOrCreatePin(repairCode As String) As String
        Dim o As Object = DatabaseManager.ExecuteScalar(
            "SELECT Pin FROM FeedbackInvites WHERE RepairCode = @c", P("@c", repairCode))
        If o IsNot Nothing AndAlso Not IsDBNull(o) Then Return o.ToString()
        DatabaseManager.ExecuteNonQuery(
            "INSERT OR IGNORE INTO FeedbackInvites (RepairCode, Pin, CreatedAt) VALUES (@c, @p, @t)",
            P("@c", repairCode), P("@p", RandomNumberGenerator.GetInt32(1000, 10000).ToString()),
            P("@t", DatabaseManager.NowText()))
        Return Convert.ToString(DatabaseManager.ExecuteScalar(
            "SELECT Pin FROM FeedbackInvites WHERE RepairCode = @c", P("@c", repairCode)))
    End Function

    ' ---------- The page ----------

    Private Function SimplePage(title As String, message As String) As String
        Return "<!DOCTYPE html><html><head><meta charset='utf-8'>" &
               "<meta name='viewport' content='width=device-width, initial-scale=1'><title>" &
               WebUtility.HtmlEncode(title) & "</title></head>" &
               "<body style='font-family:Segoe UI,Arial,sans-serif;background:#f0f2f5;margin:0'>" &
               "<div style='background:#0a193c;color:#fff;padding:20px;text-align:center;font-size:20px'>" &
               WebUtility.HtmlEncode(Setting("BusinessName", "")) & "</div>" &
               "<div style='max-width:460px;margin:24px auto;background:#fff;border-radius:14px;padding:22px;" &
               "text-align:center;font-size:17px'>" & WebUtility.HtmlEncode(message) & "</div></body></html>"
    End Function

    Private Function ReviewPage(code As String, pin As String) As HttpResp
        Dim err As String = ""
        Dim ctx As ReviewContext = LoadContext(code, pin, err)
        If ctx Is Nothing Then Return Html(200, SimplePage("Review", err))

        Dim demo As Boolean = DemoMode()
        Dim cfg As Object = New With {.code = code, .pin = pin, .biz = ctx.Business, .name = ctx.CustomerFirst,
                                      .device = ctx.Device, .employee = ctx.EmployeeFirst,
                                      .demo = demo, .tipsOn = demo}
        Dim json As String = JsonSerializer.Serialize(cfg, cfg.GetType())
        Return Html(200, PageTemplate.Replace("__CONFIG__", json))
    End Function

    ' ---------- Submitting a review ----------

    Private Function JStr(root As JsonElement, name As String) As String
        Dim el As JsonElement
        If root.TryGetProperty(name, el) AndAlso el.ValueKind = JsonValueKind.String Then Return el.GetString()
        Return ""
    End Function

    Private Function JNum(root As JsonElement, name As String) As Decimal
        Dim el As JsonElement
        If root.TryGetProperty(name, el) AndAlso el.ValueKind = JsonValueKind.Number Then
            Dim d As Decimal
            If el.TryGetDecimal(d) Then Return d
        End If
        Return 0D
    End Function

    Private Function Scalar(conn As SqliteConnection, tx As SqliteTransaction, sql As String,
                            ParamArray params() As SqliteParameter) As Object
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText = sql
            For Each prm As SqliteParameter In params
                cmd.Parameters.Add(prm)
            Next
            Return cmd.ExecuteScalar()
        End Using
    End Function

    Private Function ApiSubmit(body As String) As HttpResp
        Dim root As JsonElement
        Try
            root = JsonDocument.Parse(body).RootElement
        Catch
            Return Fail("Something went wrong. Please try again.")
        End Try

        Dim code As String = JStr(root, "code").Trim().ToUpperInvariant()
        Dim pin As String = JStr(root, "pin").Trim()
        Dim ratingD As Decimal = JNum(root, "rating")
        Dim comment As String = JStr(root, "comment").Trim()
        Dim tip As Decimal = JNum(root, "tip")
        Dim network As String = JStr(root, "network").Trim()
        Dim phoneRaw As String = JStr(root, "phone").Trim()

        If Not Regex.IsMatch(code, "^PR-\d{3,8}$") Then Return Fail("This link is not valid.")
        Dim err As String = ""
        Dim ctx As ReviewContext = LoadContext(code, pin, err)
        If ctx Is Nothing Then Return Fail(err)

        If ratingD < 1D OrElse ratingD > 5D OrElse ratingD <> Math.Floor(ratingD) Then
            Return Fail("Please tap a star rating from 1 to 5.")
        End If
        Dim rating As Integer = CInt(ratingD)
        comment = Regex.Replace(comment, "[\x00-\x08\x0B\x0C\x0E-\x1F]", "")
        If comment.Length > FeedbackService.MaxComment Then Return Fail("Your comment is too long.")

        Dim demo As Boolean = DemoMode()
        If tip < 0D OrElse tip > 500D OrElse tip <> Math.Round(tip, 2) Then Return Fail("Enter a tip between 1 and 500.")
        Dim tipOn As Boolean = (tip > 0D AndAlso demo)
        Dim phone As String = ""
        If tipOn Then
            If network <> "MTN" AndAlso network <> "Telecel" AndAlso network <> "AirtelTigo" Then
                Return Fail("Choose your mobile money network.")
            End If
            phone = NotificationService.NormalizePhone(phoneRaw, Setting("CountryCode", "233"))
            If phone = "" Then Return Fail("Enter a valid mobile money number.")
        End If

        Dim tipRef As String = ""
        Dim feedbackId As Long = 0
        Dim stamp As String = DatabaseManager.NowText()

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        Dim existing As Integer = Convert.ToInt32(Scalar(conn, tx,
                            "SELECT COUNT(*) FROM Feedback WHERE RepairCode = @c", P("@c", code)))
                        If existing > 0 Then
                            tx.Rollback()
                            Return Fail("We already have your review. Thank you!")
                        End If

                        feedbackId = Convert.ToInt64(Scalar(conn, tx,
                            "INSERT INTO Feedback (RepairCode, CustomerName, EmployeeId, Rating, Comment, Verified, RecordedAt) " &
                            "VALUES (@c, @n, @e, @r, @cm, 1, @t) RETURNING FeedbackId",
                            P("@c", code), P("@n", ctx.Customer), P("@e", ctx.EmployeeId), P("@r", rating),
                            P("@cm", If(comment = "", Nothing, comment)), P("@t", stamp)))

                        If tipOn Then
                            tipRef = Convert.ToHexString(RandomNumberGenerator.GetBytes(8))
                            Scalar(conn, tx,
                                "INSERT INTO TipRequests (Ref, RepairCode, FeedbackId, EmployeeId, Amount, Network, Phone, " &
                                "Status, Demo, CreatedAt) VALUES (@ref, @c, @f, @e, @a, @net, @ph, 'Pending', 1, @t)",
                                P("@ref", tipRef), P("@c", code), P("@f", feedbackId), P("@e", ctx.EmployeeId),
                                P("@a", CDbl(tip)), P("@net", network), P("@ph", phone), P("@t", stamp))
                        End If

                        Scalar(conn, tx,
                            "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, PerformedByRole, " &
                            "RepairCode, Details, EventAt) VALUES ('CustomerReview', NULL, @n, 'Customer', @c, @d, @t)",
                            P("@n", ctx.Customer), P("@c", code),
                            P("@d", ctx.Customer & " rated repair " & code & " " & rating & " star(s) for " &
                                    ctx.EmployeeName & " on the review website." &
                                    If(comment <> "", " Comment: " & comment, "") &
                                    If(tipOn, " A tip of " & UIHelpers.Money(CDbl(tip)) & " is awaiting approval.", "")),
                            P("@t", stamp))
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            Return Fail("Your review could not be saved. Please try again.")
        End Try

        ActivityLogger.LogFor(ctx.EmployeeId, ctx.EmployeeName, "Feedback",
                              "Customer review received on the website for " & code & ": " & rating & " star(s).", code)

        If tipOn Then
            Return JsonReply(New With {.ok = True, .reference = tipRef, .amount = CDbl(tip),
                                       .phone = "*** *** " & phone.Substring(phone.Length - 3),
                                       .message = "Thank you for your review!"})
        End If
        Return JsonReply(New With {.ok = True, .reference = "", .amount = 0.0, .phone = "",
                                   .message = If(tip > 0D, "Thank you for your review! Tips by mobile money are not switched on yet.",
                                                 "Thank you for your review!")})
    End Function

    ' ---------- Tip approval ----------

    ''' <summary>Marks a tip as paid and credits the employee. A live payment provider would call this from its callback.</summary>
    Public Function MarkTipPaid(tipRef As String, ByRef message As String) As Boolean
        message = ""
        Dim code As String = ""
        Dim feedbackId As Long = 0
        Dim empId As Integer = 0
        Dim amount As Double = 0
        Dim stamp As String = DatabaseManager.NowText()
        Dim empName As String = ""

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        Dim found As Boolean = False
                        Using cmd As SqliteCommand = conn.CreateCommand()
                            cmd.Transaction = tx
                            cmd.CommandText =
                                "SELECT t.RepairCode, COALESCE(t.FeedbackId, 0), t.EmployeeId, CAST(t.Amount AS REAL), e.FullName " &
                                "FROM TipRequests t JOIN Employees e ON e.EmployeeId = t.EmployeeId " &
                                "WHERE t.Ref = @r AND t.Status = 'Pending'"
                            cmd.Parameters.Add(P("@r", tipRef))
                            Using rd As SqliteDataReader = cmd.ExecuteReader()
                                If rd.Read() Then
                                    found = True
                                    code = rd.GetString(0)
                                    feedbackId = rd.GetInt64(1)
                                    empId = rd.GetInt32(2)
                                    amount = rd.GetDouble(3)
                                    empName = rd.GetString(4)
                                End If
                            End Using
                        End Using
                        If Not found Then
                            tx.Rollback()
                            message = "This tip request is no longer open."
                            Return False
                        End If

                        Scalar(conn, tx,
                            "INSERT INTO Tips (RepairCode, FeedbackId, EmployeeId, Amount, Method, ReceivedAt) " &
                            "VALUES (@c, @f, @e, @a, 'Mobile Money', @t)",
                            P("@c", code), P("@f", If(feedbackId = 0, Nothing, CObj(feedbackId))),
                            P("@e", empId), P("@a", amount), P("@t", stamp))
                        Scalar(conn, tx, "UPDATE TipRequests SET Status = 'Paid', PaidAt = @t WHERE Ref = @r",
                               P("@t", stamp), P("@r", tipRef))
                        Scalar(conn, tx,
                            "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, PerformedByRole, " &
                            "RepairCode, Details, EventAt) VALUES ('TipReceived', NULL, 'Customer', 'Customer', @c, @d, @t)",
                            P("@c", code),
                            P("@d", "A tip of " & UIHelpers.Money(amount) & " by mobile money was approved for " & empName &
                                    " (repair " & code & ")."),
                            P("@t", stamp))
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
        Catch ex As Exception
            message = "The tip could not be recorded."
            Return False
        End Try

        ActivityLogger.LogFor(empId, empName, "Tip",
                              "Received a " & UIHelpers.Money(amount) & " tip from a customer by mobile money on " & code & ".", code)
        Return True
    End Function

    Private Function ApiApprove(body As String) As HttpResp
        If Not DemoMode() Then Return Fail("Demo approval is switched off.")
        Dim r As String = ""
        Try
            r = JStr(JsonDocument.Parse(body).RootElement, "ref")
        Catch
        End Try
        If Not Regex.IsMatch(r, "^[0-9A-F]{16}$") Then Return Fail("This request is not valid.")
        Dim msg As String = ""
        If MarkTipPaid(r, msg) Then
            Return JsonReply(New With {.ok = True, .message = "Your tip was sent. Thank you!"})
        End If
        Return Fail(msg)
    End Function

    Private Function ApiCancel(body As String) As HttpResp
        Dim r As String = ""
        Try
            r = JStr(JsonDocument.Parse(body).RootElement, "ref")
        Catch
        End Try
        If Regex.IsMatch(r, "^[0-9A-F]{16}$") Then
            DatabaseManager.ExecuteNonQuery(
                "UPDATE TipRequests SET Status = 'Cancelled' WHERE Ref = @r AND Status = 'Pending'", P("@r", r))
        End If
        Return JsonReply(New With {.ok = True, .message = "No tip was sent. Thank you for your review!"})
    End Function

    ' ---------- Administrator tools ----------

    ''' <summary>Returns "" on success. Gives the phone link (same Wi-Fi) and the link for this PC.</summary>
    Public Function CreateTestLink(repairCode As String, ByRef phoneLink As String,
                                   ByRef pcLink As String) As String
        SessionManager.RequireAdmin()
        phoneLink = ""
        pcLink = ""
        Dim code As String = If(repairCode, "").Trim().ToUpperInvariant()
        If Not Regex.IsMatch(code, "^PR-\d{3,8}$") Then Return "Enter a repair code such as PR-001."

        Dim dt As DataTable = DatabaseManager.GetTable(
            "SELECT Status FROM Repairs WHERE RepairCode = @c", P("@c", code))
        If dt.Rows.Count = 0 Then Return "That repair does not exist."
        Dim st As String = dt.Rows(0)("Status").ToString()
        If st <> "Completed" AndAlso st <> "Collected" Then
            Return "The repair must be Completed or Collected first (it is " & st & ")."
        End If
        If Convert.ToInt32(DatabaseManager.ExecuteScalar(
            "SELECT COUNT(*) FROM Feedback WHERE RepairCode = @c", P("@c", code))) > 0 Then
            Return "This repair already has a review. Use a different repair."
        End If

        Dim pin As String = GetOrCreatePin(code)
        phoneLink = LinkFor(code, pin)
        pcLink = "http://localhost:" & ConfiguredPort() & "/r/" & code & "?k=" & pin
        Return ""
    End Function

    Public Function GetTipRequests() As DataTable
        SessionManager.RequireAdmin()
        EnsureSchema()
        Return DatabaseManager.GetTable(
            "SELECT t.CreatedAt, t.RepairCode, e.FullName AS Employee, CAST(t.Amount AS REAL) AS Amount, " &
            "t.Network, t.Status, CASE WHEN t.Demo = 1 THEN 'Demo' ELSE 'Live' END AS Mode " &
            "FROM TipRequests t JOIN Employees e ON e.EmployeeId = t.EmployeeId " &
            "ORDER BY t.CreatedAt DESC LIMIT 200")
    End Function

    ' ---------- The web page (single quotes only, so it fits inside a VB string) ----------

    Private ReadOnly PageTemplate As String = "<!DOCTYPE html>
<html lang='en'><head><meta charset='utf-8'>
<meta name='viewport' content='width=device-width, initial-scale=1'>
<title>Rate your repair</title>
<style>
*{box-sizing:border-box}
body{margin:0;font-family:Segoe UI,Roboto,Arial,sans-serif;background:#f0f2f5;color:#282c34}
.top{background:#0a193c;color:#fff;padding:22px 18px;text-align:center}
.top h1{margin:0;font-size:21px}
.top p{margin:8px 0 0;color:#a0b9eb;font-size:15px}
.wrap{max-width:480px;margin:0 auto;padding:14px}
.card{background:#fff;border-radius:14px;padding:18px;margin-bottom:14px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
h2{font-size:17px;margin:0 0 12px;color:#0a193c}
.stars{display:flex;gap:6px}
.stars button{flex:1;font-size:34px;line-height:1;background:#f0f2f5;border:2px solid transparent;border-radius:10px;padding:8px 0;color:#c4cbd6}
.stars button.on{color:#f5a623;background:#fff6e0;border-color:#f5a623}
textarea,input,select{width:100%;font-size:16px;padding:12px;border:1px solid #cdd3dc;border-radius:10px;font-family:inherit}
textarea{min-height:100px;resize:vertical}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px}
.chip{padding:10px 14px;border:2px solid #2553c8;color:#2553c8;border-radius:999px;background:#fff;font-weight:700;font-size:15px}
.chip.on{background:#2553c8;color:#fff}
label{display:block;font-weight:700;font-size:14px;margin:12px 0 6px}
.btn{width:100%;padding:15px;border:0;border-radius:12px;font-size:17px;font-weight:700;color:#fff;background:#22783c}
.btn:disabled{opacity:.5}
.ghost{background:#e4e8ee;color:#282c34}
.err{color:#a0141e;font-weight:700;margin:10px 0;min-height:20px}
.muted{color:#64748b;font-size:14px;margin:0 0 10px}
.ov{position:fixed;left:0;top:0;right:0;bottom:0;background:rgba(10,25,60,.78);display:none;align-items:center;justify-content:center;padding:18px}
.ov.show{display:flex}
.phone{background:#fff;border-radius:16px;padding:22px;max-width:340px;width:100%;text-align:center}
.ussd{background:#1b1f27;color:#e6e6e6;border-radius:10px;padding:16px;text-align:left;font-family:Consolas,monospace;font-size:15px;margin:12px 0;white-space:pre-line}
.row{display:flex;gap:10px}
</style></head><body>
<div class='top'><h1 id='biz'></h1><p id='hello'></p></div>
<div class='wrap'>
<div id='form'>
<div class='card'><h2>How was our service?</h2><div class='stars' id='stars'></div></div>
<div class='card'><h2>Tell us more (optional)</h2><textarea id='comment' maxlength='300' placeholder='What did we do well? What could be better?'></textarea></div>
<div class='card' id='tipCard'><h2>Say thanks with a tip (optional)</h2>
<p class='muted' id='tipNote'></p>
<div class='chips' id='chips'></div>
<div id='custom' style='display:none'><label>Tip amount (&#8373;)</label><input id='amt' type='number' min='1' max='500' step='0.5' inputmode='decimal'></div>
<div id='momo' style='display:none'><label>Mobile money network</label>
<select id='net'><option>MTN</option><option>Telecel</option><option>AirtelTigo</option></select>
<label>Mobile money number</label><input id='phone' type='tel' inputmode='tel' maxlength='20' placeholder='024 123 4567'></div>
</div>
<div class='err' id='err'></div>
<button class='btn' id='send' type='button'>Send my review</button>
</div>
<div id='thanks' class='card' style='display:none;text-align:center'><h2>Thank you!</h2><p id='thanksText'></p></div>
</div>
<div class='ov' id='ov'><div class='phone'>
<h2>MoMo payment prompt</h2>
<p class='muted' id='demoNote'></p>
<div class='ussd' id='ussd'></div>
<div class='row'><button class='btn ghost' id='ovNo' type='button'>Cancel</button><button class='btn' id='ovYes' type='button'>Approve</button></div>
</div></div>
<script>
var C=__CONFIG__;
var rating=0,sel=0,ref='',tipAmt=0;
function $(i){return document.getElementById(i);}
function money(n){return '\u20B5'+Number(n).toFixed(2);}
function post(url,obj){return fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(obj)}).then(function(x){return x.json();});}
$('biz').textContent=C.biz;
$('hello').textContent='Hi '+C.name+', how was your '+C.device+' repair?';
var sb=$('stars');
for(var i=1;i<=5;i++){(function(n){var b=document.createElement('button');b.type='button';b.textContent='\u2605';
b.onclick=function(){rating=n;var all=sb.children;for(var k=0;k<all.length;k++){all[k].className=(k<n)?'on':'';}};sb.appendChild(b);})(i);}
function readTip(){if(sel===-1){tipAmt=parseFloat($('amt').value)||0;}else{tipAmt=sel;}return tipAmt;}
if(!C.tipsOn){$('tipCard').style.display='none';}else{
$('tipNote').textContent='Your tip goes to '+C.employee+'.';
var opts=[0,5,10,20,50,-1];var chipEls=[];
opts.forEach(function(v){var c=document.createElement('button');c.type='button';c.className='chip'+(v===0?' on':'');
c.textContent=(v===0)?'No tip':(v===-1)?'Other':'\u20B5'+v;
c.onclick=function(){sel=v;chipEls.forEach(function(e){e.className='chip';});c.className='chip on';
$('custom').style.display=(v===-1)?'block':'none';$('momo').style.display=(v===0)?'none':'block';};
chipEls.push(c);$('chips').appendChild(c);});}
function done(msg){$('form').style.display='none';$('ov').className='ov';$('thanks').style.display='block';$('thanksText').textContent=msg;window.scrollTo(0,0);}
$('send').onclick=function(){
$('err').textContent='';
if(!rating){$('err').textContent='Please tap a star rating.';return;}
var t=C.tipsOn?readTip():0;
if(t>0&&!$('phone').value.trim()){$('err').textContent='Enter your mobile money number for the tip.';return;}
$('send').disabled=true;
post('/api/submit',{code:C.code,pin:C.pin,rating:rating,comment:$('comment').value,tip:t,network:$('net').value,phone:$('phone').value})
.then(function(r){
if(!r.ok){$('err').textContent=r.message;$('send').disabled=false;return;}
if(r.reference){ref=r.reference;
$('demoNote').textContent='Demo mode: with a live provider this prompt appears on '+r.phone+'.';
$('ussd').textContent='Pay '+money(r.amount)+' to '+C.biz+'?\nEnter your MoMo PIN to approve.';
$('ov').className='ov show';}else{done(r.message);}})
.catch(function(){$('err').textContent='Connection problem. Please try again.';$('send').disabled=false;});};
$('ovYes').onclick=function(){$('ovYes').disabled=true;
post('/api/approve',{ref:ref}).then(function(r){done(r.message);}).catch(function(){done('Thank you for your review!');});};
$('ovNo').onclick=function(){post('/api/cancel',{ref:ref}).then(function(r){done(r.message);}).catch(function(){done('Thank you for your review!');});};
</script></body></html>"

End Module