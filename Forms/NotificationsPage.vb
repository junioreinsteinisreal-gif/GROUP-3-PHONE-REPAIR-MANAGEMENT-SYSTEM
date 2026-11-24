''' <summary>Administrator: SMS and email settings, test buttons and the message log.</summary>
Public Class NotificationsPage
    Inherits UserControl

    Private chkSms As CheckBox
    Private txtSmsUrl As TextBox
    Private txtSender As TextBox
    Private txtCountry As TextBox
    Private txtTestPhone As TextBox
    Private chkMail As CheckBox
    Private txtHost As TextBox
    Private txtPort As TextBox
    Private chkSsl As CheckBox
    Private txtUser As TextBox
    Private txtPass As TextBox
    Private txtFrom As TextBox
    Private txtTestMail As TextBox
    Private lblSaveMsg As Label
    Private gridLog As DataGridView
    Private lblLogMsg As Label
    Private ReadOnly boldFont As Font = UITheme.FontBold(10.5F)

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsAdmin Then
            Throw New UnauthorizedAccessException("Only administrators can manage customer messages.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(6)
        BuildPage()
        LoadSettings()
        LoadLog()
    End Sub

    ' ---------- Helpers ----------

    Private Function Field(labelText As String, ctrl As Control, height As Integer) As Panel
        Dim pnl As New Panel()
        pnl.Dock = DockStyle.Top
        pnl.Height = height
        pnl.Padding = New Padding(0, 0, 0, 6)
        ctrl.Dock = DockStyle.Top
        Dim lbl As New Label()
        lbl.AutoSize = False
        lbl.Dock = DockStyle.Top
        lbl.Height = 28
        lbl.TextAlign = ContentAlignment.BottomLeft
        lbl.Font = UITheme.FontBold(10.5F)
        lbl.ForeColor = UITheme.TextDark
        lbl.Text = labelText
        pnl.Controls.Add(ctrl)
        pnl.Controls.Add(lbl)
        Return pnl
    End Function

    Private Function Note(text As String, height As Integer) As Label
        Dim lbl As New Label()
        lbl.Dock = DockStyle.Top
        lbl.Height = height
        lbl.AutoSize = False
        lbl.Font = UITheme.FontRegular(9.5F)
        lbl.ForeColor = UITheme.TextMuted
        lbl.Text = text
        Return lbl
    End Function

    Private Function Heading(text As String) As Label
        Dim lbl As New Label()
        lbl.Dock = DockStyle.Top
        lbl.Height = 44
        lbl.AutoSize = False
        lbl.TextAlign = ContentAlignment.MiddleLeft
        lbl.Font = UITheme.FontBold(14)
        lbl.ForeColor = UITheme.Navy
        lbl.Text = text
        Return lbl
    End Function

    Private Function CheckRow(box As CheckBox) As Panel
        Dim pnl As New Panel()
        pnl.Dock = DockStyle.Top
        pnl.Height = 46
        box.AutoSize = True
        box.Left = 0
        box.Top = 10
        box.Font = UITheme.FontBold(11)
        pnl.Controls.Add(box)
        Return pnl
    End Function

    Private Function ButtonRow(btn As Button) As Panel
        Dim pnl As New Panel()
        pnl.Dock = DockStyle.Top
        pnl.Height = 64
        btn.Left = 0
        btn.Top = 8
        pnl.Controls.Add(btn)
        Return pnl
    End Function

    Private Sub StackTop(host As Control, items As List(Of Control))
        For i As Integer = items.Count - 1 To 0 Step -1
            host.Controls.Add(items(i))
        Next
    End Sub

    Private Function MakeCard() As Panel
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Margin = New Padding(4)
        card.Padding = New Padding(18)
        card.AutoScroll = True
        Return card
    End Function

    ' ---------- Building ----------

    Private Sub BuildPage()
        Dim tabs As New TabControl()
        tabs.Dock = DockStyle.Fill
        tabs.Font = UITheme.FontBold(10.5F)
        tabs.SizeMode = TabSizeMode.Fixed
        tabs.ItemSize = New Size(240, 42)

        Dim t1 As New TabPage("Settings")
        t1.BackColor = UITheme.LightGrey
        t1.Controls.Add(BuildSettingsTab())
        Dim t2 As New TabPage("Message Log")
        t2.BackColor = UITheme.LightGrey
        t2.Controls.Add(BuildLogTab())

        tabs.TabPages.Add(t1)
        tabs.TabPages.Add(t2)
        Me.Controls.Add(tabs)
    End Sub

    Private Function BuildSettingsTab() As Control
        Dim host As New Panel()
        host.Dock = DockStyle.Fill

        Dim tbl As New TableLayoutPanel()
        tbl.Dock = DockStyle.Fill
        tbl.ColumnCount = 2
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        tbl.RowCount = 1
        tbl.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        ' ----- SMS card -----
        Dim sms As Panel = MakeCard()
        chkSms = New CheckBox()
        chkSms.Text = "Send text messages (SMS) to customers"
        txtSmsUrl = UITheme.MakeTextBox(100, 11)
        txtSender = UITheme.MakeTextBox(100, 12)
        txtSender.MaxLength = 11
        txtCountry = UITheme.MakeTextBox(100, 12)
        txtCountry.MaxLength = 4
        txtTestPhone = UITheme.MakeTextBox(100, 12)
        txtTestPhone.MaxLength = 20
        Dim btnTestSms As Button = UITheme.MakeButton("Save && Send Test SMS", UITheme.RoyalBlue, 280, 48)
        btnTestSms.Text = "Save & Send Test SMS"
        AddHandler btnTestSms.Click, AddressOf TestSmsClick

        Dim smsItems As New List(Of Control)
        smsItems.Add(Heading("Text messages (SMS)"))
        smsItems.Add(CheckRow(chkSms))
        smsItems.Add(Note("Needs an account with an SMS provider (for example Arkesel, mNotify or Hubtel). " &
                          "Paste the provider's 'send SMS' web address (it must start with https://) and put " &
                          "{to}, {message} and {sender} where the phone number, text and sender name go. " &
                          "Typical shape: https://sms.example.com/send?key=YOUR_KEY&to={to}&from={sender}&sms={message}  " &
                          "Check your provider's documentation for the exact address.", 150))
        smsItems.Add(Field("SMS gateway address", txtSmsUrl, 80))
        smsItems.Add(Field("Sender name (max 11 characters)", txtSender, 80))
        smsItems.Add(Field("Country calling code (Ghana = 233)", txtCountry, 80))
        smsItems.Add(Field("Phone number for the test", txtTestPhone, 80))
        smsItems.Add(ButtonRow(btnTestSms))
        StackTop(sms, smsItems)

        ' ----- Email card -----
        Dim mail As Panel = MakeCard()
        chkMail = New CheckBox()
        chkMail.Text = "Send emails to customers (when they gave an email)"
        txtHost = UITheme.MakeTextBox(100, 12)
        txtPort = UITheme.MakeTextBox(100, 12)
        txtPort.MaxLength = 5
        chkSsl = New CheckBox()
        chkSsl.Text = "Use SSL / TLS (keep ticked for Gmail)"
        txtUser = UITheme.MakeTextBox(100, 12)
        txtPass = UITheme.MakeTextBox(100, 12)
        txtPass.UseSystemPasswordChar = True
        txtFrom = UITheme.MakeTextBox(100, 12)
        txtTestMail = UITheme.MakeTextBox(100, 12)
        txtTestMail.MaxLength = 80
        Dim btnTestMail As Button = UITheme.MakeButton("Save && Send Test Email", UITheme.RoyalBlue, 280, 48)
        btnTestMail.Text = "Save & Send Test Email"
        AddHandler btnTestMail.Click, AddressOf TestMailClick

        Dim mailItems As New List(Of Control)
        mailItems.Add(Heading("Email"))
        mailItems.Add(CheckRow(chkMail))
        mailItems.Add(Note("Free option with Gmail: server smtp.gmail.com, port 587, SSL on. Turn on 2-step " &
                           "verification, create an App Password in your Google account, and paste it below " &
                           "(never your normal password).", 90))
        mailItems.Add(Field("Email server (SMTP host)", txtHost, 80))
        mailItems.Add(Field("Port", txtPort, 80))
        mailItems.Add(CheckRow(chkSsl))
        mailItems.Add(Field("Username (usually the email address)", txtUser, 80))
        mailItems.Add(Field("Password / App Password", txtPass, 80))
        mailItems.Add(Field("Sender email address", txtFrom, 80))
        mailItems.Add(Field("Email address for the test", txtTestMail, 80))
        mailItems.Add(ButtonRow(btnTestMail))
        StackTop(mail, mailItems)

        tbl.Controls.Add(sms, 0, 0)
        tbl.Controls.Add(mail, 1, 0)

        ' ----- Bottom: save -----
        Dim bottom As New Panel()
        bottom.Dock = DockStyle.Bottom
        bottom.Height = 84
        bottom.BackColor = UITheme.CardWhite
        Dim btnSave As Button = UITheme.MakeButton("Save Settings", UITheme.SuccessGreen, 220, 50)
        btnSave.Left = 20
        btnSave.Top = 16
        AddHandler btnSave.Click, Sub(s, ev) SaveSettings()
        lblSaveMsg = New Label()
        lblSaveMsg.AutoSize = False
        lblSaveMsg.Left = 260
        lblSaveMsg.Top = 8
        lblSaveMsg.Width = 700
        lblSaveMsg.Height = 66
        lblSaveMsg.TextAlign = ContentAlignment.MiddleLeft
        lblSaveMsg.Font = UITheme.FontBold(10.5F)
        lblSaveMsg.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
        bottom.Controls.Add(btnSave)
        bottom.Controls.Add(lblSaveMsg)

        host.Controls.Add(tbl)
        host.Controls.Add(bottom)
        Return host
    End Function

    Private Function BuildLogTab() As Control
        Dim host As New Panel()
        host.Dock = DockStyle.Fill

        gridLog = New DataGridView()
        gridLog.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(gridLog)
        UIHelpers.AddTextColumn(gridLog, "CreatedAt", "Queued", 115)
        UIHelpers.AddTextColumn(gridLog, "RepairCode", "Repair", 70)
        UIHelpers.AddTextColumn(gridLog, "Kind", "Message", 90)
        UIHelpers.AddTextColumn(gridLog, "Channel", "Channel", 65)
        UIHelpers.AddTextColumn(gridLog, "Recipient", "To", 130)
        UIHelpers.AddTextColumn(gridLog, "Status", "Status", 75)
        UIHelpers.AddTextColumn(gridLog, "Attempts", "Tries", 45, True)
        UIHelpers.AddTextColumn(gridLog, "LastError", "Problem", 200)
        UIHelpers.AddTextColumn(gridLog, "SentAt", "Sent", 115)
        AddHandler gridLog.CellFormatting, AddressOf LogCellFormatting
        AddHandler gridLog.CellDoubleClick, AddressOf LogDoubleClick

        lblLogMsg = New Label()
        lblLogMsg.Dock = DockStyle.Bottom
        lblLogMsg.Height = 40
        lblLogMsg.AutoSize = False
        lblLogMsg.TextAlign = ContentAlignment.MiddleLeft
        lblLogMsg.Font = UITheme.FontRegular(10)
        lblLogMsg.ForeColor = UITheme.TextMuted
        lblLogMsg.Text = "Double-click a row to read the message. Messages that fail 3 times are marked Failed."

        Dim bar As New FlowLayoutPanel()
        bar.Dock = DockStyle.Top
        bar.AutoSize = True
        bar.AutoSizeMode = AutoSizeMode.GrowAndShrink
        bar.Padding = New Padding(0, 4, 0, 6)

        Dim btnRefresh As Button = UITheme.MakeButton("Refresh", UITheme.Navy, 120, 38)
        AddHandler btnRefresh.Click, Sub(s, ev) LoadLog()
        Dim btnSend As Button = UITheme.MakeButton("Send Pending Now", UITheme.RoyalBlue, 200, 38)
        AddHandler btnSend.Click,
            Sub(s, ev)
                NotificationService.ProcessPendingAsync()
                lblLogMsg.Text = "Sending in the background. Click Refresh in a few seconds."
            End Sub
        Dim btnRetry As Button = UITheme.MakeButton("Retry Failed", UITheme.WarningAmber, 160, 38)
        btnRetry.ForeColor = UITheme.TextDark
        AddHandler btnRetry.Click,
            Sub(s, ev)
                Dim n As Integer = NotificationService.RetryFailed()
                lblLogMsg.Text = n & " failed message(s) queued again. Click Refresh in a few seconds."
            End Sub
        bar.Controls.AddRange(New Control() {btnRefresh, btnSend, btnRetry})

        host.Controls.Add(gridLog)
        host.Controls.Add(lblLogMsg)
        host.Controls.Add(bar)
        Return host
    End Function

    ' ---------- Settings ----------

    Private Sub LoadSettings()
        chkSms.Checked = (NotificationService.GetSetting("SmsEnabled") = "1")
        txtSmsUrl.Text = NotificationService.GetSetting("SmsUrl")
        txtSender.Text = NotificationService.GetSetting("SmsSender")
        txtCountry.Text = NotificationService.GetSetting("CountryCode")
        chkMail.Checked = (NotificationService.GetSetting("EmailEnabled") = "1")
        txtHost.Text = NotificationService.GetSetting("SmtpHost")
        txtPort.Text = NotificationService.GetSetting("SmtpPort")
        chkSsl.Checked = (NotificationService.GetSetting("SmtpSsl") = "1")
        txtUser.Text = NotificationService.GetSetting("SmtpUser")
        txtFrom.Text = NotificationService.GetSetting("SmtpFrom")
        txtPass.Clear()
        txtPass.PlaceholderText = If(NotificationService.GetSetting("SmtpPassword") <> "",
                                     "(saved - leave empty to keep it)", "")
    End Sub

    Private Sub ShowSaveMsg(text As String, color As Color)
        lblSaveMsg.ForeColor = color
        lblSaveMsg.Text = text
    End Sub

    Private Function SaveSettings() As Boolean
        Dim err As String = NotificationService.SaveSettings(
            chkSms.Checked, txtSmsUrl.Text, txtSender.Text, txtCountry.Text, chkMail.Checked,
            txtHost.Text, txtPort.Text, chkSsl.Checked, txtUser.Text, txtPass.Text, txtFrom.Text)
        If err <> "" Then
            ShowSaveMsg(err, UITheme.DangerRed)
            Return False
        End If
        ShowSaveMsg("Settings saved. Messages are sent automatically from now on.", UITheme.SuccessGreen)
        LoadSettings()
        Return True
    End Function

    Private Sub TestSmsClick(sender As Object, e As EventArgs)
        If Not SaveSettings() Then Return
        If txtTestPhone.Text.Trim() = "" Then
            ShowSaveMsg("Enter a phone number to send the test to.", UITheme.DangerRed)
            Return
        End If
        Dim err As String = ""
        Me.Cursor = Cursors.WaitCursor
        Try
            err = NotificationService.SendTestSms(txtTestPhone.Text)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        If err = "" Then
            ShowSaveMsg("The gateway accepted the test SMS. Check the phone.", UITheme.SuccessGreen)
        Else
            ShowSaveMsg("SMS test failed: " & err, UITheme.DangerRed)
        End If
    End Sub

    Private Sub TestMailClick(sender As Object, e As EventArgs)
        If Not SaveSettings() Then Return
        Dim err As String = ""
        Me.Cursor = Cursors.WaitCursor
        Try
            err = NotificationService.SendTestEmail(txtTestMail.Text)
        Finally
            Me.Cursor = Cursors.Default
        End Try
        If err = "" Then
            ShowSaveMsg("Test email sent. Check the inbox (and spam).", UITheme.SuccessGreen)
        Else
            ShowSaveMsg("Email test failed: " & err, UITheme.DangerRed)
        End If
    End Sub

    ' ---------- Log ----------

    Private Sub LoadLog()
        Try
            gridLog.DataSource = NotificationService.GetLog()
        Catch ex As Exception
            lblLogMsg.Text = "Could not load the log: " & ex.Message
        End Try
    End Sub

    Private Sub LogCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 OrElse e.Value Is Nothing OrElse IsDBNull(e.Value) Then Return
        If gridLog.Columns(e.ColumnIndex).DataPropertyName <> "Status" Then Return
        Select Case e.Value.ToString()
            Case "Sent"
                e.CellStyle.ForeColor = UITheme.SuccessGreen
            Case "Failed"
                e.CellStyle.ForeColor = UITheme.DangerRed
            Case "Pending"
                e.CellStyle.ForeColor = Color.FromArgb(176, 112, 0)
            Case Else
                e.CellStyle.ForeColor = UITheme.TextMuted
        End Select
        e.CellStyle.Font = boldFont
    End Sub

    Private Sub LogDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        Dim drv As DataRowView = TryCast(gridLog.Rows(e.RowIndex).DataBoundItem, DataRowView)
        If drv Is Nothing Then Return
        MessageBox.Show(drv("Body").ToString() &
                        If(drv("LastError").ToString() <> "", vbCrLf & vbCrLf & "Problem: " & drv("LastError").ToString(), ""),
                        "Message to " & drv("Recipient").ToString(),
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class