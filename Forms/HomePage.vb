''' <summary>Dashboard. Employees see only their own figures; administrators see the whole shop.</summary>
Public Class HomePage
    Inherits UserControl

    Private ReadOnly isAdminView As Boolean
    Private ReadOnly refreshTimer As New System.Windows.Forms.Timer()
    Private tickCount As Integer = 0

    Private lblDateTime As Label
    Private lblLoginStatus As Label
    Private lblBanner As Label
    Private lblPaySummary As Label
    Private ReadOnly stats As New Dictionary(Of String, Label)
    Private gridActivity As DataGridView
    Private gridAttention As DataGridView
    Private gridSecond As DataGridView
    Private ReadOnly boldFont As Font = UITheme.FontBold(10)

    Public Sub New()
        SessionManager.RequireLogin()
        isAdminView = SessionManager.IsAdmin
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(6)

        BuildPage()
        LoadData()
        UpdateClock()

        refreshTimer.Interval = 1000
        AddHandler refreshTimer.Tick, AddressOf OnTick
        refreshTimer.Start()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            refreshTimer.Stop()
            refreshTimer.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' ---------- Layout ----------

    Private Sub BuildPage()
        Dim root As New TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.ColumnCount = 1
        root.RowCount = 4
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 84.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 116.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 52.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        root.Controls.Add(BuildInfoStrip(), 0, 0)
        root.Controls.Add(BuildCards(), 0, 1)
        root.Controls.Add(BuildBanner(), 0, 2)
        root.Controls.Add(BuildBody(), 0, 3)
        Me.Controls.Add(root)
    End Sub

    Private Function InfoCell(caption As String, value As String, ByRef valueLabel As Label) As Panel
        Dim pnl As New Panel()
        pnl.Dock = DockStyle.Fill
        pnl.Padding = New Padding(14, 8, 8, 4)

        valueLabel = New Label()
        valueLabel.Dock = DockStyle.Fill
        valueLabel.AutoSize = False
        valueLabel.AutoEllipsis = True
        valueLabel.TextAlign = ContentAlignment.TopLeft
        valueLabel.Font = UITheme.FontBold(11.5F)
        valueLabel.ForeColor = UITheme.Navy
        valueLabel.Text = value

        Dim cap As New Label()
        cap.Dock = DockStyle.Top
        cap.Height = 26
        cap.AutoSize = False
        cap.TextAlign = ContentAlignment.MiddleLeft
        cap.Font = UITheme.FontRegular(9)
        cap.ForeColor = UITheme.TextMuted
        cap.Text = caption

        pnl.Controls.Add(valueLabel)
        pnl.Controls.Add(cap)
        Return pnl
    End Function

    Private Function BuildInfoStrip() As Control
        Dim strip As New TableLayoutPanel()
        strip.Dock = DockStyle.Fill
        strip.BackColor = UITheme.CardWhite
        strip.Margin = New Padding(4)
        strip.ColumnCount = 4
        strip.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 27.0F))
        strip.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 14.0F))
        strip.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 31.0F))
        strip.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 28.0F))
        strip.RowCount = 1
        strip.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        Dim lblName As Label = Nothing
        Dim lblId As Label = Nothing
        strip.Controls.Add(InfoCell(If(isAdminView, "Administrator", "Employee"),
                                    SessionManager.CurrentUserName, lblName), 0, 0)
        strip.Controls.Add(InfoCell(If(isAdminView, "Role", "Employee ID"),
                                    If(isAdminView, "Admin", "#" & SessionManager.CurrentUserId),
                                    lblId), 1, 0)
        strip.Controls.Add(InfoCell("Date and time", "", lblDateTime), 2, 0)
        strip.Controls.Add(InfoCell("Login status", "", lblLoginStatus), 3, 0)
        Return strip
    End Function

    Private Function StatCard(key As String, caption As String, color As Color) As Panel
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Margin = New Padding(4)

        Dim value As New Label()
        value.Dock = DockStyle.Fill
        value.AutoSize = False
        value.TextAlign = ContentAlignment.MiddleCenter
        value.Font = UITheme.FontBold(15)
        value.ForeColor = color
        value.Text = "0"
        stats(key) = value

        Dim cap As New Label()
        cap.Dock = DockStyle.Bottom
        cap.Height = 48
        cap.AutoSize = False
        cap.TextAlign = ContentAlignment.TopCenter
        cap.Font = UITheme.FontRegular(9.5F)
        cap.ForeColor = UITheme.TextMuted
        cap.Text = caption

        card.Controls.Add(value)
        card.Controls.Add(cap)
        Return card
    End Function

    Private Function BuildCards() As Control
        Dim tbl As New TableLayoutPanel()
        tbl.Dock = DockStyle.Fill
        tbl.ColumnCount = 6
        For i As Integer = 1 To 6
            tbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F / 6.0F))
        Next
        tbl.RowCount = 1
        tbl.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        Dim amber As Color = Color.FromArgb(176, 112, 0)
        If isAdminView Then
            tbl.Controls.Add(StatCard("rev", "Revenue today (net of refunds)", UITheme.SuccessGreen), 0, 0)
            tbl.Controls.Add(StatCard("pays", "Payments today", UITheme.RoyalBlue), 1, 0)
            tbl.Controls.Add(StatCard("reg", "Repairs registered today", UITheme.Navy), 2, 0)
            tbl.Controls.Add(StatCard("pend", "Pending repairs", amber), 3, 0)
            tbl.Controls.Add(StatCard("over", "Overdue repairs", UITheme.DangerRed), 4, 0)
            tbl.Controls.Add(StatCard("emps", "Active employees", UITheme.Navy), 5, 0)
        Else
            tbl.Controls.Add(StatCard("reg", "Repairs registered today", UITheme.RoyalBlue), 0, 0)
            tbl.Controls.Add(StatCard("upd", "Repairs updated today", UITheme.Navy), 1, 0)
            tbl.Controls.Add(StatCard("done", "Completed today", UITheme.SuccessGreen), 2, 0)
            tbl.Controls.Add(StatCard("pend", "Pending repairs", amber), 3, 0)
            tbl.Controls.Add(StatCard("pays", "Payments received today", UITheme.RoyalBlue), 4, 0)
            tbl.Controls.Add(StatCard("rev", "Revenue collected today (net)", UITheme.SuccessGreen), 5, 0)
        End If
        Return tbl
    End Function

    Private Function BuildBanner() As Control
        lblBanner = New Label()
        lblBanner.Dock = DockStyle.Fill
        lblBanner.Margin = New Padding(4)
        lblBanner.AutoSize = False
        lblBanner.AutoEllipsis = True
        lblBanner.TextAlign = ContentAlignment.MiddleLeft
        lblBanner.Padding = New Padding(14, 0, 8, 0)
        lblBanner.Font = UITheme.FontBold(10.5F)
        lblBanner.ForeColor = Color.White
        lblBanner.BackColor = UITheme.RoyalBlue
        Return lblBanner
    End Function

    Private Function NewGrid() As DataGridView
        Dim g As New DataGridView()
        g.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(g)
        g.RowTemplate.Height = 34
        g.ColumnHeadersHeight = 38
        AddHandler g.CellFormatting, AddressOf GridCellFormatting
        Return g
    End Function

    Private Function SectionCard(title As String, content As Control,
                                 Optional topExtra As Control = Nothing) As Panel
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Margin = New Padding(4)
        card.Padding = New Padding(10, 4, 10, 10)

        Dim lbl As New Label()
        lbl.Dock = DockStyle.Top
        lbl.Height = 38
        lbl.AutoSize = False
        lbl.TextAlign = ContentAlignment.MiddleLeft
        lbl.Font = UITheme.FontBold(11.5F)
        lbl.ForeColor = UITheme.Navy
        lbl.Text = title

        content.Dock = DockStyle.Fill
        card.Controls.Add(content)
        If topExtra IsNot Nothing Then
            topExtra.Dock = DockStyle.Top
            card.Controls.Add(topExtra)
        End If
        card.Controls.Add(lbl)
        Return card
    End Function

    Private Function BuildBody() As Control
        Dim body As New TableLayoutPanel()
        body.Dock = DockStyle.Fill
        body.ColumnCount = 2
        body.RowCount = 2
        body.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 46.0F))
        body.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 54.0F))
        body.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
        body.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))

        gridActivity = NewGrid()
        gridAttention = NewGrid()
        gridSecond = NewGrid()

        UIHelpers.AddTextColumn(gridActivity, "WhenText", "When", 70)
        If isAdminView Then UIHelpers.AddTextColumn(gridActivity, "EmployeeName", "Employee", 80)
        UIHelpers.AddTextColumn(gridActivity, "Description", "Activity", 240)
        UIHelpers.AddTextColumn(gridActivity, "Result", "", 50)

        Dim activityCard As Panel = SectionCard(
            If(isAdminView, "Recent activity (all employees)", "My recent activity"), gridActivity)
        body.Controls.Add(activityCard, 0, 0)
        body.SetRowSpan(activityCard, 2)

        If isAdminView Then
            UIHelpers.AddTextColumn(gridAttention, "Employee", "Employee", 110)
            UIHelpers.AddTextColumn(gridAttention, "Registered", "Repairs", 60, True)
            UIHelpers.AddTextColumn(gridAttention, "Payments", "Payments", 65, True)
            UIHelpers.AddTextColumn(gridAttention, "Collected", "Collected", 80, True, "N2")
            UIHelpers.AddTextColumn(gridAttention, "Refunded", "Refunded", 75, True, "N2")
            UIHelpers.AddTextColumn(gridAttention, "Net", "Net", 80, True, "N2")
            body.Controls.Add(SectionCard("Employee performance today", gridAttention), 1, 0)

            UIHelpers.AddTextColumn(gridSecond, "Category", "Reminder", 105)
            UIHelpers.AddTextColumn(gridSecond, "RepairCode", "Repair", 70)
            UIHelpers.AddTextColumn(gridSecond, "CustomerName", "Customer", 110)
            UIHelpers.AddTextColumn(gridSecond, "DueDate", "Expected", 85)
            UIHelpers.AddTextColumn(gridSecond, "Employee", "Employee", 95)
            body.Controls.Add(SectionCard("Collection reminders (whole shop)", gridSecond), 1, 1)
        Else
            UIHelpers.AddTextColumn(gridAttention, "Category", "Category", 110)
            UIHelpers.AddTextColumn(gridAttention, "RepairCode", "Repair", 70)
            UIHelpers.AddTextColumn(gridAttention, "CustomerName", "Customer", 115)
            UIHelpers.AddTextColumn(gridAttention, "DueDate", "Expected", 85)
            UIHelpers.AddTextColumn(gridAttention, "Status", "Status", 85)
            body.Controls.Add(SectionCard("My repairs needing attention", gridAttention), 1, 0)

            lblPaySummary = New Label()
            lblPaySummary.Height = 56
            lblPaySummary.AutoSize = False
            lblPaySummary.TextAlign = ContentAlignment.MiddleLeft
            lblPaySummary.Font = UITheme.FontBold(10)
            lblPaySummary.ForeColor = UITheme.TextDark

            UIHelpers.AddTextColumn(gridSecond, "WhenText", "When", 70)
            UIHelpers.AddTextColumn(gridSecond, "RepairCode", "Repair", 70)
            UIHelpers.AddTextColumn(gridSecond, "Amount", "Amount", 80, True, "N2")
            UIHelpers.AddTextColumn(gridSecond, "PaymentMethod", "Method", 90)
            UIHelpers.AddTextColumn(gridSecond, "ReceiptNumber", "Receipt", 75)
            body.Controls.Add(SectionCard("Payment overview", gridSecond, lblPaySummary), 1, 1)
        End If
        Return body
    End Function

    ' ---------- Data ----------

    Private Sub LoadData()
        Try
            If isAdminView Then LoadAdmin() Else LoadEmployee()
        Catch ex As Exception
            lblBanner.BackColor = UITheme.DangerRed
            lblBanner.Text = "The dashboard could not be refreshed: " & ex.Message
        End Try
    End Sub

    Private Sub LoadEmployee()
        Dim s As EmployeeStats = DashboardService.GetEmployeeStats()
        stats("reg").Text = s.RegisteredToday.ToString()
        stats("upd").Text = s.UpdatedToday.ToString()
        stats("done").Text = s.CompletedToday.ToString()
        stats("pend").Text = s.Pending.ToString()
        stats("pays").Text = s.PaymentsToday.ToString()
        stats("rev").Text = UIHelpers.Money(s.NetToday)

        lblPaySummary.Text = "Today: " & s.PaymentsToday & " payment(s)   |   Collected " &
                             UIHelpers.Money(s.CollectedToday) & "   |   Refunded " &
                             UIHelpers.Money(s.RefundedToday) & "   |   Net " &
                             UIHelpers.Money(s.NetToday)

        gridActivity.DataSource = DashboardService.GetRecentActivity(15)
        gridAttention.DataSource = ReminderService.GetReminders("All", True, True)
        gridSecond.DataSource = DashboardService.GetRecentPayments(10)
        ShowBanner()
    End Sub

    Private Sub LoadAdmin()
        Dim a As AdminStats = DashboardService.GetAdminStats()
        stats("rev").Text = UIHelpers.Money(a.RevenueToday)
        stats("pays").Text = a.PaymentsToday.ToString()
        stats("reg").Text = a.RegisteredToday.ToString()
        stats("pend").Text = a.Pending.ToString()
        stats("emps").Text = a.ActiveEmployees.ToString()
        stats("over").Text = ReminderService.GetCounts()(0).ToString()

        gridActivity.DataSource = DashboardService.GetRecentActivity(15)
        gridAttention.DataSource = DashboardService.GetEmployeeRevenueToday()
        gridSecond.DataSource = ReminderService.GetReminders("All", False, False)
        ShowBanner()
    End Sub

    Private Sub ShowBanner()
        Dim c As Integer() = ReminderService.GetCounts()   ' overdue, due tomorrow, awaiting
        If c(0) = 0 AndAlso c(1) = 0 AndAlso c(2) = 0 Then
            lblBanner.BackColor = UITheme.SuccessGreen
            lblBanner.Text = "No collection reminders right now."
            Return
        End If
        lblBanner.BackColor = If(c(0) > 0, UITheme.DangerRed, Color.FromArgb(176, 112, 0))
        lblBanner.Text = If(isAdminView, "Collection reminders (whole shop):   ", "My collection reminders:   ") &
                         c(0) & " overdue   |   " &
        c(1) & " due tomorrow   |   " & c(2) & " awaiting collection" &
                         If(isAdminView, "", "   -   open Collection Reminders for details")
    End Sub

    Private Sub GridCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 OrElse e.Value Is Nothing OrElse IsDBNull(e.Value) Then Return
        Dim g As DataGridView = DirectCast(sender, DataGridView)
        Dim name As String = g.Columns(e.ColumnIndex).DataPropertyName

        If name = "Result" AndAlso e.Value.ToString() = "Failed" Then
            e.CellStyle.ForeColor = UITheme.DangerRed
            e.CellStyle.Font = boldFont
        ElseIf name = "Category" Then
            Select Case e.Value.ToString()
                Case "Overdue"
                    e.CellStyle.ForeColor = UITheme.DangerRed
                Case "Due tomorrow"
                    e.CellStyle.ForeColor = Color.FromArgb(176, 112, 0)
                Case "Awaiting collection"
                    e.CellStyle.ForeColor = UITheme.SuccessGreen
                Case Else
                    e.CellStyle.ForeColor = UITheme.TextMuted
            End Select
            e.CellStyle.Font = boldFont
        End If
    End Sub

    ' ---------- Clock and auto refresh ----------

    Private Sub UpdateClock()
        If Not SessionManager.IsLoggedIn Then Return
        lblDateTime.Text = DateTime.Now.ToString("ddd, dd MMM yyyy   HH:mm:ss")
        Dim span As TimeSpan = DateTime.Now - SessionManager.LoginTime
        lblLoginStatus.Text = "Online since " & SessionManager.LoginTime.ToString("HH:mm") &
                              "  (" & CInt(Math.Floor(span.TotalHours)) & "h " & span.Minutes & "m)"
    End Sub

    Private Sub OnTick(sender As Object, e As EventArgs)
        If Not SessionManager.IsLoggedIn Then
            refreshTimer.Stop()
            Return
        End If
        UpdateClock()
        tickCount += 1
        If tickCount >= 20 Then
            tickCount = 0
            LoadData()
        End If
    End Sub

End Class