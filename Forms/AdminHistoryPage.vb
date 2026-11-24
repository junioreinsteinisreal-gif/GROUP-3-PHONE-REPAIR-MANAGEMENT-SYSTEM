''' <summary>Administrator audit trail and deleted repair records.</summary>
Public Class AdminHistoryPage
    Inherits UserControl

    Private txtSearch As TextBox
    Private cmbType As ComboBox
    Private dtpFrom As DateTimePicker
    Private dtpTo As DateTimePicker
    Private gridEvents As DataGridView
    Private txtDetail As TextBox
    Private lblEventCount As Label
    Private txtDelSearch As TextBox
    Private gridDeleted As DataGridView
    Private lblDelCount As Label

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsAdmin Then
            Throw New UnauthorizedAccessException("Only administrators can view the administrator history.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)
        BuildPage()
        LoadTypes()
        LoadEvents()
        LoadDeleted()
    End Sub

    Private Sub BuildPage()
        Dim tabs As New TabControl()
        tabs.Dock = DockStyle.Fill
        tabs.Font = UITheme.FontBold(10.5F)
        tabs.SizeMode = TabSizeMode.Fixed
        tabs.ItemSize = New Size(260, 42)

        Dim t1 As New TabPage("Audit Events")
        t1.BackColor = UITheme.LightGrey
        t1.Controls.Add(BuildEventsTab())
        Dim t2 As New TabPage("Deleted Repairs")
        t2.BackColor = UITheme.LightGrey
        t2.Controls.Add(BuildDeletedTab())

        tabs.TabPages.Add(t1)
        tabs.TabPages.Add(t2)
        Me.Controls.Add(tabs)
    End Sub

    Private Function Cap(text As String) As Label
        Dim l As Label = UITheme.MakeLabel(text, 10, True, UITheme.TextDark)
        l.Margin = New Padding(10, 13, 2, 0)
        Return l
    End Function

    Private Function MakePicker() As DateTimePicker
        Dim d As New DateTimePicker()
        d.Font = UITheme.FontRegular(10.5F)
        d.Format = DateTimePickerFormat.Custom
        d.CustomFormat = "dd MMM yyyy"
        d.ShowCheckBox = True
        d.Checked = False
        d.Width = 170
        d.Margin = New Padding(3, 9, 3, 3)
        Return d
    End Function

    Private Function Btn(text As String, color As Color, width As Integer,
                         handler As EventHandler) As Button
        Dim b As Button = UITheme.MakeButton(text, color, width, 38)
        b.Margin = New Padding(6, 5, 0, 3)
        AddHandler b.Click, handler
        Return b
    End Function

    ' ---------- Audit events tab (Fill first, then Bottom, then Top) ----------

    Private Function BuildEventsTab() As Panel
        Dim pnl As New Panel()
        pnl.Dock = DockStyle.Fill

        gridEvents = New DataGridView()
        gridEvents.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(gridEvents)
        UIHelpers.AddTextColumn(gridEvents, "EventAt", "When", 120)
        UIHelpers.AddTextColumn(gridEvents, "EventType", "Event", 125)
        UIHelpers.AddTextColumn(gridEvents, "PerformedByName", "By", 100)
        UIHelpers.AddTextColumn(gridEvents, "PerformedByRole", "Role", 85)
        UIHelpers.AddTextColumn(gridEvents, "RepairCode", "Repair", 70)
        UIHelpers.AddTextColumn(gridEvents, "Details", "Details", 280)
        UIHelpers.AddTextColumn(gridEvents, "OriginalData", "Original Record", 220)
        AddHandler gridEvents.SelectionChanged, AddressOf ShowSelectedEvent

        txtDetail = UITheme.MakeTextBox(100, 11)
        txtDetail.Multiline = True
        txtDetail.ReadOnly = True
        txtDetail.ScrollBars = ScrollBars.Vertical
        txtDetail.BackColor = UITheme.CardWhite
        txtDetail.Dock = DockStyle.Fill

        Dim detailTitle As New Label()
        detailTitle.Dock = DockStyle.Top
        detailTitle.Height = 34
        detailTitle.AutoSize = False
        detailTitle.TextAlign = ContentAlignment.MiddleLeft
        detailTitle.Font = UITheme.FontBold(11)
        detailTitle.ForeColor = UITheme.Navy
        detailTitle.Text = "Selected event: details and original record"

        Dim detailPanel As New Panel()
        detailPanel.Dock = DockStyle.Bottom
        detailPanel.Height = 190
        detailPanel.BackColor = UITheme.CardWhite
        detailPanel.Padding = New Padding(12, 4, 12, 10)
        detailPanel.Controls.Add(txtDetail)
        detailPanel.Controls.Add(detailTitle)

        txtSearch = UITheme.MakeTextBox(240, 11)
        txtSearch.MaxLength = 60
        txtSearch.PlaceholderText = "Name, repair code or text"
        txtSearch.Margin = New Padding(5, 9, 5, 3)
        AddHandler txtSearch.KeyDown,
            Sub(s, ev)
                If ev.KeyCode = Keys.Enter Then
                    ev.SuppressKeyPress = True
                    LoadEvents()
                End If
            End Sub

        cmbType = New ComboBox()
        cmbType.Font = UITheme.FontRegular(10.5F)
        cmbType.DropDownStyle = ComboBoxStyle.DropDownList
        cmbType.Width = 200
        cmbType.Margin = New Padding(3, 9, 3, 3)

        dtpFrom = MakePicker()
        dtpTo = MakePicker()

        lblEventCount = UITheme.MakeLabel("", 10, True, UITheme.TextMuted)
        lblEventCount.Margin = New Padding(14, 14, 0, 0)

        Dim toolbar As New FlowLayoutPanel()
        toolbar.Dock = DockStyle.Top
        toolbar.AutoSize = True
        toolbar.AutoSizeMode = AutoSizeMode.GrowAndShrink
        toolbar.Padding = New Padding(0, 0, 0, 4)
        toolbar.Controls.Add(txtSearch)
        toolbar.Controls.Add(Cap("Event:"))
        toolbar.Controls.Add(cmbType)
        toolbar.Controls.Add(Cap("From:"))
        toolbar.Controls.Add(dtpFrom)
        toolbar.Controls.Add(Cap("To:"))
        toolbar.Controls.Add(dtpTo)
        toolbar.Controls.Add(Btn("Search", UITheme.RoyalBlue, 110, Sub(s, ev) LoadEvents()))
        toolbar.Controls.Add(Btn("Clear", UITheme.TextMuted, 90, AddressOf ClearEventFilters))
        toolbar.Controls.Add(Btn("Refresh", UITheme.Navy, 100,
                                 Sub(s, ev)
                                     LoadTypes()
                                     LoadEvents()
                                     LoadDeleted()
                                 End Sub))
        toolbar.Controls.Add(Btn("Export to Excel", UITheme.SuccessGreen, 170,
                                 Sub(s, ev) ExcelExporter.ExportGrid(gridEvents, "AdministratorHistory", "Audit Events")))
        toolbar.Controls.Add(lblEventCount)

        pnl.Controls.Add(gridEvents)
        pnl.Controls.Add(detailPanel)
        pnl.Controls.Add(toolbar)
        Return pnl
    End Function

    Private Sub LoadTypes()
        Dim keep As String = cmbType.Text
        cmbType.Items.Clear()
        cmbType.Items.Add("All events")
        Try
            For Each t As String In AdminHistoryService.GetEventTypes()
                cmbType.Items.Add(t)
            Next
        Catch
        End Try
        Dim i As Integer = cmbType.Items.IndexOf(keep)
        cmbType.SelectedIndex = If(i >= 0, i, 0)
    End Sub

    Private Sub LoadEvents()
        Dim fromD As Date? = Nothing
        Dim toD As Date? = Nothing
        If dtpFrom.Checked Then fromD = dtpFrom.Value.Date
        If dtpTo.Checked Then toD = dtpTo.Value.Date
        If fromD.HasValue AndAlso toD.HasValue AndAlso fromD.Value > toD.Value Then
            MessageBox.Show("The start date cannot be after the end date.", "Filters",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Dim typeFilter As String = If(cmbType.SelectedIndex > 0, cmbType.Text, "")
            Dim dt As DataTable = AdminHistoryService.GetEvents(txtSearch.Text, typeFilter, fromD, toD)
            gridEvents.DataSource = dt
            lblEventCount.Text = dt.Rows.Count & " event(s)" &
                If(dt.Rows.Count >= AdminHistoryService.MaxRows, " (limit reached, narrow the filters)", "")
            ShowSelectedEvent(Nothing, EventArgs.Empty)
        Catch ex As Exception
            MessageBox.Show("The audit events could not be loaded: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ClearEventFilters(sender As Object, e As EventArgs)
        txtSearch.Clear()
        cmbType.SelectedIndex = 0
        dtpFrom.Checked = False
        dtpTo.Checked = False
        LoadEvents()
    End Sub

    Private Sub ShowSelectedEvent(sender As Object, e As EventArgs)
        If gridEvents.CurrentRow Is Nothing Then
            txtDetail.Clear()
            Return
        End If
        Dim drv As DataRowView = TryCast(gridEvents.CurrentRow.DataBoundItem, DataRowView)
        If drv Is Nothing Then
            txtDetail.Clear()
            Return
        End If
        Dim original As String = drv("OriginalData").ToString()
        txtDetail.Text = drv("EventAt").ToString() & "   |   " & drv("EventType").ToString() & vbCrLf &
                         drv("Details").ToString() & vbCrLf & vbCrLf &
                         "Original record:" & vbCrLf & If(original = "", "(none)", original)
    End Sub

    ' ---------- Deleted repairs tab ----------

    Private Function BuildDeletedTab() As Panel
        Dim pnl As New Panel()
        pnl.Dock = DockStyle.Fill

        gridDeleted = New DataGridView()
        gridDeleted.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(gridDeleted)
        UIHelpers.AddTextColumn(gridDeleted, "DeletedAt", "Deleted At", 115)
        UIHelpers.AddTextColumn(gridDeleted, "RepairCode", "Repair Code", 80)
        UIHelpers.AddTextColumn(gridDeleted, "CustomerName", "Customer", 110)
        UIHelpers.AddTextColumn(gridDeleted, "CustomerPhone", "Phone", 100)
        UIHelpers.AddTextColumn(gridDeleted, "Device", "Device", 100)
        UIHelpers.AddTextColumn(gridDeleted, "ProblemDescription", "Problem", 150)
        UIHelpers.AddTextColumn(gridDeleted, "EstimatedCost", "Estimated", 80, True, "N2")
        UIHelpers.AddTextColumn(gridDeleted, "Status", "Status", 80)
        UIHelpers.AddTextColumn(gridDeleted, "TotalPaid", "Paid", 70, True, "N2")
        UIHelpers.AddTextColumn(gridDeleted, "RegisteredBy", "Registered By", 100)
        UIHelpers.AddTextColumn(gridDeleted, "DeletedByName", "Deleted By", 100)
        UIHelpers.AddTextColumn(gridDeleted, "Reason", "Reason", 160)

        lblDelCount = UITheme.MakeLabel("", 10, True, UITheme.TextMuted)
        lblDelCount.Margin = New Padding(14, 14, 0, 0)

        txtDelSearch = UITheme.MakeTextBox(280, 11)
        txtDelSearch.MaxLength = 60
        txtDelSearch.PlaceholderText = "Code, customer, phone, employee or reason"
        txtDelSearch.Margin = New Padding(5, 9, 5, 3)
        AddHandler txtDelSearch.KeyDown,
            Sub(s, ev)
                If ev.KeyCode = Keys.Enter Then
                    ev.SuppressKeyPress = True
                    LoadDeleted()
                End If
            End Sub

        Dim toolbar As New FlowLayoutPanel()
        toolbar.Dock = DockStyle.Top
        toolbar.AutoSize = True
        toolbar.AutoSizeMode = AutoSizeMode.GrowAndShrink
        toolbar.Padding = New Padding(0, 0, 0, 4)
        toolbar.Controls.Add(txtDelSearch)
        toolbar.Controls.Add(Btn("Search", UITheme.RoyalBlue, 110, Sub(s, ev) LoadDeleted()))
        toolbar.Controls.Add(Btn("Clear", UITheme.TextMuted, 90,
                                 Sub(s, ev)
                                     txtDelSearch.Clear()
                                     LoadDeleted()
                                 End Sub))
        toolbar.Controls.Add(Btn("Export to Excel", UITheme.SuccessGreen, 170,
                                 Sub(s, ev) ExcelExporter.ExportGrid(gridDeleted, "DeletedRepairs", "Deleted Repairs")))
        toolbar.Controls.Add(lblDelCount)

        pnl.Controls.Add(gridDeleted)
        pnl.Controls.Add(toolbar)
        Return pnl
    End Function

    Private Sub LoadDeleted()
        Try
            Dim dt As DataTable = AdminHistoryService.GetDeletedRepairs(txtDelSearch.Text)
            gridDeleted.DataSource = dt
            lblDelCount.Text = dt.Rows.Count & " deleted repair(s)"
        Catch ex As Exception
            MessageBox.Show("The deleted repairs could not be loaded: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class