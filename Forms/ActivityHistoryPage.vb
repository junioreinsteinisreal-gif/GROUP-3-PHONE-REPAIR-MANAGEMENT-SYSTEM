''' <summary>Activity history. Admins see everyone; employees see only their own.</summary>
Public Class ActivityHistoryPage
    Inherits UserControl

    Private ReadOnly isAdminView As Boolean
    Private dtpDate As DateTimePicker
    Private dtpFrom As DateTimePicker
    Private dtpTo As DateTimePicker
    Private cmbDay As ComboBox
    Private cmbType As ComboBox
    Private cmbSort As ComboBox
    Private txtEmpName As TextBox
    Private txtEmpId As TextBox
    Private txtRepair As TextBox
    Private grid As DataGridView
    Private lblCount As Label
    Private ReadOnly boldFont As Font = UITheme.FontBold(10.5F)

    Public Sub New(Optional presetEmployeeId As Integer = 0)
        SessionManager.RequireLogin()
        isAdminView = SessionManager.IsAdmin
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)
        BuildPage()
        If isAdminView AndAlso presetEmployeeId > 0 Then txtEmpId.Text = presetEmployeeId.ToString()
        LoadTypes()
        LoadData()
    End Sub

    ' ---------- Building ----------

    Private Function MakePicker() As DateTimePicker
        Dim d As New DateTimePicker()
        d.Font = UITheme.FontRegular(10.5F)
        d.Format = DateTimePickerFormat.Custom
        d.CustomFormat = "dd MMM yyyy"
        d.ShowCheckBox = True
        d.Checked = False
        Return d
    End Function

    Private Function MakeCombo(ParamArray items() As String) As ComboBox
        Dim c As New ComboBox()
        c.Font = UITheme.FontRegular(10.5F)
        c.DropDownStyle = ComboBoxStyle.DropDownList
        c.Items.AddRange(items)
        c.SelectedIndex = 0
        Return c
    End Function

    Private Function FieldBox(labelText As String, ctrl As Control, width As Integer) As Panel
        Dim pnl As New Panel()
        pnl.Size = New Size(width, 72)
        pnl.Margin = New Padding(0, 0, 12, 4)
        Dim lbl As Label = UITheme.MakeLabel(labelText, 9.5F, True, UITheme.TextDark)
        lbl.Left = 0
        lbl.Top = 0
        ctrl.Left = 0
        ctrl.Top = 28
        ctrl.Width = width
        pnl.Controls.Add(lbl)
        pnl.Controls.Add(ctrl)
        Return pnl
    End Function

    Private Sub DigitsOnly(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then e.Handled = True
    End Sub

    Private Sub SearchOnEnter(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            LoadData()
        End If
    End Sub

    Private Function MakeBox(width As Integer, maxLen As Integer) As TextBox
        Dim t As TextBox = UITheme.MakeTextBox(width, 10.5F)
        t.MaxLength = maxLen
        AddHandler t.KeyDown, AddressOf SearchOnEnter
        Return t
    End Function

    ' Docking: grid (Fill) first, then footer, then the button row, then filters (top-most).
    Private Sub BuildPage()
        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        UIHelpers.AddTextColumn(grid, "ActivityId", "ID", 50)
        UIHelpers.AddTextColumn(grid, "EmployeeId", "Emp. ID", 50)
        UIHelpers.AddTextColumn(grid, "EmployeeName", "Employee", 110)
        UIHelpers.AddTextColumn(grid, "ActionType", "Action", 105)
        UIHelpers.AddTextColumn(grid, "RepairCode", "Repair", 70)
        UIHelpers.AddTextColumn(grid, "Description", "Description", 300)
        UIHelpers.AddTextColumn(grid, "ActivityDate", "Date", 90)
        UIHelpers.AddTextColumn(grid, "DayName", "Day", 85)
        UIHelpers.AddTextColumn(grid, "ActivityTime", "Time", 75)
        UIHelpers.AddTextColumn(grid, "Result", "Result", 75)
        UIHelpers.AddTextColumn(grid, "Removed", "Removed By Employee", 90)
        For Each c As DataGridViewColumn In grid.Columns
            c.SortMode = DataGridViewColumnSortMode.Automatic
        Next
        If Not isAdminView Then
            grid.Columns("EmployeeId").Visible = False
            grid.Columns("Removed").Visible = False
        End If
        AddHandler grid.CellFormatting, AddressOf GridCellFormatting
        AddHandler grid.CellDoubleClick, AddressOf GridDoubleClick

        Dim footer As New Panel()
        footer.Dock = DockStyle.Bottom
        footer.Height = 34
        lblCount = UITheme.MakeLabel("", 10, False, UITheme.TextMuted)
        lblCount.Left = 5
        lblCount.Top = 9
        footer.Controls.Add(lblCount)

        ' ----- Filter controls -----
        dtpDate = MakePicker()
        dtpFrom = MakePicker()
        dtpTo = MakePicker()
        Dim dayItems As New List(Of String)
        dayItems.Add("Any day")
        dayItems.AddRange(ActivityService.DayNames)
        cmbDay = MakeCombo(dayItems.ToArray())
        cmbType = MakeCombo("All types")
        cmbSort = MakeCombo("Newest first", "Oldest first")
        txtEmpName = MakeBox(180, 60)
        txtEmpId = MakeBox(100, 9)
        AddHandler txtEmpId.KeyPress, AddressOf DigitsOnly
        txtRepair = MakeBox(140, 20)

        Dim filters As New FlowLayoutPanel()
        filters.Dock = DockStyle.Top
        filters.AutoSize = True
        filters.AutoSizeMode = AutoSizeMode.GrowAndShrink
        filters.WrapContents = True
        filters.BackColor = UITheme.CardWhite
        filters.Padding = New Padding(12, 8, 12, 4)
        filters.Controls.Add(FieldBox("Specific date (tick to use)", dtpDate, 200))
        filters.Controls.Add(FieldBox("Day of week", cmbDay, 150))
        filters.Controls.Add(FieldBox("From date (tick to use)", dtpFrom, 200))
        filters.Controls.Add(FieldBox("To date (tick to use)", dtpTo, 200))
        If isAdminView Then
            filters.Controls.Add(FieldBox("Employee name", txtEmpName, 180))
            filters.Controls.Add(FieldBox("Employee ID", txtEmpId, 100))
        End If
        filters.Controls.Add(FieldBox("Activity type", cmbType, 190))
        filters.Controls.Add(FieldBox("Repair code", txtRepair, 140))
        filters.Controls.Add(FieldBox("Sort", cmbSort, 150))

        ' ----- Buttons -----
        Dim buttons As New FlowLayoutPanel()
        buttons.Dock = DockStyle.Top
        buttons.AutoSize = True
        buttons.AutoSizeMode = AutoSizeMode.GrowAndShrink
        buttons.Padding = New Padding(0, 6, 0, 6)

        Dim btnSearch As Button = UITheme.MakeButton("Search", UITheme.RoyalBlue, 120, 40)
        AddHandler btnSearch.Click, Sub(s, ev) LoadData()
        Dim btnRefresh As Button = UITheme.MakeButton("Refresh", UITheme.Navy, 120, 40)
        AddHandler btnRefresh.Click,
            Sub(s, ev)
                LoadTypes()
                LoadData()
            End Sub
        Dim btnClear As Button = UITheme.MakeButton("Clear Filters", UITheme.TextMuted, 160, 40)
        AddHandler btnClear.Click, AddressOf ClearFilters
        Dim btnExport As Button = UITheme.MakeButton("Export to Excel", UITheme.SuccessGreen, 190, 40)
        AddHandler btnExport.Click,
            Sub(s, ev)
                If ExcelExporter.ExportGrid(grid, If(isAdminView, "ActivityHistory", "MyActivity"), "Activity") Then
                    ActivityLogger.Log("ExportExcel", "Exported the activity history to Excel.")
                End If
            End Sub

        buttons.Controls.Add(btnSearch)
        buttons.Controls.Add(btnRefresh)
        buttons.Controls.Add(btnClear)
        buttons.Controls.Add(btnExport)
        If Not isAdminView Then
            Dim btnRemove As Button = UITheme.MakeButton("Remove Selected From My History",
                                                         UITheme.DangerRed, 340, 40)
            AddHandler btnRemove.Click, AddressOf RemoveClick
            buttons.Controls.Add(btnRemove)
        End If

        Me.Controls.Add(grid)
        Me.Controls.Add(footer)
        Me.Controls.Add(buttons)
        Me.Controls.Add(filters)
    End Sub

    ' ---------- Filters ----------

    Private Sub LoadTypes()
        Dim keep As String = cmbType.Text
        cmbType.Items.Clear()
        cmbType.Items.Add("All types")
        Try
            For Each t As String In ActivityService.GetActionTypes()
                cmbType.Items.Add(t)
            Next
        Catch
        End Try
        Dim i As Integer = cmbType.Items.IndexOf(keep)
        cmbType.SelectedIndex = If(i >= 0, i, 0)
    End Sub

    Private Function BuildFilter(ByRef err As String) As ActivityFilter
        Dim f As New ActivityFilter()
        If dtpDate.Checked Then f.SpecificDate = dtpDate.Value.Date
        If cmbDay.SelectedIndex > 0 Then f.DayName = cmbDay.Text
        If dtpFrom.Checked Then f.FromDate = dtpFrom.Value.Date
        If dtpTo.Checked Then f.ToDate = dtpTo.Value.Date
        If f.FromDate.HasValue AndAlso f.ToDate.HasValue AndAlso f.FromDate.Value > f.ToDate.Value Then
            err = "The start date cannot be after the end date."
            Return Nothing
        End If
        If isAdminView Then
            f.EmployeeName = txtEmpName.Text.Trim()
            Dim id As Integer
            If txtEmpId.Text.Trim() <> "" Then
                If Not Integer.TryParse(txtEmpId.Text.Trim(), id) Then
                    err = "Enter a valid employee ID."
                    Return Nothing
                End If
                f.EmployeeId = id
            End If
        End If
        If cmbType.SelectedIndex > 0 Then f.ActionType = cmbType.Text
        f.RepairCode = txtRepair.Text.Trim()
        f.NewestFirst = (cmbSort.SelectedIndex = 0)
        Return f
    End Function

    Private Sub ClearFilters(sender As Object, e As EventArgs)
        dtpDate.Checked = False
        dtpFrom.Checked = False
        dtpTo.Checked = False
        cmbDay.SelectedIndex = 0
        cmbType.SelectedIndex = 0
        cmbSort.SelectedIndex = 0
        txtEmpName.Clear()
        txtEmpId.Clear()
        txtRepair.Clear()
        LoadData()
    End Sub

    Private Sub LoadData()
        Dim err As String = ""
        Dim f As ActivityFilter = BuildFilter(err)
        If f Is Nothing Then
            MessageBox.Show(err, "Filters", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Dim dt As DataTable = ActivityService.GetActivities(f)
            grid.DataSource = dt
            If dt.Rows.Count = 0 Then
                lblCount.Text = "No activities match these filters."
            Else
                lblCount.Text = dt.Rows.Count & " activities shown" &
                    If(dt.Rows.Count >= ActivityService.MaxRows, " (limit reached, narrow the filters)", "") &
                    ".  Double-click a row to read the full description."
            End If
        Catch ex As Exception
            MessageBox.Show("The activity history could not be loaded: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ---------- Grid ----------

    Private Sub GridCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 OrElse e.Value Is Nothing OrElse IsDBNull(e.Value) Then Return
        Dim name As String = grid.Columns(e.ColumnIndex).DataPropertyName
        If name = "Result" AndAlso e.Value.ToString() = "Failed" Then
            e.CellStyle.ForeColor = UITheme.DangerRed
            e.CellStyle.Font = boldFont
        ElseIf name = "Removed" AndAlso e.Value.ToString() = "Yes" Then
            e.CellStyle.ForeColor = Color.FromArgb(176, 112, 0)
            e.CellStyle.Font = boldFont
        End If
    End Sub

    Private Function SelectedRow() As DataRowView
        If grid.CurrentRow Is Nothing Then Return Nothing
        Return TryCast(grid.CurrentRow.DataBoundItem, DataRowView)
    End Function

    Private Sub GridDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        Dim drv As DataRowView = SelectedRow()
        If drv Is Nothing Then Return
        MessageBox.Show(drv("Description").ToString() & vbCrLf & vbCrLf &
                        drv("ActivityDate").ToString() & " " & drv("ActivityTime").ToString(),
                        "Activity #" & drv("ActivityId").ToString(),
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub RemoveClick(sender As Object, e As EventArgs)
        Dim drv As DataRowView = SelectedRow()
        If drv Is Nothing Then
            MessageBox.Show("Select an entry in the list first.", "My Activity",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        If MessageBox.Show("Remove this entry from your history?" & vbCrLf & vbCrLf &
                           "The administrator will still keep a permanent record of it.",
                           "Remove Entry", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question) <> DialogResult.Yes Then Return

        Dim err As String = ""
        Try
            err = ActivityService.RemoveFromMyHistory(Convert.ToInt64(drv("ActivityId")))
        Catch ex As UnauthorizedAccessException
            err = ex.Message
        End Try
        If err <> "" Then
            MessageBox.Show(err, "My Activity", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        LoadTypes()
        LoadData()
    End Sub

End Class