''' <summary>Repairs table with search, status filter and details.</summary>
Public Class RepairsPage
    Inherits UserControl

    Private txtSearch As TextBox
    Private cmbStatus As ComboBox
    Private grid As DataGridView
    Private lblCount As Label
    Private suspendEvents As Boolean = False
    Private ReadOnly boldFont As Font = UITheme.FontBold(10)

    Public Sub New()
        SessionManager.RequireLogin()

        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)

        BuildPage()
        LoadData()
        ActivityLogger.Log("ViewRepairs", "Opened the repairs list.")
    End Sub

    ' ---------- Building the page ----------
    ' Docking: grid (Fill) is added first, then footer (Bottom), then toolbar (Top).

    Private Sub BuildPage()
        ' ----- Grid -----
        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        UIHelpers.AddTextColumn(grid, "RepairCode", "Repair Code", 90)
        UIHelpers.AddTextColumn(grid, "CustomerName", "Customer", 130)
        UIHelpers.AddTextColumn(grid, "CustomerPhone", "Phone", 110)
        UIHelpers.AddTextColumn(grid, "Brand", "Brand", 80)
        UIHelpers.AddTextColumn(grid, "Model", "Model", 100)
        UIHelpers.AddTextColumn(grid, "Status", "Status", 95)
        UIHelpers.AddTextColumn(grid, "EstimatedCost", "Estimated", 90, True, "N2")
        UIHelpers.AddTextColumn(grid, "Paid", "Paid", 80, True, "N2")
        UIHelpers.AddTextColumn(grid, "Balance", "Balance", 80, True, "N2")
        UIHelpers.AddTextColumn(grid, "DateReceived", "Received", 90)
        UIHelpers.AddTextColumn(grid, "ExpectedCollectionDate", "Expected Collection", 105)
        UIHelpers.AddTextColumn(grid, "Employee", "Employee", 100)
        AddHandler grid.CellFormatting, AddressOf GridCellFormatting
        AddHandler grid.CellDoubleClick, AddressOf GridDoubleClick

        ' ----- Footer -----
        Dim footer As New Panel()
        footer.Dock = DockStyle.Bottom
        footer.Height = 34
        lblCount = UITheme.MakeLabel("", 10, False, UITheme.TextMuted)
        lblCount.Left = 5
        lblCount.Top = 9
        footer.Controls.Add(lblCount)

        ' ----- Toolbar -----
        Dim toolbar As New FlowLayoutPanel()
        toolbar.Dock = DockStyle.Top
        toolbar.AutoSize = True
        toolbar.AutoSizeMode = AutoSizeMode.GrowAndShrink
        toolbar.WrapContents = True
        toolbar.FlowDirection = FlowDirection.LeftToRight
        toolbar.Padding = New Padding(0, 0, 0, 6)

        txtSearch = UITheme.MakeTextBox(300, 11)
        txtSearch.MaxLength = 60
        txtSearch.PlaceholderText = "Code, customer, phone, brand or model"
        txtSearch.Margin = New Padding(5, 5, 5, 3)
        AddHandler txtSearch.KeyDown, AddressOf SearchKeyDown

        cmbStatus = New ComboBox()
        cmbStatus.Font = UITheme.FontRegular(11)
        cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList
        cmbStatus.Width = 160
        cmbStatus.Margin = New Padding(5, 5, 5, 3)
        cmbStatus.Items.Add("All")
        cmbStatus.Items.AddRange(RepairService.RepairStatuses)
        cmbStatus.SelectedIndex = 0
        AddHandler cmbStatus.SelectedIndexChanged,
            Sub(s, ev)
                If Not suspendEvents Then LoadData()
            End Sub

        Dim btnSearch As Button = UITheme.MakeButton("Search", UITheme.RoyalBlue, 100, 34)
        AddHandler btnSearch.Click, Sub(s, ev) RunSearch()

        Dim btnClear As Button = UITheme.MakeButton("Clear", UITheme.TextMuted, 100, 34)
        AddHandler btnClear.Click, Sub(s, ev) ClearFilters()

        Dim btnRefresh As Button = UITheme.MakeButton("Refresh", UITheme.Navy, 100, 34)
        AddHandler btnRefresh.Click, Sub(s, ev) LoadData()

        Dim btnDetails As Button = UITheme.MakeButton("View Details", UITheme.SuccessGreen, 130, 34)
        AddHandler btnDetails.Click, Sub(s, ev) ShowDetails()

        toolbar.Controls.AddRange(New Control() {txtSearch, cmbStatus, btnSearch, btnClear,
                                                 btnRefresh, btnDetails})

        Me.Controls.Add(grid)
        Me.Controls.Add(footer)
        Me.Controls.Add(toolbar)
    End Sub

    ' ---------- Data ----------

    Private Sub LoadData()
        Try
            Dim dt As DataTable = RepairService.GetRepairs(txtSearch.Text, cmbStatus.Text)
            grid.DataSource = dt
            If dt.Rows.Count = 0 Then
                lblCount.Text = "No repairs found."
            ElseIf dt.Rows.Count = 1 Then
                lblCount.Text = "1 repair shown"
            Else
                lblCount.Text = dt.Rows.Count & " repairs shown"
            End If
        Catch ex As Exception
            MessageBox.Show("The repairs could not be loaded: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RunSearch()
        LoadData()
        Dim q As String = txtSearch.Text.Trim()
        If q <> "" Then
            ActivityLogger.Log("SearchRepairs", "Searched repairs for '" & q & "'.")
        End If
    End Sub

    Private Sub ClearFilters()
        suspendEvents = True
        txtSearch.Clear()
        cmbStatus.SelectedIndex = 0
        suspendEvents = False
        LoadData()
        txtSearch.Focus()
    End Sub

    Private Sub SearchKeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RunSearch()
        End If
    End Sub

    ' ---------- Grid behaviour ----------

    Private Sub GridCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 OrElse e.Value Is Nothing OrElse IsDBNull(e.Value) Then Return
        Dim colName As String = grid.Columns(e.ColumnIndex).DataPropertyName

        If colName = "Status" Then
            Select Case e.Value.ToString()
                Case "Completed"
                    e.CellStyle.ForeColor = UITheme.SuccessGreen
                Case "Collected"
                    e.CellStyle.ForeColor = UITheme.TextMuted
                Case "Unrepairable"
                    e.CellStyle.ForeColor = UITheme.DangerRed
                Case "In Progress"
                    e.CellStyle.ForeColor = UITheme.RoyalBlue
                Case Else
                    e.CellStyle.ForeColor = Color.FromArgb(176, 112, 0)
            End Select
            e.CellStyle.Font = boldFont
        ElseIf colName = "Balance" Then
            If Convert.ToDouble(e.Value) > 0.005 Then
                e.CellStyle.ForeColor = UITheme.DangerRed
            Else
                e.CellStyle.ForeColor = UITheme.SuccessGreen
            End If
            e.CellStyle.Font = boldFont
        End If
    End Sub

    Private Sub GridDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then ShowDetails()
    End Sub

    Private Sub ShowDetails()
        If grid.CurrentRow Is Nothing Then
            MessageBox.Show("Select a repair in the list first.", "View Details",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim drv As DataRowView = TryCast(grid.CurrentRow.DataBoundItem, DataRowView)
        If drv Is Nothing Then Return
        Dim code As String = drv("RepairCode").ToString()

        Dim details As String = RepairService.GetRepairDetails(code)
        If details = "" Then
            MessageBox.Show("That repair could not be found.", "View Details",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ActivityLogger.Log("ViewRepair", "Viewed repair " & code & ".", code)
        MessageBox.Show(details, "Repair " & code,
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class