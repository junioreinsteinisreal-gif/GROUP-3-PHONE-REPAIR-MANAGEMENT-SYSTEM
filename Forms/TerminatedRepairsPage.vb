''' <summary>Administrator view of repairs that employees terminated as beyond repair.</summary>
Public Class TerminatedRepairsPage
    Inherits UserControl

    Private txtSearch As TextBox
    Private grid As DataGridView
    Private txtReason As TextBox
    Private lblCount As Label
    Private suspendEvents As Boolean = False

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsAdmin Then
            Throw New UnauthorizedAccessException("Only administrators can view terminated repairs.")
        End If

        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)

        BuildPage()
        LoadData()
    End Sub

    ' Docking: grid (Fill) first, then the reason panel (Bottom), then the toolbar (Top).
    Private Sub BuildPage()
        ' ----- Grid -----
        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        UIHelpers.AddTextColumn(grid, "TerminatedAt", "Terminated At", 115)
        UIHelpers.AddTextColumn(grid, "RepairCode", "Repair Code", 80)
        UIHelpers.AddTextColumn(grid, "CustomerName", "Customer", 120)
        UIHelpers.AddTextColumn(grid, "CustomerPhone", "Phone", 105)
        UIHelpers.AddTextColumn(grid, "Device", "Device", 115)
        UIHelpers.AddTextColumn(grid, "EstimatedCost", "Estimated", 85, True, "N2")
        UIHelpers.AddTextColumn(grid, "AmountPaid", "Paid", 80, True, "N2")
        UIHelpers.AddTextColumn(grid, "Refunded", "Refunded", 80, True, "N2")
        UIHelpers.AddTextColumn(grid, "PreviousStatus", "Status Before", 95)
        UIHelpers.AddTextColumn(grid, "TerminatedByName", "Terminated By", 110)
        UIHelpers.AddTextColumn(grid, "Reason", "Reason", 220)
        AddHandler grid.SelectionChanged, AddressOf ShowSelectedReason

        ' ----- Reason panel -----
        txtReason = UITheme.MakeTextBox(100, 11)
        txtReason.Multiline = True
        txtReason.ReadOnly = True
        txtReason.ScrollBars = ScrollBars.Vertical
        txtReason.BackColor = UITheme.CardWhite
        txtReason.Dock = DockStyle.Fill

        Dim lblReasonTitle As New Label()
        lblReasonTitle.AutoSize = False
        lblReasonTitle.Dock = DockStyle.Top
        lblReasonTitle.Height = 38
        lblReasonTitle.TextAlign = ContentAlignment.MiddleLeft
        lblReasonTitle.Font = UITheme.FontBold(11)
        lblReasonTitle.ForeColor = UITheme.Navy
        lblReasonTitle.Text = "Full reason given by the employee (select a row above)"

        Dim detail As New Panel()
        detail.Dock = DockStyle.Bottom
        detail.Height = 200
        detail.BackColor = UITheme.CardWhite
        detail.Padding = New Padding(15, 8, 15, 12)
        detail.Controls.Add(txtReason)
        detail.Controls.Add(lblReasonTitle)

        ' ----- Toolbar -----
        Dim toolbar As New FlowLayoutPanel()
        toolbar.Dock = DockStyle.Top
        toolbar.AutoSize = True
        toolbar.AutoSizeMode = AutoSizeMode.GrowAndShrink
        toolbar.WrapContents = True
        toolbar.FlowDirection = FlowDirection.LeftToRight
        toolbar.Padding = New Padding(0, 0, 0, 6)

        txtSearch = UITheme.MakeTextBox(320, 11)
        txtSearch.MaxLength = 60
        txtSearch.PlaceholderText = "Code, customer, phone, employee or reason"
        txtSearch.Margin = New Padding(5, 5, 5, 3)
        AddHandler txtSearch.KeyDown, AddressOf SearchKeyDown

        Dim btnSearch As Button = UITheme.MakeButton("Search", UITheme.RoyalBlue, 100, 34)
        AddHandler btnSearch.Click, Sub(s, ev) LoadData()

        Dim btnClear As Button = UITheme.MakeButton("Clear", UITheme.TextMuted, 100, 34)
        AddHandler btnClear.Click,
            Sub(s, ev)
                txtSearch.Clear()
                LoadData()
                txtSearch.Focus()
            End Sub

        Dim btnRefresh As Button = UITheme.MakeButton("Refresh", UITheme.Navy, 100, 34)
        AddHandler btnRefresh.Click, Sub(s, ev) LoadData()

        lblCount = UITheme.MakeLabel("", 10, True, UITheme.TextMuted)
        lblCount.Margin = New Padding(15, 12, 0, 0)

        toolbar.Controls.AddRange(New Control() {txtSearch, btnSearch, btnClear, btnRefresh, lblCount})

        Me.Controls.Add(grid)
        Me.Controls.Add(detail)
        Me.Controls.Add(toolbar)
    End Sub

    Private Sub LoadData()
        Try
            Dim dt As DataTable = TerminationService.GetTerminations(txtSearch.Text)
            grid.DataSource = dt
            If dt.Rows.Count = 0 Then
                lblCount.Text = "No terminated repairs found."
            ElseIf dt.Rows.Count = 1 Then
                lblCount.Text = "1 terminated repair"
            Else
                lblCount.Text = dt.Rows.Count & " terminated repairs"
            End If
            ShowSelectedReason(Nothing, EventArgs.Empty)
        Catch ex As Exception
            MessageBox.Show("The terminated repairs could not be loaded: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SearchKeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            LoadData()
        End If
    End Sub

    Private Sub ShowSelectedReason(sender As Object, e As EventArgs)
        If grid.CurrentRow Is Nothing Then
            txtReason.Clear()
            Return
        End If
        Dim drv As DataRowView = TryCast(grid.CurrentRow.DataBoundItem, DataRowView)
        If drv Is Nothing Then
            txtReason.Clear()
            Return
        End If
        txtReason.Text = drv("Reason").ToString()
    End Sub

End Class