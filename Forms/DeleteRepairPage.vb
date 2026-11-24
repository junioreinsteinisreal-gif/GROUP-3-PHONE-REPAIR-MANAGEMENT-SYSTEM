''' <summary>Employees delete repairs registered by mistake. A copy is kept for administrators.</summary>
Public Class DeleteRepairPage
    Inherits UserControl

    Private txtSearch As TextBox
    Private grid As DataGridView
    Private lblCount As Label
    Private lblSelected As Label
    Private txtReason As TextBox
    Private lblMessage As Label
    Private btnDelete As Button
    Private currentCode As String = ""
    Private binding As Boolean = False

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can delete repairs.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)
        BuildPage()
    End Sub

    ' Docking: grid (Fill) first, then Bottom, then Top controls (last added = top-most).
    Private Sub BuildPage()
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Padding = New Padding(20, 12, 20, 14)

        Dim title As New Label()
        title.Dock = DockStyle.Top
        title.Height = 48
        title.AutoSize = False
        title.TextAlign = ContentAlignment.MiddleLeft
        title.Font = UITheme.FontBold(16)
        title.ForeColor = UITheme.Navy
        title.Text = "Delete Repair"

        Dim note As New Label()
        note.Dock = DockStyle.Top
        note.Height = 62
        note.AutoSize = False
        note.TextAlign = ContentAlignment.TopLeft
        note.Font = UITheme.FontRegular(10)
        note.ForeColor = UITheme.TextMuted
        note.Text = "For repairs registered by mistake. You can delete a repair only if you registered it, it is still " &
                    "in Received status, and no payment has been taken. A permanent copy is kept for the " &
                    "administrator. If the phone cannot be repaired, use Terminate Repair instead."

        ' ----- Search row -----
        Dim searchRow As New TableLayoutPanel()
        searchRow.Dock = DockStyle.Top
        searchRow.Height = 54
        searchRow.ColumnCount = 2
        searchRow.RowCount = 1
        searchRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        searchRow.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 130.0F))
        searchRow.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        txtSearch = UITheme.MakeTextBox(100, 12)
        txtSearch.MaxLength = 60
        txtSearch.Anchor = AnchorStyles.Left Or AnchorStyles.Right
        txtSearch.Margin = New Padding(0, 8, 10, 0)
        txtSearch.PlaceholderText = "Repair code, customer, phone, brand or model"
        AddHandler txtSearch.KeyDown,
            Sub(s, ev)
                If ev.KeyCode = Keys.Enter Then
                    ev.SuppressKeyPress = True
                    RunSearch()
                End If
            End Sub

        Dim btnSearch As Button = UITheme.MakeButton("Search", UITheme.RoyalBlue, 120, 42)
        btnSearch.Margin = New Padding(0, 3, 0, 0)
        AddHandler btnSearch.Click, Sub(s, ev) RunSearch()
        searchRow.Controls.Add(txtSearch, 0, 0)
        searchRow.Controls.Add(btnSearch, 1, 0)

        ' ----- Grid -----
        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        UIHelpers.AddTextColumn(grid, "RepairCode", "Repair Code", 85)
        UIHelpers.AddTextColumn(grid, "CustomerName", "Customer", 130)
        UIHelpers.AddTextColumn(grid, "CustomerPhone", "Phone", 105)
        UIHelpers.AddTextColumn(grid, "Brand", "Brand", 80)
        UIHelpers.AddTextColumn(grid, "Model", "Model", 90)
        UIHelpers.AddTextColumn(grid, "Status", "Status", 90)
        UIHelpers.AddTextColumn(grid, "Paid", "Paid", 80, True, "N2")
        UIHelpers.AddTextColumn(grid, "DateReceived", "Received", 90)
        UIHelpers.AddTextColumn(grid, "Employee", "Registered By", 110)
        AddHandler grid.SelectionChanged, AddressOf GridSelectionChanged

        ' ----- Bottom area -----
        lblSelected = New Label()
        lblSelected.Dock = DockStyle.Top
        lblSelected.Height = 40
        lblSelected.AutoSize = False
        lblSelected.TextAlign = ContentAlignment.MiddleLeft
        lblSelected.Font = UITheme.FontBold(11.5F)
        lblSelected.ForeColor = UITheme.Navy

        Dim lblReason As New Label()
        lblReason.Dock = DockStyle.Top
        lblReason.Height = 30
        lblReason.AutoSize = False
        lblReason.TextAlign = ContentAlignment.MiddleLeft
        lblReason.Font = UITheme.FontBold(10.5F)
        lblReason.ForeColor = UITheme.TextDark
        lblReason.Text = "Reason for deletion (optional, but recommended)"

        txtReason = UITheme.MakeTextBox(100, 11)
        txtReason.Multiline = True
        txtReason.ScrollBars = ScrollBars.Vertical
        txtReason.MaxLength = RepairDeletionService.MaxReasonLength
        txtReason.Dock = DockStyle.Fill

        lblMessage = New Label()
        lblMessage.Dock = DockStyle.Bottom
        lblMessage.Height = 46
        lblMessage.AutoSize = False
        lblMessage.TextAlign = ContentAlignment.MiddleLeft
        lblMessage.Font = UITheme.FontBold(10.5F)

        btnDelete = UITheme.MakeButton("Delete Repair", UITheme.DangerRed, 230, 52)
        btnDelete.Margin = New Padding(0, 4, 12, 0)
        AddHandler btnDelete.Click, AddressOf DeleteClick

        Dim btnClear As Button = UITheme.MakeButton("Clear", UITheme.LightGrey, 140, 52)
        btnClear.ForeColor = UITheme.TextDark
        btnClear.Margin = New Padding(0, 4, 0, 0)
        AddHandler btnClear.Click, Sub(s, ev) ClearAll()

        Dim buttonRow As New FlowLayoutPanel()
        buttonRow.Dock = DockStyle.Bottom
        buttonRow.Height = 62
        buttonRow.WrapContents = False
        buttonRow.Controls.Add(btnDelete)
        buttonRow.Controls.Add(btnClear)

        Dim bottom As New Panel()
        bottom.Dock = DockStyle.Bottom
        bottom.Height = 270
        bottom.Controls.Add(txtReason)
        bottom.Controls.Add(lblMessage)
        bottom.Controls.Add(buttonRow)
        bottom.Controls.Add(lblReason)
        bottom.Controls.Add(lblSelected)

        lblCount = New Label()
        lblCount.Dock = DockStyle.Bottom
        lblCount.Height = 30
        lblCount.AutoSize = False
        lblCount.TextAlign = ContentAlignment.MiddleLeft
        lblCount.Font = UITheme.FontRegular(10)
        lblCount.ForeColor = UITheme.TextMuted

        card.Controls.Add(grid)
        card.Controls.Add(lblCount)
        card.Controls.Add(bottom)
        card.Controls.Add(searchRow)
        card.Controls.Add(note)
        card.Controls.Add(title)
        Me.Controls.Add(card)
    End Sub

    Private Sub ShowMessage(text As String, color As Color)
        lblMessage.ForeColor = color
        lblMessage.Text = text
    End Sub

    Private Sub RunSearch()
        Dim q As String = txtSearch.Text.Trim()
        If q = "" Then
            ShowMessage("Enter something to search for.", UITheme.DangerRed)
            txtSearch.Focus()
            Return
        End If
        Try
            Dim dt As DataTable = RepairService.GetRepairs(q, "All")
            binding = True
            grid.DataSource = dt
            grid.ClearSelection()
            grid.CurrentCell = Nothing
            binding = False
            currentCode = ""
            lblSelected.Text = ""
            ShowMessage("", UITheme.TextDark)
            lblCount.Text = If(dt.Rows.Count = 0, "No repairs found.",
                               dt.Rows.Count & " repair(s) found. Click one to select it.")
            ActivityLogger.Log("SearchRepairs", "Searched repairs for '" & q & "' (delete screen).")
        Catch ex As Exception
            binding = False
            ShowMessage("The search failed: " & ex.Message, UITheme.DangerRed)
        End Try
    End Sub

    Private Sub GridSelectionChanged(sender As Object, e As EventArgs)
        If binding Then Return
        If grid.CurrentRow Is Nothing OrElse Not grid.CurrentRow.Selected Then Return
        Dim drv As DataRowView = TryCast(grid.CurrentRow.DataBoundItem, DataRowView)
        If drv Is Nothing Then Return
        currentCode = drv("RepairCode").ToString()
        lblSelected.Text = "Selected: " & currentCode & "  |  " & drv("CustomerName").ToString() &
                           "  |  " & drv("Brand").ToString() & " " & drv("Model").ToString() &
                           "  |  " & drv("Status").ToString()
        ShowMessage("", UITheme.TextDark)
    End Sub

    Private Sub DeleteClick(sender As Object, e As EventArgs)
        ShowMessage("", UITheme.TextDark)
        If currentCode = "" Then
            ShowMessage("Search for a repair and select it first.", UITheme.DangerRed)
            Return
        End If

        If MessageBox.Show("Delete repair " & currentCode & "?" & vbCrLf & vbCrLf &
                           "It will be removed from the repair list. A permanent copy is kept " &
                           "for the administrator.",
                           "Confirm Deletion", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Warning) <> DialogResult.Yes Then Return

        Dim code As String = currentCode
        Dim err As String = ""
        btnDelete.Enabled = False
        Try
            err = RepairDeletionService.DeleteRepair(code, txtReason.Text)
        Catch ex As UnauthorizedAccessException
            err = ex.Message
        Finally
            btnDelete.Enabled = True
        End Try

        ' On failure nothing is cleared.
        If err <> "" Then
            ShowMessage(err, UITheme.DangerRed)
            Return
        End If

        MessageBox.Show("Repair " & code & " has been deleted." & vbCrLf & vbCrLf &
                        "The administrator keeps a permanent record of it.",
                        "Repair Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ClearAll()
        ShowMessage("Repair " & code & " deleted.", UITheme.SuccessGreen)
    End Sub

    Private Sub ClearAll()
        txtSearch.Clear()
        txtReason.Clear()
        binding = True
        grid.DataSource = Nothing
        binding = False
        currentCode = ""
        lblSelected.Text = ""
        lblCount.Text = ""
        ShowMessage("", UITheme.TextDark)
        txtSearch.Focus()
    End Sub

End Class