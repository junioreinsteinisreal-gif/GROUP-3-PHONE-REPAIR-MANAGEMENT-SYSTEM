''' <summary>Collection reminders for the whole shop (employees).</summary>
Public Class RemindersPage
    Inherits UserControl

    Private cmbCategory As ComboBox
    Private grid As DataGridView
    Private lblCount As Label
    Private lblOverdue As Label
    Private lblTomorrow As Label
    Private lblAwaiting As Label
    Private suspendEvents As Boolean = False
    Private ReadOnly boldFont As Font = UITheme.FontBold(10.5F)

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Collection reminders are for employees.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)
        BuildPage()
        LoadData()
    End Sub

    Private Function MakeCard(caption As String, color As Color, ByRef valueLabel As Label) As Panel
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Margin = New Padding(4)

        valueLabel = New Label()
        valueLabel.Dock = DockStyle.Fill
        valueLabel.AutoSize = False
        valueLabel.TextAlign = ContentAlignment.MiddleCenter
        valueLabel.Font = UITheme.FontBold(22)
        valueLabel.ForeColor = color
        valueLabel.Text = "0"

        Dim cap As New Label()
        cap.Dock = DockStyle.Bottom
        cap.Height = 34
        cap.AutoSize = False
        cap.TextAlign = ContentAlignment.MiddleCenter
        cap.Font = UITheme.FontRegular(11)
        cap.ForeColor = UITheme.TextMuted
        cap.Text = caption

        card.Controls.Add(valueLabel)
        card.Controls.Add(cap)
        Return card
    End Function

    ' Docking: grid (Fill) first, then footer, then toolbar, then the cards (top-most).
    Private Sub BuildPage()
        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        UIHelpers.AddTextColumn(grid, "Category", "Reminder", 120)
        UIHelpers.AddTextColumn(grid, "RepairCode", "Repair Code", 85)
        UIHelpers.AddTextColumn(grid, "CustomerName", "Customer", 140)
        UIHelpers.AddTextColumn(grid, "CustomerPhone", "Phone", 110)
        UIHelpers.AddTextColumn(grid, "Device", "Device", 120)
        UIHelpers.AddTextColumn(grid, "Status", "Status", 90)
        UIHelpers.AddTextColumn(grid, "DueDate", "Expected Collection", 105)
        UIHelpers.AddTextColumn(grid, "Balance", "Balance", 80, True, "N2")
        UIHelpers.AddTextColumn(grid, "Employee", "Registered By", 110)
        AddHandler grid.CellFormatting, AddressOf GridCellFormatting

        Dim footer As New Panel()
        footer.Dock = DockStyle.Bottom
        footer.Height = 34
        lblCount = UITheme.MakeLabel("", 10, False, UITheme.TextMuted)
        lblCount.Left = 5
        lblCount.Top = 9
        footer.Controls.Add(lblCount)

        Dim toolbar As New FlowLayoutPanel()
        toolbar.Dock = DockStyle.Top
        toolbar.AutoSize = True
        toolbar.AutoSizeMode = AutoSizeMode.GrowAndShrink
        toolbar.Padding = New Padding(0, 0, 0, 6)

        Dim lbl As Label = UITheme.MakeLabel("Show:", 11, True, UITheme.TextDark)
        lbl.Margin = New Padding(6, 12, 4, 0)
        cmbCategory = New ComboBox()
        cmbCategory.Font = UITheme.FontRegular(11)
        cmbCategory.DropDownStyle = ComboBoxStyle.DropDownList
        cmbCategory.Width = 230
        cmbCategory.Margin = New Padding(3, 8, 8, 3)
        cmbCategory.Items.Add("All")
        cmbCategory.Items.AddRange(ReminderService.Categories)
        cmbCategory.SelectedIndex = 0
        AddHandler cmbCategory.SelectedIndexChanged,
            Sub(s, ev)
                If Not suspendEvents Then LoadData()
            End Sub

        Dim btnRefresh As Button = UITheme.MakeButton("Refresh", UITheme.Navy, 120, 38)
        AddHandler btnRefresh.Click, Sub(s, ev) LoadData()

        toolbar.Controls.AddRange(New Control() {lbl, cmbCategory, btnRefresh})

        Dim cards As New TableLayoutPanel()
        cards.Dock = DockStyle.Top
        cards.Height = 110
        cards.ColumnCount = 3
        For i As Integer = 1 To 3
            cards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.33F))
        Next
        cards.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        cards.Controls.Add(MakeCard("Overdue", UITheme.DangerRed, lblOverdue), 0, 0)
        cards.Controls.Add(MakeCard("Due tomorrow", Color.FromArgb(176, 112, 0), lblTomorrow), 1, 0)
        cards.Controls.Add(MakeCard("Awaiting collection", UITheme.SuccessGreen, lblAwaiting), 2, 0)

        Me.Controls.Add(grid)
        Me.Controls.Add(footer)
        Me.Controls.Add(toolbar)
        Me.Controls.Add(cards)
    End Sub

    Private Sub LoadData()
        Try
            Dim c As Integer() = ReminderService.GetCounts()
            lblOverdue.Text = c(0).ToString()
            lblTomorrow.Text = c(1).ToString()
            lblAwaiting.Text = c(2).ToString()

            Dim dt As DataTable = ReminderService.GetReminders(cmbCategory.Text, False, False)
            grid.DataSource = dt
            lblCount.Text = If(dt.Rows.Count = 0, "Nothing to show. All caught up!",
                               dt.Rows.Count & " repair(s) shown. Collected repairs are never listed.")
        Catch ex As Exception
            MessageBox.Show("The reminders could not be loaded: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 OrElse e.Value Is Nothing OrElse IsDBNull(e.Value) Then Return
        If grid.Columns(e.ColumnIndex).DataPropertyName <> "Category" Then Return
        Select Case e.Value.ToString()
            Case "Overdue"
                e.CellStyle.ForeColor = UITheme.DangerRed
            Case "Due tomorrow"
                e.CellStyle.ForeColor = Color.FromArgb(176, 112, 0)
            Case Else
                e.CellStyle.ForeColor = UITheme.SuccessGreen
        End Select
        e.CellStyle.Font = boldFont
    End Sub

End Class