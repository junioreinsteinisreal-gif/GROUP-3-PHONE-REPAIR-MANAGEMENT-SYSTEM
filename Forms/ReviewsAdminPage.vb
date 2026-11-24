Imports System.Globalization

''' <summary>Administrator view of every customer review and tip, with a per-employee summary.</summary>
Public Class ReviewsAdminPage
    Inherits UserControl

    Private ReadOnly stats As New Dictionary(Of String, Label)
    Private txtSearch As TextBox
    Private gridSummary As DataGridView
    Private gridList As DataGridView
    Private lblCount As Label
    Private ReadOnly boldFont As Font = UITheme.FontBold(10.5F)

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsAdmin Then
            Throw New UnauthorizedAccessException("Only administrators can view all reviews and tips.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(6)
        BuildPage()
        LoadData()
    End Sub

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
        cap.Height = 44
        cap.AutoSize = False
        cap.TextAlign = ContentAlignment.TopCenter
        cap.Font = UITheme.FontRegular(9.5F)
        cap.ForeColor = UITheme.TextMuted
        cap.Text = caption

        card.Controls.Add(value)
        card.Controls.Add(cap)
        Return card
    End Function

    Private Function Section(title As String, content As Control) As Panel
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
        card.Controls.Add(lbl)
        Return card
    End Function

    Private Function NewGrid() As DataGridView
        Dim g As New DataGridView()
        g.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(g)
        g.RowTemplate.Height = 36
        g.ColumnHeadersHeight = 42
        AddHandler g.CellFormatting, AddressOf GridCellFormatting
        Return g
    End Function

    Private Sub BuildPage()
        Dim root As New TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.ColumnCount = 1
        root.RowCount = 4
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 120.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 58.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 34.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 66.0F))

        Dim cards As New TableLayoutPanel()
        cards.Dock = DockStyle.Fill
        cards.ColumnCount = 4
        For i As Integer = 1 To 4
            cards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
        Next
        cards.RowCount = 1
        cards.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        cards.Controls.Add(StatCard("reviews", "Customer reviews", UITheme.RoyalBlue), 0, 0)
        cards.Controls.Add(StatCard("avg", "Average rating (out of 5)", UITheme.Navy), 1, 0)
        cards.Controls.Add(StatCard("good", "Good reviews (4-5 stars)", UITheme.SuccessGreen), 2, 0)
        cards.Controls.Add(StatCard("tips", "Tips received (all employees)", UITheme.SuccessGreen), 3, 0)
        root.Controls.Add(cards, 0, 0)

        ' ----- Toolbar -----
        Dim bar As New FlowLayoutPanel()
        bar.Dock = DockStyle.Fill
        bar.BackColor = UITheme.CardWhite
        bar.Margin = New Padding(4)
        bar.Padding = New Padding(8, 8, 8, 2)
        bar.WrapContents = False

        txtSearch = UITheme.MakeTextBox(280, 11)
        txtSearch.MaxLength = 60
        txtSearch.PlaceholderText = "Repair, customer, employee or comment"
        txtSearch.Margin = New Padding(3, 4, 6, 3)
        AddHandler txtSearch.KeyDown,
            Sub(s, ev)
                If ev.KeyCode = Keys.Enter Then
                    ev.SuppressKeyPress = True
                    LoadData()
                End If
            End Sub

        Dim btnSearch As Button = UITheme.MakeButton("Search", UITheme.RoyalBlue, 110, 38)
        AddHandler btnSearch.Click, Sub(s, ev) LoadData()
        Dim btnClear As Button = UITheme.MakeButton("Clear", UITheme.TextMuted, 90, 38)
        AddHandler btnClear.Click,
            Sub(s, ev)
                txtSearch.Clear()
                LoadData()
            End Sub
        Dim btnExport As Button = UITheme.MakeButton("Export to Excel", UITheme.SuccessGreen, 170, 38)
        AddHandler btnExport.Click, Sub(s, ev) ExcelExporter.ExportGrid(gridList, "ReviewsAndTips", "Reviews")
        lblCount = UITheme.MakeLabel("", 10, True, UITheme.TextMuted)
        lblCount.Margin = New Padding(12, 10, 0, 0)
        bar.Controls.AddRange(New Control() {txtSearch, btnSearch, btnClear, btnExport, lblCount})
        root.Controls.Add(bar, 0, 1)

        gridSummary = NewGrid()
        UIHelpers.AddTextColumn(gridSummary, "Employee", "Employee", 160)
        UIHelpers.AddTextColumn(gridSummary, "Reviews", "Reviews", 70, True)
        UIHelpers.AddTextColumn(gridSummary, "AvgRating", "Average Rating", 90, True, "0.0")
        UIHelpers.AddTextColumn(gridSummary, "GoodReviews", "Good Reviews (4-5)", 100, True)
        UIHelpers.AddTextColumn(gridSummary, "VerifiedGood", "Of Which Verified", 100, True)
        UIHelpers.AddTextColumn(gridSummary, "Tips", "Tips Received", 100, True, "N2")
        root.Controls.Add(Section("Employee summary", gridSummary), 0, 2)

        gridList = NewGrid()
        UIHelpers.AddTextColumn(gridList, "RecordedAt", "Recorded", 115)
        UIHelpers.AddTextColumn(gridList, "RepairCode", "Repair", 70)
        UIHelpers.AddTextColumn(gridList, "CustomerName", "Customer", 110)
        UIHelpers.AddTextColumn(gridList, "Employee", "Employee", 110)
        UIHelpers.AddTextColumn(gridList, "Rating", "Rating", 60, True)
        UIHelpers.AddTextColumn(gridList, "Comment", "Comment", 240)
        UIHelpers.AddTextColumn(gridList, "Tip", "Tip", 75, True, "N2")
        UIHelpers.AddTextColumn(gridList, "Verified", "Verified", 70)
        root.Controls.Add(Section("All customer feedback", gridList), 0, 3)

        Me.Controls.Add(root)
    End Sub

    Private Sub LoadData()
        Try
            Dim s As FeedbackStats = FeedbackService.GetShopStats()
            stats("reviews").Text = s.Reviews.ToString()
            stats("avg").Text = If(s.Reviews > 0, s.AvgRating.ToString("0.0", CultureInfo.InvariantCulture), "-")
            stats("good").Text = s.GoodReviews.ToString()
            stats("tips").Text = UIHelpers.Money(s.TipsTotal)

            gridSummary.DataSource = FeedbackService.GetEmployeeSummary()
            Dim dt As DataTable = FeedbackService.GetFeedbackList(txtSearch.Text)
            gridList.DataSource = dt
            lblCount.Text = dt.Rows.Count & " review(s)"
        Catch ex As Exception
            MessageBox.Show("Could not load the reviews: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 OrElse e.Value Is Nothing OrElse IsDBNull(e.Value) Then Return
        Dim g As DataGridView = DirectCast(sender, DataGridView)
        Dim name As String = g.Columns(e.ColumnIndex).DataPropertyName
        If name = "Rating" Then
            Dim r As Integer = Convert.ToInt32(e.Value)
            e.CellStyle.ForeColor = If(r >= 4, UITheme.SuccessGreen, If(r <= 2, UITheme.DangerRed, UITheme.TextDark))
            e.CellStyle.Font = boldFont
        ElseIf name = "Verified" AndAlso e.Value.ToString() = "No" Then
            e.CellStyle.ForeColor = Color.FromArgb(176, 112, 0)
        ElseIf name = "Tips" OrElse name = "Tip" Then
            e.CellStyle.Font = boldFont
        End If
    End Sub

End Class