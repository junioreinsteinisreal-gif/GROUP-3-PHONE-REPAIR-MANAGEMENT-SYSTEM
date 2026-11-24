Imports System.Globalization

''' <summary>Read-only: the employee's own customer reviews and tips. Nothing can be recorded here.</summary>
Public Class FeedbackPage
    Inherits UserControl

    Private ReadOnly stats As New Dictionary(Of String, Label)
    Private grid As DataGridView
    Private ReadOnly refreshTimer As New System.Windows.Forms.Timer()
    Private ReadOnly boldFont As Font = UITheme.FontBold(10.5F)

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("This screen is for employees.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(6)
        BuildPage()
        LoadAll()

        refreshTimer.Interval = 20000
        AddHandler refreshTimer.Tick, Sub(s, e) LoadAll()
        refreshTimer.Start()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            refreshTimer.Stop()
            refreshTimer.Dispose()
        End If
        MyBase.Dispose(disposing)
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

    Private Sub BuildPage()
        Dim root As New TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.ColumnCount = 1
        root.RowCount = 3
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 120.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 50.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        Dim cards As New TableLayoutPanel()
        cards.Dock = DockStyle.Fill
        cards.ColumnCount = 4
        For i As Integer = 1 To 4
            cards.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25.0F))
        Next
        cards.RowCount = 1
        cards.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        cards.Controls.Add(StatCard("tips", "My tips received (all time)", UITheme.SuccessGreen), 0, 0)
        cards.Controls.Add(StatCard("reviews", "My customer reviews", UITheme.RoyalBlue), 1, 0)
        cards.Controls.Add(StatCard("avg", "My average rating (out of 5)", UITheme.Navy), 2, 0)
        cards.Controls.Add(StatCard("good", "Good reviews (4-5 stars)", UITheme.SuccessGreen), 3, 0)
        root.Controls.Add(cards, 0, 0)

        Dim note As New Label()
        note.Dock = DockStyle.Fill
        note.Margin = New Padding(4, 0, 4, 0)
        note.AutoSize = False
        note.TextAlign = ContentAlignment.MiddleLeft
        note.Font = UITheme.FontRegular(10)
        note.ForeColor = UITheme.TextMuted
        note.Text = "Customers leave their review and tip on the review website, using the link in their " &
                    "message. They appear here automatically, credited to you. You cannot add or change them."
        root.Controls.Add(note, 0, 1)

        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        UIHelpers.AddTextColumn(grid, "RecordedAt", "Received", 120)
        UIHelpers.AddTextColumn(grid, "RepairCode", "Repair", 80)
        UIHelpers.AddTextColumn(grid, "CustomerName", "Customer", 130)
        UIHelpers.AddTextColumn(grid, "Rating", "Rating", 60, True)
        UIHelpers.AddTextColumn(grid, "Comment", "Comment", 280)
        UIHelpers.AddTextColumn(grid, "Tip", "Tip", 80, True, "N2")
        AddHandler grid.CellFormatting,
            Sub(s, e)
                If e.RowIndex < 0 OrElse e.Value Is Nothing OrElse IsDBNull(e.Value) Then Return
                Dim name As String = grid.Columns(e.ColumnIndex).DataPropertyName
                If name = "Rating" Then
                    Dim r As Integer = Convert.ToInt32(e.Value)
                    e.CellStyle.ForeColor = If(r >= 4, UITheme.SuccessGreen, If(r <= 2, UITheme.DangerRed, UITheme.TextDark))
                    e.CellStyle.Font = boldFont
                ElseIf name = "Tip" Then
                    e.CellStyle.Font = boldFont
                End If
            End Sub

        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Margin = New Padding(4)
        card.Padding = New Padding(10)
        card.Controls.Add(grid)
        root.Controls.Add(card, 0, 2)

        Me.Controls.Add(root)
    End Sub

    Private Sub LoadAll()
        If Not SessionManager.IsLoggedIn Then
            refreshTimer.Stop()
            Return
        End If
        Try
            Dim s As FeedbackStats = FeedbackService.GetMyStats()
            stats("tips").Text = UIHelpers.Money(s.TipsTotal)
            stats("reviews").Text = s.Reviews.ToString()
            stats("avg").Text = If(s.Reviews > 0, s.AvgRating.ToString("0.0", CultureInfo.InvariantCulture), "-")
            stats("good").Text = s.GoodReviews.ToString()
            grid.DataSource = FeedbackService.GetFeedbackList("")
        Catch
        End Try
    End Sub

End Class