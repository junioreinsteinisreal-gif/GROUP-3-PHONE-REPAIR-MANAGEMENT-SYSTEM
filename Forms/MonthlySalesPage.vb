Imports System.IO

''' <summary>Monthly sales dashboard (administrator only).</summary>
Public Class MonthlySalesPage
    Inherits UserControl

    Private dtp As DateTimePicker
    Private ReadOnly stats As New Dictionary(Of String, Label)
    Private gridWeeks As DataGridView
    Private gridEmp As DataGridView
    Private ReadOnly boldFont As Font = UITheme.FontBold(10.5F)

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsAdmin Then
            Throw New UnauthorizedAccessException("Only administrators can view monthly sales.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(6)
        BuildPage()
        LoadData()
    End Sub

    Private Sub BuildPage()
        Dim root As New TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.ColumnCount = 1
        root.RowCount = 4
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 70.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 120.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 38.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 62.0F))
        root.Controls.Add(BuildToolbar(), 0, 0)
        root.Controls.Add(BuildCards(), 0, 1)

        gridWeeks = NewGrid()
        UIHelpers.AddTextColumn(gridWeeks, "Week", "Week", 190)
        UIHelpers.AddTextColumn(gridWeeks, "Payments", "Payments", 80, True)
        UIHelpers.AddTextColumn(gridWeeks, "Collected", "Collected", 100, True, "N2")
        UIHelpers.AddTextColumn(gridWeeks, "Refunded", "Refunded", 100, True, "N2")
        UIHelpers.AddTextColumn(gridWeeks, "Net", "Net Revenue", 110, True, "N2")
        root.Controls.Add(Section("Weekly revenue breakdown", gridWeeks), 0, 2)

        gridEmp = NewGrid()
        UIHelpers.AddTextColumn(gridEmp, "Employee", "Employee", 160)
        UIHelpers.AddTextColumn(gridEmp, "Registered", "Repairs Registered", 90, True)
        UIHelpers.AddTextColumn(gridEmp, "Payments", "Payments", 80, True)
        UIHelpers.AddTextColumn(gridEmp, "Collected", "Collected", 100, True, "N2")
        UIHelpers.AddTextColumn(gridEmp, "Refunded", "Refunded", 100, True, "N2")
        UIHelpers.AddTextColumn(gridEmp, "Net", "Net Revenue", 110, True, "N2")
        UIHelpers.AddTextColumn(gridEmp, "AvgPayment", "Average Payment", 100, True, "N2")
        root.Controls.Add(Section("Employee sales performance", gridEmp), 0, 3)

        Me.Controls.Add(root)
    End Sub

    Private Function BuildToolbar() As Control
        Dim bar As New FlowLayoutPanel()
        bar.Dock = DockStyle.Fill
        bar.BackColor = UITheme.CardWhite
        bar.Margin = New Padding(4)
        bar.Padding = New Padding(10, 8, 10, 4)
        bar.WrapContents = False

        Dim cap As Label = UITheme.MakeLabel("Month:", 11, True, UITheme.TextDark)
        cap.Margin = New Padding(4, 12, 6, 0)

        dtp = New DateTimePicker()
        dtp.Font = UITheme.FontRegular(11)
        dtp.Format = DateTimePickerFormat.Custom
        dtp.CustomFormat = "MMMM yyyy"
        dtp.ShowUpDown = True
        dtp.MaxDate = Date.Today
        dtp.Value = Date.Today
        dtp.Width = 220
        dtp.Margin = New Padding(3, 8, 12, 3)
        AddHandler dtp.ValueChanged, Sub(s, ev) LoadData()

        Dim btnRefresh As Button = UITheme.MakeButton("Refresh", UITheme.Navy, 120, 40)
        btnRefresh.Margin = New Padding(3, 5, 6, 3)
        AddHandler btnRefresh.Click, Sub(s, ev) LoadData()

        Dim btnExport As Button = UITheme.MakeButton("Export to Excel", UITheme.SuccessGreen, 190, 40)
        btnExport.Margin = New Padding(3, 5, 6, 3)
        AddHandler btnExport.Click, AddressOf ExportClick

        Dim note As Label = UITheme.MakeLabel("Average per day = net revenue / days so far in the month.",
                                              9.5F, False, UITheme.TextMuted)
        note.Margin = New Padding(14, 14, 0, 0)

        bar.Controls.AddRange(New Control() {cap, dtp, btnRefresh, btnExport, note})
        Return bar
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
        tbl.Controls.Add(StatCard("rev", "Monthly revenue (net of refunds)", UITheme.SuccessGreen), 0, 0)
        tbl.Controls.Add(StatCard("pays", "Total payments", UITheme.RoyalBlue), 1, 0)
        tbl.Controls.Add(StatCard("ref", "Refunds given", UITheme.DangerRed), 2, 0)
        tbl.Controls.Add(StatCard("avgday", "Average revenue per day", UITheme.Navy), 3, 0)
        tbl.Controls.Add(StatCard("avgpay", "Average payment", UITheme.Navy), 4, 0)
        tbl.Controls.Add(StatCard("reg", "Repairs registered", Color.FromArgb(176, 112, 0)), 5, 0)
        Return tbl
    End Function

    Private Function NewGrid() As DataGridView
        Dim g As New DataGridView()
        g.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(g)
        g.RowTemplate.Height = 38
        g.ColumnHeadersHeight = 44
        AddHandler g.CellFormatting,
            Sub(s, e)
                If e.RowIndex >= 0 AndAlso g.Columns(e.ColumnIndex).DataPropertyName = "Net" Then
                    e.CellStyle.Font = boldFont
                End If
            End Sub
        Return g
    End Function

    Private Function Section(title As String, content As Control) As Panel
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Margin = New Padding(4)
        card.Padding = New Padding(10, 4, 10, 10)
        Dim lbl As New Label()
        lbl.Dock = DockStyle.Top
        lbl.Height = 40
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

    Private Sub LoadData()
        Try
            Dim yr As Integer = dtp.Value.Year
            Dim mo As Integer = dtp.Value.Month
            Dim f As MonthFigures = MonthlySalesService.GetMonthFigures(yr, mo)
            stats("rev").Text = UIHelpers.Money(f.NetRevenue)
            stats("pays").Text = f.PaymentCount.ToString()
            stats("ref").Text = UIHelpers.Money(f.Refunded)
            stats("avgday").Text = UIHelpers.Money(f.AvgPerDay)
            stats("avgpay").Text = UIHelpers.Money(f.AvgPerPayment)
            stats("reg").Text = f.Registered.ToString()
            gridWeeks.DataSource = MonthlySalesService.GetWeekly(yr, mo)
            gridEmp.DataSource = MonthlySalesService.GetEmployeeMonth(yr, mo)
        Catch ex As Exception
            MessageBox.Show("The monthly figures could not be loaded: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExportClick(sender As Object, e As EventArgs)
        Dim reportPath As String = ""
        Try
            reportPath = MonthlySalesService.BuildMonthReport(dtp.Value.Year, dtp.Value.Month)
        Catch ioEx As IOException
            MessageBox.Show("The report could not be written. If it is open in Excel, close it and try again.",
                            "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        Catch ex As Exception
            MessageBox.Show("The export failed: " & ex.Message, "Export to Excel",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        If MessageBox.Show("Report saved:" & vbCrLf & reportPath & vbCrLf & vbCrLf & "Open the folder now?",
                           "Export to Excel", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Information) = DialogResult.Yes Then
            Try
                Process.Start("explorer.exe", "/select,""" & reportPath & """")
            Catch
            End Try
        End If
    End Sub

End Class