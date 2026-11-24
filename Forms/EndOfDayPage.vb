Imports System.IO

''' <summary>End of day. Administrators close the business day; employees see only their own figures.</summary>
Public Class EndOfDayPage
    Inherits UserControl

    Private ReadOnly isAdminView As Boolean
    Private dtp As DateTimePicker
    Private lblBanner As Label
    Private ReadOnly stats As New Dictionary(Of String, Label)
    Private grid As DataGridView
    Private btnAction As Button
    Private btnReopen As Button
    Private ReadOnly boldFont As Font = UITheme.FontBold(10.5F)

    Public Sub New()
        SessionManager.RequireLogin()
        isAdminView = SessionManager.IsAdmin
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(6)
        BuildPage()
        LoadData()
    End Sub

    ' ---------- Layout ----------

    Private Sub BuildPage()
        Dim root As New TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.ColumnCount = 1
        root.RowCount = 4
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 70.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 52.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Absolute, 120.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        root.Controls.Add(BuildToolbar(), 0, 0)
        root.Controls.Add(BuildBanner(), 0, 1)
        root.Controls.Add(BuildCards(), 0, 2)
        root.Controls.Add(BuildBody(), 0, 3)
        Me.Controls.Add(root)
    End Sub

    Private Function BuildToolbar() As Control
        Dim bar As New FlowLayoutPanel()
        bar.Dock = DockStyle.Fill
        bar.BackColor = UITheme.CardWhite
        bar.Margin = New Padding(4)
        bar.Padding = New Padding(10, 8, 10, 4)
        bar.WrapContents = False

        Dim cap As Label = UITheme.MakeLabel(If(isAdminView, "Business date:", "Today:"), 11, True, UITheme.TextDark)
        cap.Margin = New Padding(4, 12, 6, 0)
        bar.Controls.Add(cap)

        If isAdminView Then
            dtp = New DateTimePicker()
            dtp.Font = UITheme.FontRegular(11)
            dtp.Format = DateTimePickerFormat.Custom
            dtp.CustomFormat = "dddd, dd MMM yyyy"
            dtp.MaxDate = Date.Today
            dtp.Value = Date.Today
            dtp.Width = 290
            dtp.Margin = New Padding(3, 8, 12, 3)
            AddHandler dtp.ValueChanged, Sub(s, ev) LoadData()
            bar.Controls.Add(dtp)
        Else
            Dim lblToday As Label = UITheme.MakeLabel(Date.Today.ToString("dddd, dd MMMM yyyy"), 11, False, UITheme.Navy)
            lblToday.Margin = New Padding(3, 12, 20, 0)
            bar.Controls.Add(lblToday)
        End If

        Dim btnRefresh As Button = UITheme.MakeButton("Refresh", UITheme.Navy, 120, 40)
        btnRefresh.Margin = New Padding(3, 5, 6, 3)
        AddHandler btnRefresh.Click, Sub(s, ev) LoadData()
        bar.Controls.Add(btnRefresh)

        If isAdminView Then
            Dim btnExport As Button = UITheme.MakeButton("Export Report", UITheme.RoyalBlue, 170, 40)
            btnExport.Margin = New Padding(3, 5, 6, 3)
            AddHandler btnExport.Click, AddressOf ExportClick
            bar.Controls.Add(btnExport)

            btnReopen = UITheme.MakeButton("Reopen Day", UITheme.WarningAmber, 160, 40)
            btnReopen.ForeColor = UITheme.TextDark
            btnReopen.Margin = New Padding(3, 5, 6, 3)
            AddHandler btnReopen.Click, AddressOf ReopenClick
            bar.Controls.Add(btnReopen)

            btnAction = UITheme.MakeButton("Close the Day", UITheme.SuccessGreen, 190, 40)
            AddHandler btnAction.Click, AddressOf CloseClick
        Else
            btnAction = UITheme.MakeButton("Finish My Shift && Log Out", UITheme.DangerRed, 300, 40)
            AddHandler btnAction.Click, AddressOf FinishClick
        End If
        btnAction.Margin = New Padding(3, 5, 6, 3)
        bar.Controls.Add(btnAction)
        Return bar
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
        tbl.Controls.Add(StatCard("rev", If(isAdminView, "Total revenue (net of refunds)", "My revenue (net of refunds)"), UITheme.SuccessGreen), 0, 0)
        tbl.Controls.Add(StatCard("pays", "Number of payments", UITheme.RoyalBlue), 1, 0)
        tbl.Controls.Add(StatCard("ref", "Refunds given", UITheme.DangerRed), 2, 0)
        tbl.Controls.Add(StatCard("reg", "Repairs registered", UITheme.Navy), 3, 0)
        tbl.Controls.Add(StatCard("done", "Repairs completed", UITheme.SuccessGreen), 4, 0)
        tbl.Controls.Add(StatCard("out", If(isAdminView, "Outstanding balances", "Outstanding on my repairs"), amber), 5, 0)
        Return tbl
    End Function

    Private Function BuildBody() As Control
        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        If isAdminView Then
            UIHelpers.AddTextColumn(grid, "Employee", "Employee", 160)
            UIHelpers.AddTextColumn(grid, "Payments", "Payments", 80, True)
            UIHelpers.AddTextColumn(grid, "Collected", "Collected", 100, True, "N2")
            UIHelpers.AddTextColumn(grid, "Refunded", "Refunded", 100, True, "N2")
            UIHelpers.AddTextColumn(grid, "Net", "Net Revenue", 110, True, "N2")
        Else
            UIHelpers.AddTextColumn(grid, "PaidAt", "Date and Time", 130)
            UIHelpers.AddTextColumn(grid, "ReceiptNumber", "Receipt", 80)
            UIHelpers.AddTextColumn(grid, "RepairCode", "Repair", 80)
            UIHelpers.AddTextColumn(grid, "CustomerName", "Customer", 140)
            UIHelpers.AddTextColumn(grid, "Amount", "Amount", 90, True, "N2")
            UIHelpers.AddTextColumn(grid, "PaymentMethod", "Method", 100)
        End If
        AddHandler grid.CellFormatting,
            Sub(s, e)
                If e.RowIndex >= 0 AndAlso grid.Columns(e.ColumnIndex).DataPropertyName = "Net" Then
                    e.CellStyle.Font = boldFont
                End If
            End Sub

        Dim lbl As New Label()
        lbl.Dock = DockStyle.Top
        lbl.Height = 40
        lbl.AutoSize = False
        lbl.TextAlign = ContentAlignment.MiddleLeft
        lbl.Font = UITheme.FontBold(11.5F)
        lbl.ForeColor = UITheme.Navy
        lbl.Text = If(isAdminView, "Employee revenue breakdown", "My payments today")

        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Margin = New Padding(4)
        card.Padding = New Padding(10, 4, 10, 10)
        card.Controls.Add(grid)
        card.Controls.Add(lbl)
        Return card
    End Function

    ' ---------- Data ----------

    Private Sub LoadData()
        Try
            Dim bizDate As Date = If(isAdminView, dtp.Value.Date, Date.Today)
            Dim f As DayFigures = EndOfDayService.GetDayFigures(bizDate, 0)
            stats("rev").Text = UIHelpers.Money(f.NetRevenue)
            stats("pays").Text = f.PaymentCount.ToString()
            stats("ref").Text = UIHelpers.Money(f.Refunded)
            stats("reg").Text = f.Registered.ToString()
            stats("done").Text = f.Completed.ToString()
            stats("out").Text = UIHelpers.Money(f.Outstanding)

            If isAdminView Then
                grid.DataSource = EndOfDayService.GetEmployeeBreakdown(bizDate)
                Dim c As DataTable = EndOfDayService.GetClosing(bizDate)
                If c.Rows.Count > 0 Then
                    lblBanner.BackColor = UITheme.SuccessGreen
                    lblBanner.Text = "Day closed by " & c.Rows(0)("ClosedByName").ToString() & " on " &
                                     c.Rows(0)("ClosedAt").ToString() & ". It cannot be closed again."
                    btnAction.Enabled = False
                Else
                    lblBanner.BackColor = Color.FromArgb(176, 112, 0)
                    lblBanner.Text = "This day has not been closed yet."
                    btnAction.Enabled = True
                End If
            Else
                grid.DataSource = EndOfDayService.GetPaymentsForDate(bizDate, 0)
                lblBanner.BackColor = UITheme.RoyalBlue
                lblBanner.Text = "These are your own figures for today. When you are done, use Finish My Shift."
            End If
        Catch ex As Exception
            lblBanner.BackColor = UITheme.DangerRed
            lblBanner.Text = "The figures could not be loaded: " & ex.Message
        End Try
    End Sub

    ' ---------- Actions ----------

    Private Sub OfferOpen(reportPath As String)
        If MessageBox.Show("Report saved:" & vbCrLf & reportPath & vbCrLf & vbCrLf & "Open the folder now?",
                           "Excel Report", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Information) <> DialogResult.Yes Then Return
        Try
            Process.Start("explorer.exe", "/select,""" & reportPath & """")
        Catch
        End Try
    End Sub

    Private Sub CloseClick(sender As Object, e As EventArgs)
        Dim bizDate As Date = dtp.Value.Date
        Dim warn As String = ""
        Try
            Dim f As DayFigures = EndOfDayService.GetDayFigures(bizDate, 0)
            If f.PaymentCount = 0 AndAlso f.Registered = 0 Then
                warn = vbCrLf & vbCrLf & "Note: no payments or new repairs were recorded on this day."
            End If
        Catch
        End Try

        If MessageBox.Show("Close the working day for " & bizDate.ToString("dddd, dd MMMM yyyy") & "?" &
                           vbCrLf & vbCrLf & "This saves the day's totals and creates an Excel report. " &
                          "A day can only be closed once. Employees will be logged out and cannot " &
                           "log in again until 6:00 AM tomorrow." & warn,
                           "Close the Day", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question) <> DialogResult.Yes Then Return

        Dim reportPath As String = ""
        Dim reportError As String = ""
        Dim err As String = ""
        btnAction.Enabled = False
        Try
            err = EndOfDayService.CloseDay(bizDate, reportPath, reportError)
        Catch ex As UnauthorizedAccessException
            err = ex.Message
        End Try
        LoadData()

        If err <> "" Then
            MessageBox.Show(err, "Close the Day", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If reportError <> "" Then
            MessageBox.Show("The day was closed and saved, but the Excel report could not be created:" &
                            vbCrLf & reportError & vbCrLf & vbCrLf &
                            "Use Export Report to try again.",
                            "Day Closed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        OfferOpen(reportPath)
    End Sub

    Private Sub ReopenClick(sender As Object, e As EventArgs)
        Dim bizDate As Date = dtp.Value.Date
        If MessageBox.Show("Reopen the working day for " & bizDate.ToString("dddd, dd MMMM yyyy") & "?" &
                           vbCrLf & vbCrLf & "Employees will be able to log in again. The closing " &
                           "figures are kept in the Administrator History, and the day can be closed again.",
                           "Reopen the Day", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question) <> DialogResult.Yes Then Return

        Dim err As String = ""
        Try
            err = EndOfDayService.ReopenDay(bizDate)
        Catch ex As UnauthorizedAccessException
            err = ex.Message
        End Try
        If err <> "" Then
            MessageBox.Show(err, "Reopen the Day", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        LoadData()
        MessageBox.Show("The day has been reopened. Employees can log in again.",
                        "Reopen the Day", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ExportClick(sender As Object, e As EventArgs)
        Dim reportPath As String = ""
        Try
            reportPath = EndOfDayService.BuildReport(dtp.Value.Date)
        Catch ioEx As IOException
            MessageBox.Show("The report could not be written. If it is open in Excel, close it and try again.",
                            "Export Report", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        Catch ex As Exception
            MessageBox.Show("The report failed: " & ex.Message, "Export Report",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try
        OfferOpen(reportPath)
    End Sub

    Private Sub FinishClick(sender As Object, e As EventArgs)
        If MessageBox.Show("Finish your shift?" & vbCrLf & vbCrLf &
                           "Your figures for today are saved in your activity history and you will be logged out.",
                           "Finish My Shift", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question) <> DialogResult.Yes Then Return
        Try
            Dim f As DayFigures = EndOfDayService.GetDayFigures(Date.Today, 0)
            ActivityLogger.Log("EndOfDay",
                               "Finished the day: " & f.PaymentCount & " payment(s), net " &
                               UIHelpers.Money(f.NetRevenue) & ", " & f.Registered &
                               " repair(s) registered, " & f.Completed & " completed.")
        Catch
        End Try
        AuthService.Logout("End of day")
        Dim frm As Form = Me.FindForm()
        If frm IsNot Nothing Then frm.Close()
    End Sub

End Class