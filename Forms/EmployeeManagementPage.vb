Imports System.Text.RegularExpressions

''' <summary>Administrator-only employee directory and management.</summary>
Public Class EmployeeManagementPage
    Inherits UserControl

    Private txtSearch As TextBox
    Private grid As DataGridView
    Private lblTotal As Label
    Private lblActive As Label
    Private lblInactive As Label
    Private ReadOnly boldFont As Font = UITheme.FontBold(10.5F)

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsAdmin Then
            Throw New UnauthorizedAccessException("Only administrators can manage employees.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)
        BuildPage()
        LoadData()
    End Sub

    ' ---------- Layout (Fill first, then toolbar, then stats = top-most) ----------

    Private Sub BuildPage()
        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        UIHelpers.AddTextColumn(grid, "EmployeeId", "ID", 40)
        UIHelpers.AddTextColumn(grid, "FullName", "Employee Name", 200)
        UIHelpers.AddTextColumn(grid, "AccessCode", "Access Code", 90)
        UIHelpers.AddTextColumn(grid, "RegisteredDate", "Registered", 100)
        UIHelpers.AddTextColumn(grid, "StatusText", "Status", 90)
        AddHandler grid.CellFormatting, AddressOf GridCellFormatting

        Dim toolbar As New FlowLayoutPanel()
        toolbar.Dock = DockStyle.Top
        toolbar.AutoSize = True
        toolbar.AutoSizeMode = AutoSizeMode.GrowAndShrink
        toolbar.WrapContents = True
        toolbar.Padding = New Padding(0, 0, 0, 6)

        txtSearch = UITheme.MakeTextBox(250, 11)
        txtSearch.MaxLength = 60
        txtSearch.PlaceholderText = "Name or employee ID"
        txtSearch.Margin = New Padding(5, 5, 5, 3)
        AddHandler txtSearch.KeyDown,
            Sub(s, ev)
                If ev.KeyCode = Keys.Enter Then
                    ev.SuppressKeyPress = True
                    LoadData()
                End If
            End Sub

        toolbar.Controls.Add(txtSearch)
        toolbar.Controls.Add(Btn("Search", UITheme.RoyalBlue, 100, Sub(s, ev) LoadData()))
        toolbar.Controls.Add(Btn("Clear", UITheme.TextMuted, 90,
                                 Sub(s, ev)
                                     txtSearch.Clear()
                                     LoadData()
                                 End Sub))
        toolbar.Controls.Add(Btn("Register Employee", UITheme.SuccessGreen, 190, AddressOf RegisterClick))
        toolbar.Controls.Add(Btn("Reset Code", UITheme.RoyalBlue, 130, AddressOf ResetClick))
        toolbar.Controls.Add(Btn("Activate", UITheme.SuccessGreen, 110, AddressOf ActivateClick))
        toolbar.Controls.Add(Btn("Deactivate", UITheme.DangerRed, 130, AddressOf DeactivateClick))
        toolbar.Controls.Add(Btn("Performance && Revenue", UITheme.Navy, 230, AddressOf PerformanceClick))
        toolbar.Controls.Add(Btn("View Activity", UITheme.Navy, 150, AddressOf ActivityClick))

        Dim stats As New TableLayoutPanel()
        stats.Dock = DockStyle.Top
        stats.Height = 110
        stats.ColumnCount = 3
        For i As Integer = 1 To 3
            stats.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.33F))
        Next
        stats.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        stats.Controls.Add(StatCard("Total employees", UITheme.Navy, lblTotal), 0, 0)
        stats.Controls.Add(StatCard("Active employees", UITheme.SuccessGreen, lblActive), 1, 0)
        stats.Controls.Add(StatCard("Inactive employees", UITheme.DangerRed, lblInactive), 2, 0)

        Me.Controls.Add(grid)
        Me.Controls.Add(toolbar)
        Me.Controls.Add(stats)
    End Sub

    Private Function Btn(text As String, color As Color, width As Integer,
                         handler As EventHandler) As Button
        Dim b As Button = UITheme.MakeButton(text.Replace("&&", "&"), color, width, 38)
        AddHandler b.Click, handler
        Return b
    End Function

    Private Function StatCard(caption As String, valueColor As Color,
                              ByRef valueLabel As Label) As Panel
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Margin = New Padding(5, 0, 5, 8)

        valueLabel = New Label()
        valueLabel.Dock = DockStyle.Fill
        valueLabel.AutoSize = False
        valueLabel.TextAlign = ContentAlignment.MiddleCenter
        valueLabel.Font = UITheme.FontBold(24)
        valueLabel.ForeColor = valueColor
        valueLabel.Text = "0"

        Dim cap As New Label()
        cap.Dock = DockStyle.Bottom
        cap.Height = 36
        cap.AutoSize = False
        cap.TextAlign = ContentAlignment.MiddleCenter
        cap.Font = UITheme.FontRegular(11)
        cap.ForeColor = UITheme.TextMuted
        cap.Text = caption

        card.Controls.Add(valueLabel)
        card.Controls.Add(cap)
        Return card
    End Function

    ' ---------- Data ----------

    Private Sub LoadData()
        Try
            grid.DataSource = EmployeeService.GetEmployees(txtSearch.Text)
            Dim s As Integer() = EmployeeService.GetStats()
            lblTotal.Text = s(0).ToString()
            lblActive.Text = s(1).ToString()
            lblInactive.Text = s(2).ToString()
        Catch ex As Exception
            MessageBox.Show("The employees could not be loaded: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GridCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 OrElse e.Value Is Nothing Then Return
        If grid.Columns(e.ColumnIndex).DataPropertyName = "StatusText" Then
            e.CellStyle.ForeColor = If(e.Value.ToString() = "Active", UITheme.SuccessGreen, UITheme.DangerRed)
            e.CellStyle.Font = boldFont
        End If
    End Sub

    Private Function Selected(ByRef id As Integer, ByRef name As String) As Boolean
        If grid.CurrentRow Is Nothing Then
            MessageBox.Show("Select an employee in the list first.", "Employees",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If
        Dim drv As DataRowView = TryCast(grid.CurrentRow.DataBoundItem, DataRowView)
        If drv Is Nothing Then Return False
        id = Convert.ToInt32(drv("EmployeeId"))
        name = drv("FullName").ToString()
        Return True
    End Function

    ' ---------- Actions ----------

    Private Sub RegisterClick(sender As Object, e As EventArgs)
        Dim name As String = ""
        Dim code As String = ""
        If Not AskDetails("Register Employee",
                          "Enter the employee's name and choose a unique 4-digit access code. " &
                          "They will log in with this code only.", True, name, code) Then Return

        Dim err As String = EmployeeService.RegisterEmployee(name, code)
        If err <> "" Then
            MessageBox.Show(err, "Register Employee", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        LoadData()
        MessageBox.Show("Employee '" & name & "' registered." & vbCrLf & vbCrLf &
                        "Give them their code privately. It cannot be displayed again.",
                        "Register Employee", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ResetClick(sender As Object, e As EventArgs)
        Dim id As Integer
        Dim nm As String = ""
        If Not Selected(id, nm) Then Return

        Dim dummy As String = ""
        Dim code As String = ""
        If Not AskDetails("Reset Access Code",
                          "Choose a new unique 4-digit code for " & nm & ". The old code stops working immediately.",
                          False, dummy, code) Then Return

        Dim err As String = EmployeeService.ResetAccessCode(id, code)
        If err <> "" Then
            MessageBox.Show(err, "Reset Code", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        MessageBox.Show("The access code for " & nm & " has been reset.", "Reset Code",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ActivateClick(sender As Object, e As EventArgs)
        ChangeStatus(True)
    End Sub

    Private Sub DeactivateClick(sender As Object, e As EventArgs)
        ChangeStatus(False)
    End Sub

    Private Sub ChangeStatus(active As Boolean)
        Dim id As Integer
        Dim nm As String = ""
        If Not Selected(id, nm) Then Return

        Dim question As String = If(active,
            "Activate " & nm & "? They will be able to log in again.",
            "Deactivate " & nm & "? They will no longer be able to log in. " &
            "Their history and records are kept.")
        If MessageBox.Show(question, "Confirm", MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question) <> DialogResult.Yes Then Return

        Dim err As String = EmployeeService.SetActive(id, active)
        If err <> "" Then
            MessageBox.Show(err, "Employees", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        LoadData()
    End Sub

    Private Sub PerformanceClick(sender As Object, e As EventArgs)
        Dim id As Integer
        Dim nm As String = ""
        If Not Selected(id, nm) Then Return
        Dim text As String = EmployeeService.GetPerformanceSummary(id)
        If text = "" Then Return
        MessageBox.Show(text, "Performance: " & nm, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ActivityClick(sender As Object, e As EventArgs)
        Dim id As Integer
        Dim nm As String = ""
        If Not Selected(id, nm) Then Return

        Using f As New Form()
            f.Text = "Activity history: " & nm
            f.WindowState = FormWindowState.Maximized
            f.BackColor = UITheme.LightGrey
            Dim page As New ActivityHistoryPage(id)
            page.Dock = DockStyle.Fill
            f.Controls.Add(page)
            f.ShowDialog(Me.FindForm())
        End Using
    End Sub

    ' ---------- Dialog for name + code ----------

    Private Sub DigitsOnly(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then e.Handled = True
    End Sub

    Private Function AskDetails(title As String, note As String, askName As Boolean,
                                ByRef nameOut As String, ByRef codeOut As String) As Boolean
        Dim accepted As Boolean = False
        Dim nameText As String = ""
        Dim codeText As String = ""

        Using dlg As New Form()
            dlg.Text = title
            dlg.StartPosition = FormStartPosition.CenterParent
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.MaximizeBox = False
            dlg.MinimizeBox = False
            dlg.ShowInTaskbar = False
            dlg.BackColor = UITheme.CardWhite

            Dim y As Integer = 18
            Dim lblTitle As Label = UITheme.MakeLabel(title, 15, True, UITheme.Navy)
            lblTitle.Left = 24
            lblTitle.Top = y
            y += 50
            Dim lblNote As Label = UITheme.MakeLabel(note, 10.5F, False, UITheme.TextDark)
            lblNote.MaximumSize = New Size(472, 0)
            lblNote.Left = 24
            lblNote.Top = y
            y += 88
            dlg.Controls.Add(lblTitle)
            dlg.Controls.Add(lblNote)

            Dim txtName As TextBox = Nothing
            If askName Then
                Dim l As Label = UITheme.MakeLabel("Employee name *", 11, True, UITheme.TextDark)
                l.Left = 24
                l.Top = y
                txtName = UITheme.MakeTextBox(472, 12)
                txtName.MaxLength = 80
                txtName.Left = 24
                txtName.Top = y + 30
                dlg.Controls.Add(l)
                dlg.Controls.Add(txtName)
                y += 84
            End If

            Dim lc As Label = UITheme.MakeLabel("Access code (4 digits) *", 11, True, UITheme.TextDark)
            lc.Left = 24
            lc.Top = y
            Dim txtCode As TextBox = UITheme.MakeTextBox(472, 12)
            txtCode.MaxLength = 4
            txtCode.UseSystemPasswordChar = True
            txtCode.Left = 24
            txtCode.Top = y + 30
            AddHandler txtCode.KeyPress, AddressOf DigitsOnly
            y += 84

            Dim lf As Label = UITheme.MakeLabel("Confirm access code *", 11, True, UITheme.TextDark)
            lf.Left = 24
            lf.Top = y
            Dim txtConfirm As TextBox = UITheme.MakeTextBox(472, 12)
            txtConfirm.MaxLength = 4
            txtConfirm.UseSystemPasswordChar = True
            txtConfirm.Left = 24
            txtConfirm.Top = y + 30
            AddHandler txtConfirm.KeyPress, AddressOf DigitsOnly
            y += 84

            Dim lblError As Label = UITheme.MakeLabel("", 10.5F, True, UITheme.DangerRed)
            lblError.MaximumSize = New Size(472, 0)
            lblError.Left = 24
            lblError.Top = y
            y += 40

            Dim btnOk As Button = UITheme.MakeButton("Save", UITheme.SuccessGreen, 200, 48)
            btnOk.Left = 24
            btnOk.Top = y
            AddHandler btnOk.Click,
                Sub(s, ev)
                    If askName AndAlso txtName.Text.Trim() = "" Then
                        lblError.Text = "Enter the employee's name."
                        Return
                    End If
                    If Not Regex.IsMatch(txtCode.Text, "^\d{4}$") Then
                        lblError.Text = "The code must be exactly 4 digits."
                        Return
                    End If
                    If txtCode.Text <> txtConfirm.Text Then
                        lblError.Text = "The two codes do not match."
                        Return
                    End If
                    dlg.DialogResult = DialogResult.OK
                End Sub

            Dim btnCancel As Button = UITheme.MakeButton("Cancel", UITheme.LightGrey, 140, 48)
            btnCancel.ForeColor = UITheme.TextDark
            btnCancel.Left = 240
            btnCancel.Top = y
            btnCancel.DialogResult = DialogResult.Cancel
            dlg.CancelButton = btnCancel

            dlg.Controls.AddRange(New Control() {lc, txtCode, lf, txtConfirm, lblError, btnOk, btnCancel})
            dlg.ClientSize = New Size(520, y + 76)

            If dlg.ShowDialog(Me.FindForm()) = DialogResult.OK Then
                nameText = If(txtName Is Nothing, "", txtName.Text.Trim())
                codeText = txtCode.Text
                accepted = True
            End If
        End Using

        nameOut = nameText
        codeOut = codeText
        Return accepted
    End Function

End Class