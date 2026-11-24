Imports System.Globalization

''' <summary>Search for a repair, then update it or terminate it as beyond repair.</summary>
Public Class UpdateRepairPage
    Inherits UserControl

    ' "Unrepairable" is not selectable here: employees use the Terminate button.
    Private ReadOnly StatusChoices As String() =
        {"Received", "In Progress", "Completed", "Collected"}

    ' Search side
    Private txtSearch As TextBox
    Private grid As DataGridView
    Private lblResultCount As Label

    ' Details side
    Private lblCode As Label
    Private lblCustomer As Label
    Private lblDevice As Label
    Private lblDates As Label
    Private lblCost As Label
    Private lblPaid As Label
    Private lblBalance As Label
    Private txtProblem As TextBox

    ' Update fields
    Private cmbStatus As ComboBox
    Private dtpExpected As DateTimePicker
    Private txtPayment As TextBox
    Private cmbMethod As ComboBox
    Private lblMessage As Label
    Private btnSave As Button
    Private btnClear As Button
    Private btnTerminate As Button

    Private currentCode As String = ""
    Private currentInfo As RepairInfo = Nothing
    Private reasonDraft As String = ""
    Private binding As Boolean = False


    Public Sub New()
        ' Permission check inside the page itself, not just a hidden button.
        SessionManager.RequireLogin()
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can update repairs.")
        End If

        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)
        Me.AutoScrollMinSize = New Size(1000, 780)

        BuildPage()
        ClearDetails()
    End Sub

    ' ---------- Building the page ----------

    Private Sub BuildPage()
        Dim root As New TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.ColumnCount = 2
        root.RowCount = 1
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 42.0F))
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 58.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        root.Controls.Add(BuildSearchCard(), 0, 0)
        root.Controls.Add(BuildUpdateCard(), 1, 0)
        Me.Controls.Add(root)
    End Sub

    Private Function MakeSectionTitle(text As String) As Label
        Dim lbl As New Label()
        lbl.AutoSize = False
        lbl.Dock = DockStyle.Top
        lbl.Height = 50
        lbl.TextAlign = ContentAlignment.MiddleLeft
        lbl.Font = UITheme.FontBold(15)
        lbl.ForeColor = UITheme.Navy
        lbl.Text = text
        Return lbl
    End Function

    Private Function BuildSearchCard() As Panel
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Padding = New Padding(20)
        card.Margin = New Padding(0, 0, 10, 0)

        ' Results grid
        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        UIHelpers.AddTextColumn(grid, "RepairCode", "Code", 80)
        UIHelpers.AddTextColumn(grid, "CustomerName", "Customer", 130)
        UIHelpers.AddTextColumn(grid, "CustomerPhone", "Phone", 105)
        UIHelpers.AddTextColumn(grid, "Status", "Status", 90)
        UIHelpers.AddTextColumn(grid, "Balance", "Balance", 80, True, "N2")
        AddHandler grid.SelectionChanged, AddressOf GridSelectionChanged

        lblResultCount = New Label()
        lblResultCount.AutoSize = False
        lblResultCount.Dock = DockStyle.Bottom
        lblResultCount.Height = 32
        lblResultCount.TextAlign = ContentAlignment.MiddleLeft
        lblResultCount.Font = UITheme.FontRegular(10)
        lblResultCount.ForeColor = UITheme.TextMuted

        Dim lblHint As New Label()
        lblHint.AutoSize = False
        lblHint.Dock = DockStyle.Top
        lblHint.Height = 32
        lblHint.TextAlign = ContentAlignment.MiddleLeft
        lblHint.Font = UITheme.FontRegular(9.5F)
        lblHint.ForeColor = UITheme.TextMuted
        lblHint.Text = "Search by repair code, customer, phone, brand or model."

        ' Search row: text box + button
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
        txtSearch.PlaceholderText = "Type and press Enter"
        AddHandler txtSearch.KeyDown, AddressOf SearchKeyDown

        Dim btnSearch As Button = UITheme.MakeButton("Search", UITheme.RoyalBlue, 120, 42)
        btnSearch.Margin = New Padding(0, 3, 0, 0)
        AddHandler btnSearch.Click, Sub(s, ev) RunSearch()

        searchRow.Controls.Add(txtSearch, 0, 0)
        searchRow.Controls.Add(btnSearch, 1, 0)

        ' Docking order: Fill first, then Bottom, then Top controls (last = top-most)
        card.Controls.Add(grid)
        card.Controls.Add(lblResultCount)
        card.Controls.Add(lblHint)
        card.Controls.Add(searchRow)
        card.Controls.Add(MakeSectionTitle("Find Repair"))
        Return card
    End Function

    Private Function MakeValueLabel() As Label
        Dim lbl As New Label()
        lbl.AutoSize = False
        lbl.TextAlign = ContentAlignment.MiddleLeft
        lbl.Font = UITheme.FontBold(11)
        lbl.ForeColor = UITheme.TextDark
        lbl.AutoEllipsis = True
        Return lbl
    End Function

    Private Sub AddDetailRow(tbl As TableLayoutPanel, row As Integer,
                             key As String, value As Control)
        Dim k As New Label()
        k.AutoSize = False
        k.Dock = DockStyle.Fill
        k.TextAlign = ContentAlignment.MiddleLeft
        k.Font = UITheme.FontRegular(10.5F)
        k.ForeColor = UITheme.TextMuted
        k.Text = key
        tbl.Controls.Add(k, 0, row)

        value.Dock = DockStyle.Fill
        tbl.Controls.Add(value, 1, row)
    End Sub

    ''' <summary>A label on top of an input control, sized to fill its table cell.</summary>
    Private Function Field(labelText As String, ctrl As Control) As Panel
        Dim pnl As New Panel()
        pnl.Dock = DockStyle.Fill
        pnl.Margin = New Padding(0, 0, 12, 6)

        ctrl.Dock = DockStyle.Top

        Dim lbl As New Label()
        lbl.AutoSize = False
        lbl.Dock = DockStyle.Top
        lbl.Height = 34
        lbl.TextAlign = ContentAlignment.BottomLeft
        lbl.Font = UITheme.FontBold(11)
        lbl.ForeColor = UITheme.TextDark
        lbl.Text = labelText

        pnl.Controls.Add(ctrl)
        pnl.Controls.Add(lbl)
        Return pnl
    End Function

    Private Function BuildUpdateCard() As Panel
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Padding = New Padding(20)

        ' ----- Read-only details -----
        lblCode = MakeValueLabel()
        lblCode.Font = UITheme.FontBold(13)
        lblCode.ForeColor = UITheme.Navy
        lblCustomer = MakeValueLabel()
        lblDevice = MakeValueLabel()
        lblDates = MakeValueLabel()
        lblCost = MakeValueLabel()
        lblPaid = MakeValueLabel()
        lblBalance = MakeValueLabel()

        txtProblem = UITheme.MakeTextBox(100, 11)
        txtProblem.Multiline = True
        txtProblem.ReadOnly = True
        txtProblem.TabStop = False
        txtProblem.ScrollBars = ScrollBars.Vertical
        txtProblem.BackColor = UITheme.LightGrey
        txtProblem.Margin = New Padding(3, 4, 3, 4)

        Dim details As New TableLayoutPanel()
        details.Dock = DockStyle.Fill
        details.ColumnCount = 2
        details.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 160.0F))
        details.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        details.RowCount = 8
        For i As Integer = 1 To 7
            details.RowStyles.Add(New RowStyle(SizeType.Absolute, 36.0F))
        Next
        details.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        AddDetailRow(details, 0, "Repair code", lblCode)
        AddDetailRow(details, 1, "Customer", lblCustomer)
        AddDetailRow(details, 2, "Device", lblDevice)
        AddDetailRow(details, 3, "Received / by", lblDates)
        AddDetailRow(details, 4, "Estimated cost", lblCost)
        AddDetailRow(details, 5, "Total paid", lblPaid)
        AddDetailRow(details, 6, "Balance", lblBalance)
        AddDetailRow(details, 7, "Problem", txtProblem)

        ' ----- Update fields (2 x 2) -----
        cmbStatus = New ComboBox()
        cmbStatus.Font = UITheme.FontRegular(12)
        cmbStatus.DropDownStyle = ComboBoxStyle.DropDownList
        cmbStatus.Items.AddRange(StatusChoices)

        dtpExpected = New DateTimePicker()
        dtpExpected.Font = UITheme.FontRegular(12)
        dtpExpected.Format = DateTimePickerFormat.Custom
        dtpExpected.CustomFormat = "dd MMM yyyy"

        txtPayment = UITheme.MakeTextBox(100, 12)
        txtPayment.MaxLength = 10
        txtPayment.PlaceholderText = "0.00"
        AddHandler txtPayment.KeyPress, AddressOf MoneyKeyPress

        cmbMethod = New ComboBox()
        cmbMethod.Font = UITheme.FontRegular(12)
        cmbMethod.DropDownStyle = ComboBoxStyle.DropDownList
        cmbMethod.Items.AddRange(RepairService.PaymentMethods)
        cmbMethod.SelectedIndex = 0

        Dim updateTbl As New TableLayoutPanel()
        updateTbl.Dock = DockStyle.Bottom
        updateTbl.Height = 190
        updateTbl.ColumnCount = 2
        updateTbl.RowCount = 2
        updateTbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        updateTbl.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        updateTbl.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
        updateTbl.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0F))
        updateTbl.Controls.Add(Field("New status", cmbStatus), 0, 0)
        updateTbl.Controls.Add(Field("Expected collection date", dtpExpected), 1, 0)
        updateTbl.Controls.Add(Field("Payment amount (" & UIHelpers.CurrencySymbol & ")", txtPayment), 0, 1)
        updateTbl.Controls.Add(Field("Payment method", cmbMethod), 1, 1)

        ' ----- Message and buttons -----
        lblMessage = New Label()
        lblMessage.AutoSize = False
        lblMessage.Dock = DockStyle.Fill
        lblMessage.TextAlign = ContentAlignment.MiddleLeft
        lblMessage.Font = UITheme.FontBold(10.5F)
        lblMessage.AutoEllipsis = True

        btnSave = UITheme.MakeButton("Save Update", UITheme.SuccessGreen, 190, 54)
        btnSave.Margin = New Padding(0, 0, 12, 0)
        AddHandler btnSave.Click, AddressOf SaveClick

        btnClear = UITheme.MakeButton("Clear", UITheme.LightGrey, 120, 54)
        btnClear.ForeColor = UITheme.TextDark
        btnClear.Margin = New Padding(0, 0, 12, 0)
        AddHandler btnClear.Click, Sub(s, ev) ClearAll()

        btnTerminate = UITheme.MakeButton("Terminate Repair", UITheme.DangerRed, 210, 54)
        btnTerminate.Margin = New Padding(0)
        AddHandler btnTerminate.Click, AddressOf TerminateClick

        Dim buttonRow As New FlowLayoutPanel()
        buttonRow.Dock = DockStyle.Bottom
        buttonRow.Height = 62
        buttonRow.FlowDirection = FlowDirection.LeftToRight
        buttonRow.WrapContents = False
        buttonRow.Controls.Add(btnSave)
        buttonRow.Controls.Add(btnClear)
        buttonRow.Controls.Add(btnTerminate)

        Dim bottom As New Panel()
        bottom.Dock = DockStyle.Bottom
        bottom.Height = 130
        bottom.Controls.Add(lblMessage)
        bottom.Controls.Add(buttonRow)

        ' Docking order: Fill first, then Bottom (last added = lowest), then Top
        card.Controls.Add(details)
        card.Controls.Add(updateTbl)
        card.Controls.Add(bottom)
        card.Controls.Add(MakeSectionTitle("Repair Details && Update"))
        Return card
    End Function

    ' ---------- Small helpers ----------

    Private Sub ShowMessage(text As String, color As Color)
        lblMessage.ForeColor = color
        lblMessage.Text = text
    End Sub

    ''' <summary>Money boxes accept only digits and one decimal point.</summary>
    Private Sub MoneyKeyPress(sender As Object, e As KeyPressEventArgs)
        Dim tb As TextBox = DirectCast(sender, TextBox)
        If Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar) Then Return
        If e.KeyChar = "."c AndAlso Not tb.Text.Contains(".") Then Return
        e.Handled = True
    End Sub

    Private Function TryMoney(text As String, ByRef value As Decimal) As Boolean
        Dim t As String = text.Trim()
        If t = "" Then
            value = 0D
            Return True
        End If
        Return Decimal.TryParse(t, NumberStyles.AllowDecimalPoint,
                                CultureInfo.InvariantCulture, value)
    End Function

    Private Sub SetEditEnabled(enabled As Boolean)
        cmbStatus.Enabled = enabled
        dtpExpected.Enabled = enabled
        txtPayment.Enabled = enabled
        cmbMethod.Enabled = enabled
        btnSave.Enabled = enabled
        btnTerminate.Enabled = enabled
    End Sub

    ' ---------- Search ----------

    Private Sub SearchKeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            RunSearch()
        End If
    End Sub

    Private Sub RunSearch()
        Dim q As String = txtSearch.Text.Trim()
        If q = "" Then
            ShowMessage("Enter a repair code, customer name, phone, brand or model to search.",
                        UITheme.DangerRed)
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

            ClearDetails()
            ShowMessage("", UITheme.TextDark)
            ActivityLogger.Log("SearchRepairs", "Searched repairs for '" & q & "'.")

            If dt.Rows.Count = 0 Then
                lblResultCount.Text = "No repairs found."
            ElseIf dt.Rows.Count = 1 Then
                lblResultCount.Text = "1 repair found"
                ' A single result opens straight away.
                grid.CurrentCell = grid.Rows(0).Cells(0)
            Else
                lblResultCount.Text = dt.Rows.Count & " repairs found. Click one to open it."
            End If
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
        LoadRepair(drv("RepairCode").ToString())
    End Sub

    ' ---------- Showing one repair ----------

    Private Sub LoadRepair(code As String)
        If code = currentCode Then Return

        Dim info As RepairInfo = Nothing
        Try
            info = RepairUpdateService.FindRepair(code)
        Catch ex As Exception
            ShowMessage("The repair could not be loaded: " & ex.Message, UITheme.DangerRed)
            Return
        End Try

        If info Is Nothing Then
            ClearDetails()
            ShowMessage("That repair could not be found.", UITheme.DangerRed)
            Return
        End If

        ShowMessage("", UITheme.TextDark)
        currentCode = info.RepairCode
        currentInfo = info
        reasonDraft = ""

        lblCode.Text = info.RepairCode & "   (" & info.Status & ")"
        lblCustomer.Text = info.CustomerName & "   |   " & info.CustomerPhone
        lblDevice.Text = info.Brand & " " & info.Model
        lblDates.Text = info.DateReceived.ToString("dd MMM yyyy") & "   |   " & info.EmployeeName
        lblCost.Text = UIHelpers.Money(CDbl(info.EstimatedCost))
        lblPaid.Text = UIHelpers.Money(CDbl(info.TotalPaid))
        lblBalance.Text = UIHelpers.Money(CDbl(info.Balance))
        lblBalance.ForeColor = If(info.Balance > 0D, UITheme.DangerRed, UITheme.SuccessGreen)
        txtProblem.Text = info.Problem

        cmbStatus.SelectedIndex = -1
        cmbStatus.SelectedItem = info.Status
        dtpExpected.Value = info.ExpectedCollection
        txtPayment.Clear()
        cmbMethod.SelectedIndex = 0

        If info.Status = "Unrepairable" Then
            SetEditEnabled(False)
            ShowMessage(TerminationService.GetTerminationSummary(info.RepairCode),
                        UITheme.DangerRed)
        ElseIf info.Status = "Collected" Then
            SetEditEnabled(False)
            ShowMessage("Collected on " & info.CollectedAt &
                        ". No further changes are allowed.", UITheme.RoyalBlue)
        Else
            SetEditEnabled(True)
            If info.Balance <= 0D Then
                txtPayment.Enabled = False
                cmbMethod.Enabled = False
            End If
            ' Only phones still being worked on can be terminated.
            btnTerminate.Enabled = (info.Status = "Received" OrElse info.Status = "In Progress")
        End If

        ActivityLogger.Log("ViewRepair", "Opened repair " & info.RepairCode &
                           " on the update screen.", info.RepairCode)
    End Sub

    Private Sub ClearDetails()
        currentCode = ""
        currentInfo = Nothing
        reasonDraft = ""
        lblCode.Text = ""
        lblCustomer.Text = ""
        lblDevice.Text = ""
        lblDates.Text = ""
        lblCost.Text = ""
        lblPaid.Text = ""
        lblBalance.Text = ""
        lblBalance.ForeColor = UITheme.TextDark
        txtProblem.Clear()
        cmbStatus.SelectedIndex = -1
        dtpExpected.Value = Date.Today
        txtPayment.Clear()
        cmbMethod.SelectedIndex = 0
        SetEditEnabled(False)
    End Sub

    ''' <summary>Clears EVERYTHING: search box, results, details, status and payment.</summary>
    Private Sub ClearAll()
        txtSearch.Clear()
        binding = True
        grid.DataSource = Nothing
        binding = False
        lblResultCount.Text = ""
        ClearDetails()
        ShowMessage("", UITheme.TextDark)
        txtSearch.Focus()
    End Sub

    ' ---------- Saving ----------

    Private Sub SaveClick(sender As Object, e As EventArgs)
        ShowMessage("", UITheme.TextDark)

        If currentCode = "" Then
            ShowMessage("Search for a repair and select it first.", UITheme.DangerRed)
            txtSearch.Focus()
            Return
        End If
        If cmbStatus.SelectedIndex < 0 Then
            ShowMessage("Choose a status.", UITheme.DangerRed)
            Return
        End If

        Dim payment As Decimal
        If Not TryMoney(txtPayment.Text, payment) Then
            ShowMessage("Enter a valid payment, for example 100 or 100.50.", UITheme.DangerRed)
            txtPayment.Focus()
            Return
        End If

        Dim newStatus As String = cmbStatus.Text
        If newStatus = "Collected" Then
            If MessageBox.Show("Mark repair " & currentCode & " as Collected?" & vbCrLf & vbCrLf &
                               "This is final. The repair cannot be changed afterwards.",
                               "Confirm Collection", MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If
        End If

        Dim code As String = currentCode
        Dim err As String = ""
        btnSave.Enabled = False
        Try
            err = RepairUpdateService.UpdateRepair(code, newStatus, payment,
                                                   cmbMethod.Text, dtpExpected.Value.Date)
        Catch ex As UnauthorizedAccessException
            err = ex.Message
        Finally
            btnSave.Enabled = True
        End Try

        ' On failure the form is NOT cleared, so nothing has to be retyped.
        If err <> "" Then
            ShowMessage(err, UITheme.DangerRed)
            Return
        End If

        Dim extra As String = ""
        If payment > 0D Then
            extra = vbCrLf & "Payment recorded: " & UIHelpers.Money(CDbl(payment))
        End If
        MessageBox.Show("Repair " & code & " updated successfully." & vbCrLf & vbCrLf &
                        "New status: " & newStatus & extra,
                        "Repair Updated", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ClearAll()
        ShowMessage("Repair " & code & " updated.", UITheme.SuccessGreen)
    End Sub

    ' ---------- Terminating ----------

    Private Sub TerminateClick(sender As Object, e As EventArgs)
        ShowMessage("", UITheme.TextDark)

        If currentCode = "" OrElse currentInfo Is Nothing Then
            ShowMessage("Search for a repair and select it first.", UITheme.DangerRed)
            txtSearch.Focus()
            Return
        End If

        Dim code As String = currentCode
        Dim reason As String = ""
        Dim refund As Decimal = 0D
        Dim method As String = "Cash"
        If Not AskTermination(currentInfo, reason, refund, method) Then Return   ' cancelled

        Dim err As String = ""
        btnTerminate.Enabled = False
        Try
            err = TerminationService.TerminateRepair(code, reason, refund, method)
        Catch ex As UnauthorizedAccessException
            err = ex.Message
        Finally
            btnTerminate.Enabled = True
        End Try

        ' On failure nothing is cleared and the typed reason is kept for the next try.
        If err <> "" Then
            reasonDraft = reason
            ShowMessage(err, UITheme.DangerRed)
            Return
        End If

        Dim refundText As String = ""
        If refund > 0D Then
            refundText = vbCrLf & vbCrLf & "GIVE THE CUSTOMER A REFUND OF " &
                         UIHelpers.Money(CDbl(refund)) & " (" & method & ")."
        End If
        MessageBox.Show("Repair " & code & " has been terminated." & refundText & vbCrLf & vbCrLf &
                        "The administrator can now see your reason.",
                        "Repair Terminated", MessageBoxButtons.OK, MessageBoxIcon.Information)

        ClearAll()
        ShowMessage("Repair " & code & " terminated.", UITheme.SuccessGreen)
    End Sub

    ''' <summary>Pop-up asking for the reason and refund. Returns False if cancelled.</summary>
    Private Function AskTermination(info As RepairInfo, ByRef reasonOut As String,
                                    ByRef refundOut As Decimal,
                                    ByRef methodOut As String) As Boolean
        Dim accepted As Boolean = False
        Dim reasonText As String = ""
        Dim refundValue As Decimal = 0D
        Dim methodText As String = "Cash"

        Using dlg As New Form()
            dlg.Text = "Terminate Repair " & info.RepairCode
            dlg.StartPosition = FormStartPosition.CenterParent
            dlg.FormBorderStyle = FormBorderStyle.FixedDialog
            dlg.MaximizeBox = False
            dlg.MinimizeBox = False
            dlg.ShowInTaskbar = False
            dlg.ClientSize = New Size(640, 560)
            dlg.BackColor = UITheme.CardWhite

            Dim lblTitle As Label = UITheme.MakeLabel(
                "Terminate " & info.RepairCode & " (" & info.Brand & " " & info.Model & ")",
                15, True, UITheme.DangerRed)
            lblTitle.Left = 24
            lblTitle.Top = 16

            Dim lblInfo As Label = UITheme.MakeLabel(
                "Use this only when the phone is beyond repair. The repair will be closed " &
                "and nothing more will be owed. The administrator will see your reason. " &
                "This cannot be undone.",
                10.5F, False, UITheme.TextDark)
            lblInfo.MaximumSize = New Size(592, 0)
            lblInfo.Left = 24
            lblInfo.Top = 62

            Dim lblReason As Label = UITheme.MakeLabel(
                "Reason for termination * (at least " & TerminationService.MinReasonLength &
                " characters)", 11, True, UITheme.TextDark)
            lblReason.Left = 24
            lblReason.Top = 150

            Dim txtReason As TextBox = UITheme.MakeTextBox(592, 12)
            txtReason.Multiline = True
            txtReason.ScrollBars = ScrollBars.Vertical
            txtReason.MaxLength = TerminationService.MaxReasonLength
            txtReason.Left = 24
            txtReason.Top = 182
            txtReason.Height = 120
            txtReason.Text = reasonDraft

            Dim lblRefund As Label = UITheme.MakeLabel(
                "Refund to give the customer (" & UIHelpers.CurrencySymbol & ")",
                11, True, UITheme.TextDark)
            lblRefund.Left = 24
            lblRefund.Top = 318

            Dim txtRefund As TextBox = UITheme.MakeTextBox(280, 12)
            txtRefund.MaxLength = 10
            txtRefund.Left = 24
            txtRefund.Top = 350
            txtRefund.Text = info.TotalPaid.ToString("0.00", CultureInfo.InvariantCulture)
            AddHandler txtRefund.KeyPress, AddressOf MoneyKeyPress

            Dim lblRefundMethod As Label = UITheme.MakeLabel("Refund method", 11, True, UITheme.TextDark)
            lblRefundMethod.Left = 330
            lblRefundMethod.Top = 318

            Dim cmbRefundMethod As New ComboBox()
            cmbRefundMethod.Font = UITheme.FontRegular(12)
            cmbRefundMethod.DropDownStyle = ComboBoxStyle.DropDownList
            cmbRefundMethod.Items.AddRange(RepairService.PaymentMethods)
            cmbRefundMethod.SelectedIndex = 0
            cmbRefundMethod.Left = 330
            cmbRefundMethod.Top = 350
            cmbRefundMethod.Width = 286

            Dim lblPaidNote As Label = UITheme.MakeLabel(
                "The customer has paid " & UIHelpers.Money(CDbl(info.TotalPaid)) &
                ". Enter 0 if no refund is due.", 10, False, UITheme.TextMuted)
            lblPaidNote.MaximumSize = New Size(592, 0)
            lblPaidNote.Left = 24
            lblPaidNote.Top = 396

            Dim lblError As Label = UITheme.MakeLabel("", 10.5F, True, UITheme.DangerRed)
            lblError.MaximumSize = New Size(592, 0)
            lblError.Left = 24
            lblError.Top = 432

            Dim btnConfirm As Button = UITheme.MakeButton("Terminate Repair",
                                                          UITheme.DangerRed, 250, 50)
            btnConfirm.Left = 24
            btnConfirm.Top = 490
            AddHandler btnConfirm.Click,
                Sub(s, ev)
                    If txtReason.Text.Trim().Length < TerminationService.MinReasonLength Then
                        lblError.Text = "Please type a reason of at least " &
                                        TerminationService.MinReasonLength & " characters."
                        txtReason.Focus()
                        Return
                    End If
                    Dim amt As Decimal
                    If Not TryMoney(txtRefund.Text, amt) Then
                        lblError.Text = "Enter a valid refund amount, for example 100 or 100.50."
                        txtRefund.Focus()
                        Return
                    End If
                    If amt > info.TotalPaid Then
                        lblError.Text = "The refund cannot be more than the " &
                                        UIHelpers.Money(CDbl(info.TotalPaid)) & " the customer paid."
                        txtRefund.Focus()
                        Return
                    End If
                    dlg.DialogResult = DialogResult.OK
                End Sub

            Dim btnCancel As Button = UITheme.MakeButton("Cancel", UITheme.LightGrey, 160, 50)
            btnCancel.ForeColor = UITheme.TextDark
            btnCancel.Left = 290
            btnCancel.Top = 490
            btnCancel.DialogResult = DialogResult.Cancel
            dlg.CancelButton = btnCancel

            dlg.Controls.AddRange(New Control() {lblTitle, lblInfo, lblReason, txtReason,
                                                 lblRefund, txtRefund, lblRefundMethod,
                                                 cmbRefundMethod, lblPaidNote, lblError,
                                                 btnConfirm, btnCancel})

            If dlg.ShowDialog(Me.FindForm()) = DialogResult.OK Then
                reasonText = txtReason.Text.Trim()
                TryMoney(txtRefund.Text, refundValue)
                methodText = cmbRefundMethod.Text
                accepted = True
            Else
                reasonDraft = txtReason.Text.Trim()
            End If
        End Using

        reasonOut = reasonText
        refundOut = refundValue
        methodOut = methodText
        Return accepted
    End Function
End Class