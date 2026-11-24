''' <summary>Payments list with receipt preview and printing (employees only).</summary>
Public Class PaymentsPage
    Inherits UserControl

    Private txtSearch As TextBox
    Private grid As DataGridView
    Private lblCount As Label
    Private chkThermal As CheckBox

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can use payments and receipts.")
        End If

        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)
        BuildPage()
        LoadData(0)
    End Sub

    ' Docking: grid (Fill) first, then the footer (Bottom), then the toolbar (Top).
    Private Sub BuildPage()
        grid = New DataGridView()
        grid.Dock = DockStyle.Fill
        UIHelpers.StyleGrid(grid)
        UIHelpers.AddTextColumn(grid, "ReceiptNumber", "Receipt No.", 85)
        UIHelpers.AddTextColumn(grid, "PaidAt", "Date && Time", 130)
        UIHelpers.AddTextColumn(grid, "RepairCode", "Repair Code", 85)
        UIHelpers.AddTextColumn(grid, "CustomerName", "Customer", 140)
        UIHelpers.AddTextColumn(grid, "Amount", "Amount", 90, True, "N2")
        UIHelpers.AddTextColumn(grid, "PaymentMethod", "Method", 100)
        UIHelpers.AddTextColumn(grid, "Employee", "Received By", 120)
        UIHelpers.AddTextColumn(grid, "PrintCount", "Times Printed", 80, True)
        AddHandler grid.CellDoubleClick, Sub(s, ev) If ev.RowIndex >= 0 Then ShowReceipt(False)

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
        toolbar.WrapContents = True
        toolbar.Padding = New Padding(0, 0, 0, 6)

        txtSearch = UITheme.MakeTextBox(280, 11)
        txtSearch.MaxLength = 60
        txtSearch.PlaceholderText = "Code, customer, phone or receipt no."
        txtSearch.Margin = New Padding(5, 5, 5, 3)
        AddHandler txtSearch.KeyDown,
            Sub(s, ev)
                If ev.KeyCode = Keys.Enter Then
                    ev.SuppressKeyPress = True
                    LoadData(0)
                End If
            End Sub

        Dim btnSearch As Button = UITheme.MakeButton("Search", UITheme.RoyalBlue, 100, 34)
        AddHandler btnSearch.Click, Sub(s, ev) LoadData(0)

        Dim btnClear As Button = UITheme.MakeButton("Clear", UITheme.TextMuted, 100, 34)
        AddHandler btnClear.Click,
            Sub(s, ev)
                txtSearch.Clear()
                LoadData(0)
            End Sub

        Dim btnPreview As Button = UITheme.MakeButton("Preview Receipt", UITheme.SuccessGreen, 170, 34)
        AddHandler btnPreview.Click, Sub(s, ev) ShowReceipt(False)

        Dim btnPrint As Button = UITheme.MakeButton("Print Receipt", UITheme.Navy, 150, 34)
        AddHandler btnPrint.Click, Sub(s, ev) ShowReceipt(True)

        chkThermal = New CheckBox()
        chkThermal.Text = "80 mm thermal paper"
        chkThermal.AutoSize = True
        chkThermal.Font = UITheme.FontBold(10)
        chkThermal.Margin = New Padding(12, 8, 0, 0)

        toolbar.Controls.AddRange(New Control() {txtSearch, btnSearch, btnClear, btnPreview,
                                                 btnPrint, chkThermal})

        Me.Controls.Add(grid)
        Me.Controls.Add(footer)
        Me.Controls.Add(toolbar)
    End Sub

    Private Sub LoadData(selectId As Long)
        Try
            Dim dt As DataTable = ReceiptService.GetPayments(txtSearch.Text)
            grid.DataSource = dt
            lblCount.Text = If(dt.Rows.Count = 0, "No payments found.",
                               dt.Rows.Count & " payments shown (newest first, up to 300). " &
                               "Select one, then Preview or Print.")
            If selectId > 0 Then
                For Each r As DataGridViewRow In grid.Rows
                    Dim drv As DataRowView = TryCast(r.DataBoundItem, DataRowView)
                    If drv IsNot Nothing AndAlso Convert.ToInt64(drv("PaymentId")) = selectId Then
                        grid.CurrentCell = r.Cells(0)
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("The payments could not be loaded: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function SelectedPaymentId() As Long
        If grid.CurrentRow Is Nothing Then Return 0
        Dim drv As DataRowView = TryCast(grid.CurrentRow.DataBoundItem, DataRowView)
        If drv Is Nothing Then Return 0
        Return Convert.ToInt64(drv("PaymentId"))
    End Function

    Private Sub ShowReceipt(print As Boolean)
        Dim id As Long = SelectedPaymentId()
        If id = 0 Then
            MessageBox.Show("Select a payment in the list first.", "Receipt",
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim data As ReceiptData = Nothing
        Try
            data = ReceiptService.GetReceipt(id)
        Catch ex As Exception
            MessageBox.Show("The receipt could not be prepared: " & ex.Message, "Receipt",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        If data Is Nothing Then
            MessageBox.Show("The repair for this payment is no longer available.", "Receipt",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            Dim printer As New ReceiptPrinter(data, chkThermal.Checked)
            If print Then
                printer.PrintNow(Me.FindForm())
            Else
                printer.Preview(Me.FindForm())
            End If
        Catch ex As Exception
            MessageBox.Show("Printing problem: " & ex.Message & vbCrLf &
                            "Check that a printer is installed.", "Receipt",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

        LoadData(id)
    End Sub

End Class