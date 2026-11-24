Imports System.Globalization

''' <summary>Form for registering a new repair (employees only).</summary>
Public Class RegisterRepairPage
    Inherits UserControl

    Private txtName As TextBox
    Private txtPhone As TextBox
    Private txtEmail As TextBox
    Private cmbBrand As ComboBox
    Private txtModel As TextBox
    Private txtProblem As TextBox
    Private dtpReceived As DateTimePicker
    Private dtpExpected As DateTimePicker
    Private txtCost As TextBox
    Private txtPaid As TextBox
    Private cmbMethod As ComboBox
    Private txtEmployee As TextBox
    Private lblMessage As Label
    Private btnRegister As Button
    Private btnClear As Button

    Public Sub New()
        ' Permission check inside the page itself, not just a hidden button.
        SessionManager.RequireLogin()
        If Not SessionManager.IsEmployee Then
            Throw New UnauthorizedAccessException("Only employees can register repairs.")
        End If

        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)
        Me.AutoScrollMinSize = New Size(900, 760)

        BuildForm()
    End Sub

    ' ---------- Building the form ----------

    ''' <summary>A label on top of an input control, sized to fill its table cell.</summary>
    Private Function Field(labelText As String, ctrl As Control,
                           Optional fillHeight As Boolean = False) As Panel
        Dim pnl As New Panel()
        pnl.Dock = DockStyle.Fill
        pnl.Margin = New Padding(12, 0, 12, 8)

        ctrl.Dock = If(fillHeight, DockStyle.Fill, DockStyle.Top)

        Dim lbl As New Label()
        lbl.AutoSize = False
        lbl.Dock = DockStyle.Top
        lbl.Height = 34
        lbl.TextAlign = ContentAlignment.BottomLeft
        lbl.Font = UITheme.FontBold(11)
        lbl.ForeColor = UITheme.TextDark
        lbl.Text = labelText

        pnl.Controls.Add(ctrl)   ' added first, so it docks last
        pnl.Controls.Add(lbl)    ' added second, so it docks first (at the top)
        Return pnl
    End Function

    Private Function MakeDatePicker(initial As Date) As DateTimePicker
        Dim dtp As New DateTimePicker()
        dtp.Font = UITheme.FontRegular(12)
        dtp.Format = DateTimePickerFormat.Custom
        dtp.CustomFormat = "dd MMM yyyy"
        dtp.Value = initial
        Return dtp
    End Function

    Private Sub BuildForm()
        ' ----- Input controls -----
        txtName = UITheme.MakeTextBox(100, 12)
        txtName.MaxLength = 80

        txtPhone = UITheme.MakeTextBox(100, 12)
        txtPhone.MaxLength = 20

        txtEmail = UITheme.MakeTextBox(100, 12)
        txtEmail.MaxLength = 80

        cmbBrand = New ComboBox()
        cmbBrand.Font = UITheme.FontRegular(12)
        cmbBrand.DropDownStyle = ComboBoxStyle.DropDown
        cmbBrand.MaxLength = 40
        cmbBrand.Items.AddRange(New Object() {"Apple", "Samsung", "Tecno", "Infinix", "itel",
                                              "Huawei", "Xiaomi", "Oppo", "Vivo", "Nokia",
                                              "Google", "Other"})

        txtModel = UITheme.MakeTextBox(100, 12)
        txtModel.MaxLength = 60

        txtProblem = UITheme.MakeTextBox(100, 12)
        txtProblem.Multiline = True
        txtProblem.ScrollBars = ScrollBars.Vertical
        txtProblem.MaxLength = 500

        dtpReceived = MakeDatePicker(Date.Today)
        dtpExpected = MakeDatePicker(Date.Today.AddDays(3))

        txtCost = UITheme.MakeTextBox(100, 12)
        txtCost.MaxLength = 10
        txtCost.PlaceholderText = "0.00"
        AddHandler txtCost.KeyPress, AddressOf MoneyKeyPress

        txtPaid = UITheme.MakeTextBox(100, 12)
        txtPaid.MaxLength = 10
        txtPaid.PlaceholderText = "0.00"
        AddHandler txtPaid.KeyPress, AddressOf MoneyKeyPress

        cmbMethod = New ComboBox()
        cmbMethod.Font = UITheme.FontRegular(12)
        cmbMethod.DropDownStyle = ComboBoxStyle.DropDownList
        cmbMethod.Items.AddRange(RepairService.PaymentMethods)
        cmbMethod.SelectedIndex = 0

        txtEmployee = UITheme.MakeTextBox(100, 12)
        txtEmployee.ReadOnly = True
        txtEmployee.TabStop = False
        txtEmployee.BackColor = UITheme.LightGrey
        txtEmployee.Text = SessionManager.CurrentUserName

        ' ----- Two-column grid that fills the window -----
        Dim tlp As New TableLayoutPanel()
        tlp.Dock = DockStyle.Fill
        tlp.ColumnCount = 2
        tlp.RowCount = 7
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        tlp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        For i As Integer = 1 To 7
            tlp.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F / 7.0F))
        Next

        ' Left column first, then right, so the Tab key runs down each column.
        tlp.Controls.Add(Field("Customer name *", txtName), 0, 0)
        tlp.Controls.Add(Field("Customer phone number *", txtPhone), 0, 1)
        tlp.Controls.Add(Field("Phone brand *", cmbBrand), 0, 2)
        tlp.Controls.Add(Field("Phone model *", txtModel), 0, 3)
        Dim problemField As Panel = Field("Problem description *", txtProblem, True)
        tlp.Controls.Add(problemField, 0, 4)
        tlp.SetRowSpan(problemField, 3)

        tlp.Controls.Add(Field("Date received *", dtpReceived), 1, 0)
        tlp.Controls.Add(Field("Expected collection date *", dtpExpected), 1, 1)
        tlp.Controls.Add(Field("Estimated repair cost (" & UIHelpers.CurrencySymbol & ") *", txtCost), 1, 2)
        tlp.Controls.Add(Field("Initial payment (" & UIHelpers.CurrencySymbol & ")", txtPaid), 1, 3)
        tlp.Controls.Add(Field("Payment method", cmbMethod), 1, 4)
        tlp.Controls.Add(Field("Employee responsible", txtEmployee), 1, 5)
        tlp.Controls.Add(Field("Customer email (optional, for status updates)", txtEmail), 1, 6)

        ' ----- Bottom strip: message and buttons -----
        Dim bottom As New Panel()
        bottom.Dock = DockStyle.Bottom
        bottom.Height = 110

        lblMessage = New Label()
        lblMessage.AutoSize = False
        lblMessage.Dock = DockStyle.Top
        lblMessage.Height = 42
        lblMessage.Padding = New Padding(12, 0, 0, 0)
        lblMessage.TextAlign = ContentAlignment.MiddleLeft
        lblMessage.Font = UITheme.FontBold(11)
        lblMessage.ForeColor = UITheme.DangerRed

        btnRegister = UITheme.MakeButton("Register Repair", UITheme.SuccessGreen, 240, 54)
        btnRegister.Left = 12
        btnRegister.Top = 48
        AddHandler btnRegister.Click, AddressOf RegisterClick

        btnClear = UITheme.MakeButton("Clear Form", UITheme.LightGrey, 200, 54)
        btnClear.ForeColor = UITheme.TextDark
        btnClear.Left = 270
        btnClear.Top = 48
        AddHandler btnClear.Click, Sub(s, ev) ClearForm()

        bottom.Controls.Add(btnRegister)
        bottom.Controls.Add(btnClear)
        bottom.Controls.Add(lblMessage)

        ' ----- Card with title -----
        Dim lblHint As New Label()
        lblHint.AutoSize = False
        lblHint.Dock = DockStyle.Top
        lblHint.Height = 30
        lblHint.TextAlign = ContentAlignment.MiddleLeft
        lblHint.Font = UITheme.FontRegular(10)
        lblHint.ForeColor = UITheme.TextMuted
        lblHint.Text = "Fields marked * are required."

        Dim lblTitle As New Label()
        lblTitle.AutoSize = False
        lblTitle.Dock = DockStyle.Top
        lblTitle.Height = 54
        lblTitle.TextAlign = ContentAlignment.MiddleLeft
        lblTitle.Font = UITheme.FontBold(18)
        lblTitle.ForeColor = UITheme.Navy
        lblTitle.Text = "Register New Repair"

        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Padding = New Padding(25, 15, 25, 10)

        ' Docking order: Fill first, then Bottom, then the Top controls.
        card.Controls.Add(tlp)
        card.Controls.Add(bottom)
        card.Controls.Add(lblHint)
        card.Controls.Add(lblTitle)
        Me.Controls.Add(card)
    End Sub

    ' ---------- Input helpers ----------

    ''' <summary>Money boxes accept only digits and one decimal point.</summary>
    Private Sub MoneyKeyPress(sender As Object, e As KeyPressEventArgs)
        Dim tb As TextBox = DirectCast(sender, TextBox)
        If Char.IsControl(e.KeyChar) OrElse Char.IsDigit(e.KeyChar) Then Return
        If e.KeyChar = "."c AndAlso Not tb.Text.Contains(".") Then Return
        e.Handled = True
    End Sub

    Private Function TryMoney(text As String, required As Boolean,
                              ByRef value As Decimal) As Boolean
        Dim t As String = text.Trim()
        If t = "" Then
            value = 0D
            Return Not required
        End If
        Return Decimal.TryParse(t, NumberStyles.AllowDecimalPoint,
                                CultureInfo.InvariantCulture, value)
    End Function

    Private Sub ShowMessage(text As String, isError As Boolean)
        lblMessage.ForeColor = If(isError, UITheme.DangerRed, UITheme.SuccessGreen)
        lblMessage.Text = text
    End Sub

    ' ---------- Actions ----------

    Private Sub RegisterClick(sender As Object, e As EventArgs)
        ShowMessage("", False)

        Dim cost As Decimal
        Dim paid As Decimal
        If Not TryMoney(txtCost.Text, True, cost) Then
            ShowMessage("Enter a valid estimated cost, for example 350 or 350.50.", True)
            txtCost.Focus()
            Return
        End If
        If Not TryMoney(txtPaid.Text, False, paid) Then
            ShowMessage("Enter a valid initial payment, for example 100 or 100.50.", True)
            txtPaid.Focus()
            Return
        End If
        If txtEmail.Text.Trim() <> "" AndAlso Not NotificationService.IsValidEmail(txtEmail.Text) Then
            ShowMessage("Enter a valid email address, or leave it empty.", True)
            txtEmail.Focus()
            Return
        End If

        Dim code As String = ""
        Dim err As String = ""
        btnRegister.Enabled = False
        Try
            err = RepairService.RegisterRepair(txtName.Text, txtPhone.Text, cmbBrand.Text,
                                               txtModel.Text, txtProblem.Text, cost, paid,
                                               cmbMethod.Text, dtpReceived.Value.Date,
                                               dtpExpected.Value.Date, code)
        Catch ex As UnauthorizedAccessException
            err = ex.Message
        Finally
            btnRegister.Enabled = True
        End Try

        ' On failure the form is NOT cleared, so nothing has to be retyped.
        If err <> "" Then
            ShowMessage(err, True)
            Return
        End If

        ' Save the email first, then send the "we received your phone" message.
        NotificationService.SaveCustomerEmail(code, txtEmail.Text)
        NotificationService.OnStatusChanged(code, "Received")

        Dim extra As String = ""
        If paid > 0 Then
            extra = vbCrLf & "Initial payment recorded: " & UIHelpers.Money(CDbl(paid))
        End If
        MessageBox.Show("Repair registered successfully." & vbCrLf & vbCrLf &
                        "Repair code: " & code & extra & vbCrLf & vbCrLf &
                        "Give this code to the customer.",
                        "Repair Registered", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ClearForm()
        ShowMessage("Repair " & code & " registered.", False)
    End Sub

    Private Sub ClearForm()
        txtName.Clear()
        txtPhone.Clear()
        txtEmail.Clear()
        cmbBrand.SelectedIndex = -1
        cmbBrand.Text = ""
        txtModel.Clear()
        txtProblem.Clear()
        dtpReceived.Value = Date.Today
        dtpExpected.Value = Date.Today.AddDays(3)
        txtCost.Clear()
        txtPaid.Clear()
        cmbMethod.SelectedIndex = 0
        ShowMessage("", False)
        txtName.Focus()
    End Sub

End Class