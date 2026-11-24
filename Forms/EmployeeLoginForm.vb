Public Class EmployeeLoginForm

    Private card As Panel
    Private txtCode As TextBox
    Private lblMessage As Label
    Private btnLogin As Button
    Private btnBack As Button

    Private Sub EmployeeLoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Employee Login"
        Me.WindowState = FormWindowState.Maximized
        Me.MinimumSize = New Size(800, 600)
        Me.DoubleBuffered = True
        UITheme.ApplyBackdrop(Me)

        BuildScreen()
        UITheme.KeepCentered(Me, card)
        txtCode.Focus()
    End Sub

    Private Sub BuildScreen()
        card = UITheme.MakeCard(440, 380)

        Dim lblTitle As Label = UITheme.MakeCenteredLabel("Employee Login", 18, True,
                                                          UITheme.Navy, card.Width, 40)
        lblTitle.Top = 30
        Dim lblSub As Label = UITheme.MakeCenteredLabel("Enter your 4-digit access code",
                                                        10, False, UITheme.TextMuted, card.Width, 24)
        lblSub.Top = 75

        txtCode = UITheme.MakeTextBox(260, 28)
        txtCode.Left = 90
        txtCode.Top = 120
        txtCode.MaxLength = 4
        txtCode.TextAlign = HorizontalAlignment.Center
        txtCode.UseSystemPasswordChar = True
        AddHandler txtCode.KeyPress, AddressOf CodeKeyPress

        lblMessage = UITheme.MakeCenteredLabel("", 10, False, UITheme.DangerRed, 360, 40)
        lblMessage.Left = 40
        lblMessage.Top = 195

        btnLogin = UITheme.MakeButton("Log In", UITheme.RoyalBlue, 360, 48)
        btnLogin.Left = 40
        btnLogin.Top = 245
        AddHandler btnLogin.Click, AddressOf LoginClick

        btnBack = UITheme.MakeButton("Back", UITheme.LightGrey, 360, 42)
        btnBack.ForeColor = UITheme.TextDark
        btnBack.Left = 40
        btnBack.Top = 305
        btnBack.DialogResult = DialogResult.Cancel

        card.Controls.AddRange(New Control() {lblTitle, lblSub, txtCode, lblMessage,
                                              btnLogin, btnBack})
        Me.Controls.Add(card)

        Me.AcceptButton = btnLogin
        Me.CancelButton = btnBack
    End Sub

    ''' <summary>Only digits are allowed in the code box.</summary>
    Private Sub CodeKeyPress(sender As Object, ev As KeyPressEventArgs)
        If Not Char.IsControl(ev.KeyChar) AndAlso Not Char.IsDigit(ev.KeyChar) Then
            ev.Handled = True
        End If
    End Sub

    Private Sub LoginClick(sender As Object, ev As EventArgs)
        lblMessage.Text = ""
        Dim msg As String = ""
        If AuthService.LoginEmployee(txtCode.Text, msg) Then
            Me.DialogResult = DialogResult.OK
        Else
            lblMessage.Text = msg
            txtCode.Clear()
            txtCode.Focus()
        End If
    End Sub

End Class