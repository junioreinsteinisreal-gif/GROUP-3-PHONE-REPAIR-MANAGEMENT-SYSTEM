Public Class AdminLoginForm

    Private card As Panel
    Private lblTitle As Label
    Private lblSub As Label
    Private txtUsername As TextBox
    Private txtPassword As TextBox
    Private lblConfirm As Label
    Private txtConfirm As TextBox
    Private lblMessage As Label
    Private btnPrimary As Button
    Private btnBack As Button
    Private setupMode As Boolean = False

    Private Sub AdminLoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Administrator Login"
        Me.WindowState = FormWindowState.Maximized
        Me.MinimumSize = New Size(800, 600)
        Me.DoubleBuffered = True
        UITheme.ApplyBackdrop(Me)

        BuildScreen()
        setupMode = Not AuthService.AdminExists()
        ApplyMode()
        UITheme.KeepCentered(Me, card)
        txtUsername.Focus()
    End Sub

    Private Sub BuildScreen()
        card = UITheme.MakeCard(440, 500)

        lblTitle = UITheme.MakeCenteredLabel("", 18, True, UITheme.Navy, card.Width, 40)
        lblTitle.Top = 30
        lblSub = UITheme.MakeCenteredLabel("", 10, False, UITheme.TextMuted, card.Width, 24)
        lblSub.Top = 72

        Dim lblUser As Label = UITheme.MakeLabel("Username", 10, True, UITheme.TextDark)
        lblUser.Left = 40
        lblUser.Top = 120
        txtUsername = UITheme.MakeTextBox(360)
        txtUsername.Left = 40
        txtUsername.Top = 145

        Dim lblPass As Label = UITheme.MakeLabel("Password", 10, True, UITheme.TextDark)
        lblPass.Left = 40
        lblPass.Top = 195
        txtPassword = UITheme.MakeTextBox(360)
        txtPassword.Left = 40
        txtPassword.Top = 220
        txtPassword.UseSystemPasswordChar = True

        lblConfirm = UITheme.MakeLabel("Confirm password", 10, True, UITheme.TextDark)
        lblConfirm.Left = 40
        lblConfirm.Top = 270
        txtConfirm = UITheme.MakeTextBox(360)
        txtConfirm.Left = 40
        txtConfirm.Top = 295
        txtConfirm.UseSystemPasswordChar = True

        lblMessage = UITheme.MakeCenteredLabel("", 10, False, UITheme.DangerRed, 360, 40)
        lblMessage.Left = 40

        btnPrimary = UITheme.MakeButton("Log In", UITheme.Navy, 360, 48)
        btnPrimary.Left = 40
        AddHandler btnPrimary.Click, AddressOf PrimaryClick

        btnBack = UITheme.MakeButton("Back", UITheme.LightGrey, 360, 42)
        btnBack.ForeColor = UITheme.TextDark
        btnBack.Left = 40
        btnBack.DialogResult = DialogResult.Cancel

        card.Controls.AddRange(New Control() {lblTitle, lblSub, lblUser, txtUsername,
                                              lblPass, txtPassword, lblConfirm, txtConfirm,
                                              lblMessage, btnPrimary, btnBack})
        Me.Controls.Add(card)

        Me.AcceptButton = btnPrimary
        Me.CancelButton = btnBack
    End Sub

    ''' <summary>Switches between Log In and first-time Create Administrator layouts.</summary>
    Private Sub ApplyMode()
        lblConfirm.Visible = setupMode
        txtConfirm.Visible = setupMode

        Dim y As Integer = 270
        If setupMode Then
            lblTitle.Text = "Create Administrator"
            lblSub.Text = "First-time setup: choose your administrator login"
            btnPrimary.Text = "Create Administrator"
            y = 345
        Else
            lblTitle.Text = "Administrator Login"
            lblSub.Text = "Sign in to manage the system"
            btnPrimary.Text = "Log In"
        End If

        lblMessage.Top = y
        btnPrimary.Top = y + 50
        btnBack.Top = y + 108
        card.Height = btnBack.Bottom + 30
        UITheme.CenterIn(card, Me)
    End Sub

    Private Sub PrimaryClick(sender As Object, ev As EventArgs)
        lblMessage.ForeColor = UITheme.DangerRed
        lblMessage.Text = ""
        If setupMode Then
            DoCreateAdmin()
        Else
            DoLogin()
        End If
    End Sub

    Private Sub DoLogin()
        Dim msg As String = ""
        If AuthService.LoginAdministrator(txtUsername.Text, txtPassword.Text, msg) Then
            Me.DialogResult = DialogResult.OK
        Else
            lblMessage.Text = msg
            txtPassword.Clear()
            txtPassword.Focus()
        End If
    End Sub

    Private Sub DoCreateAdmin()
        If txtPassword.Text <> txtConfirm.Text Then
            lblMessage.Text = "The passwords do not match."
            Return
        End If

        Dim err As String = AuthService.CreateAdministrator(txtUsername.Text, txtPassword.Text)
        If err <> "" Then
            lblMessage.Text = err
            Return
        End If

        setupMode = False
        txtPassword.Clear()
        txtConfirm.Clear()
        ApplyMode()
        lblMessage.ForeColor = UITheme.SuccessGreen
        lblMessage.Text = "Administrator created. Please log in."
        txtPassword.Focus()
    End Sub

End Class