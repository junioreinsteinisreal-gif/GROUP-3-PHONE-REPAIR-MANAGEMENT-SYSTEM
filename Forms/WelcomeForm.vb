Public Class WelcomeForm

    Private card As Panel
    Private autoCloseTimer As System.Windows.Forms.Timer
    Private startupError As String = ""
    Private startupWarnings As String = ""

    Private Sub WelcomeForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Phone Repair Management System"
        Me.WindowState = FormWindowState.Maximized
        Me.MinimumSize = New Size(800, 600)
        Me.DoubleBuffered = True

        ' The database is essential. If it fails we do NOT close here (that crashes);
        ' the Shown event below shows the reason and closes safely.
        Try
            DatabaseManager.Initialize()
        Catch ex As Exception
            startupError = "The database could not be opened:" & vbCrLf & vbCrLf & ex.Message
            Return
        End Try

        ' Everything below is optional: a failure is reported but never stops the program.
        Try
            SchemaUpdates.Apply()
        Catch ex As Exception
            startupWarnings &= "Database update (reviews and messages): " & ex.Message & vbCrLf & vbCrLf
        End Try

        Try
            EndOfDayService.AutoCloseIfDue(True)
        Catch ex As Exception
            startupWarnings &= "Automatic day close: " & ex.Message & vbCrLf & vbCrLf
        End Try

        Try
            NotificationService.ProcessPendingAsync()
        Catch ex As Exception
            startupWarnings &= "Customer messages: " & ex.Message & vbCrLf & vbCrLf
        End Try

        Try
            Dim serverError As String = ReviewServer.StartServer()
            If serverError <> "" Then
                startupWarnings &= "Review website could not start: " & serverError & vbCrLf & vbCrLf
            End If
        Catch ex As Exception
            startupWarnings &= "Review website: " & ex.Message & vbCrLf & vbCrLf
        End Try

        autoCloseTimer = New System.Windows.Forms.Timer()
        autoCloseTimer.Interval = 30000
        AddHandler autoCloseTimer.Tick,
            Sub(s, ev)
                Try
                    EndOfDayService.AutoCloseIfDue(False)
                Catch
                End Try
            End Sub
        autoCloseTimer.Start()

        UITheme.ApplyBackdrop(Me)
        Try
            BuildScreen()
        Catch ex As Exception
            startupError = "The welcome screen could not be built:" & vbCrLf & vbCrLf & ex.Message
        End Try
    End Sub

    Private Sub WelcomeForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        If startupError <> "" Then
            MessageBox.Show(startupError, "Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
            Return
        End If
        If startupWarnings <> "" Then
            MessageBox.Show("The program started, but some parts had a problem:" & vbCrLf & vbCrLf &
                            startupWarnings & "Please send a photo of this message.",
                            "Startup Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            startupWarnings = ""
        End If
    End Sub

    Private Sub WelcomeForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        If autoCloseTimer IsNot Nothing Then
            autoCloseTimer.Stop()
            autoCloseTimer.Dispose()
        End If
        Try
            ReviewServer.StopServer()
        Catch
        End Try
    End Sub

    Private Sub BuildScreen()
        Dim bizName As String = Convert.ToString(DatabaseManager.ExecuteScalar(
            "SELECT Value FROM Settings WHERE Key = 'BusinessName'"))

        card = UITheme.MakeCard(480, 420)

        Dim lblBiz As Label = UITheme.MakeCenteredLabel(bizName, 13, False,
                                                        UITheme.RoyalBlue, card.Width, 30)
        lblBiz.Top = 35

        Dim lblTitle As Label = UITheme.MakeCenteredLabel("Repair Management System", 20, True,
                                                          UITheme.Navy, card.Width, 44)
        lblTitle.Top = 68

        Dim lblSub As Label = UITheme.MakeCenteredLabel("Please choose how you want to sign in",
                                                        11, False, UITheme.TextMuted, card.Width, 28)
        lblSub.Top = 120

        Dim btnAdmin As Button = UITheme.MakeButton("Administrator Login", UITheme.Navy, 340, 56)
        btnAdmin.Left = 70
        btnAdmin.Top = 175
        AddHandler btnAdmin.Click, AddressOf OpenAdminLogin

        Dim btnEmployee As Button = UITheme.MakeButton("Employee Login", UITheme.RoyalBlue, 340, 56)
        btnEmployee.Left = 70
        btnEmployee.Top = 250
        AddHandler btnEmployee.Click, AddressOf OpenEmployeeLogin

        Dim btnExit As Button = UITheme.MakeButton("Exit", UITheme.LightGrey, 340, 46)
        btnExit.ForeColor = UITheme.TextDark
        btnExit.Left = 70
        btnExit.Top = 335
        AddHandler btnExit.Click, Sub(s, ev) Me.Close()

        card.Controls.AddRange(New Control() {lblBiz, lblTitle, lblSub, btnAdmin, btnEmployee, btnExit})
        Me.Controls.Add(card)
        UITheme.KeepCentered(Me, card)
    End Sub

    Private Sub OpenAdminLogin(sender As Object, ev As EventArgs)
        Using f As New AdminLoginForm()
            Me.Hide()
            If f.ShowDialog() = DialogResult.OK Then OpenMain()
            Me.Show()
        End Using
    End Sub

    Private Sub OpenEmployeeLogin(sender As Object, ev As EventArgs)
        Using f As New EmployeeLoginForm()
            Me.Hide()
            If f.ShowDialog() = DialogResult.OK Then OpenMain()
            Me.Show()
        End Using
    End Sub

    ''' <summary>Opens the main window. Returns here when the user logs out.</summary>
    Private Sub OpenMain()
        Using m As New MainForm()
            m.ShowDialog()
        End Using
    End Sub

End Class
