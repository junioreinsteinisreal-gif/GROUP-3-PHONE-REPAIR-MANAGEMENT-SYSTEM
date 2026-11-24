Public Class MainForm
    Inherits Form

    ' Employees are logged out automatically at this time.
    Private Const AutoLogoutHour As Integer = 23
    Private Const AutoLogoutMinute As Integer = 30

    Private sidebar As Panel
    Private header As Panel
    Private contentHost As Panel
    Private lblPageTitle As Label
    Private lblUser As Label
    Private lblClock As Label
    Private clockTimer As System.Windows.Forms.Timer

    Private ReadOnly navItems As New List(Of NavItem)
    Private ReadOnly navButtons As New Dictionary(Of String, Button)
    Private currentKey As String = ""

    ' ---------- Startup ----------

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Defence in depth: the main window cannot open without a login.
        If Not SessionManager.IsLoggedIn Then
            Me.Close()
            Return
        End If

        Me.Text = "Phone Repair Management System"
        Me.WindowState = FormWindowState.Maximized
        Me.MinimumSize = New Size(1000, 650)
        Me.BackColor = UITheme.LightGrey
        Me.Font = UITheme.FontRegular(10)

        BuildNavItems()

        BuildLayout()

        clockTimer = New System.Windows.Forms.Timer()
        clockTimer.Interval = 1000
        AddHandler clockTimer.Tick, AddressOf ClockTick
        clockTimer.Start()
        ClockTick(Nothing, EventArgs.Empty)

        Navigate("dashboard")
    End Sub

    ' ---------- Menu definition ----------

    Private Sub BuildNavItems()
        '                       key          text                      admin  employee stage
        navItems.Add(New NavItem("dashboard", "Dashboard", True, True, "Stage 11"))
        navItems.Add(New NavItem("register", "Register Repair", False, True, "Stage 6"))
        navItems.Add(New NavItem("repairs", "View Repairs", True, True, "Stage 6"))
        navItems.Add(New NavItem("update", "Search && Update", False, True, "Stage 7"))
        navItems.Add(New NavItem("deleterepair", "Delete Repair", False, True, "Stage 12"))
        navItems.Add(New NavItem("payments", "Payments && Receipts", False, True, "Stage 8"))
        navItems.Add(New NavItem("myactivity", "My Activity", False, True, "Stage 10"))
        navItems.Add(New NavItem("reminders", "Collection Reminders", False, True, "Stage 13"))
        navItems.Add(New NavItem("employees", "Employee Management", True, False, "Stage 9"))
        navItems.Add(New NavItem("activityall", "Activity History", True, False, "Stage 10"))
        navItems.Add(New NavItem("audit", "Administrator History", True, False, "Stage 12"))
        navItems.Add(New NavItem("terminated", "Terminated Repairs", True, False, "Stage 7"))
        navItems.Add(New NavItem("feedback", "My Reviews && Tips", False, True, "this stage"))
        navItems.Add(New NavItem("reviewsadmin", "Reviews && Tips", True, False, "this stage"))
        navItems.Add(New NavItem("notifications", "Customer Messages", True, False, "this stage"))
        navItems.Add(New NavItem("reviewsite", "Review Website", True, False, "this stage"))
        navItems.Add(New NavItem("endofday", "End of Day", True, True, "Stage 14"))
        navItems.Add(New NavItem("monthly", "Monthly Sales", True, False, "Stage 15"))
        navItems.Add(New NavItem("backup", "Settings && Backup", True, False, "Stage 17"))
    End Sub

    ' ---------- Layout ----------
    ' Docking order: the Fill control is added first, then Top, then Left,
    ' so the sidebar gets the full height and the header spans the rest.

    Private Sub BuildLayout()
        contentHost = New Panel()
        contentHost.Dock = DockStyle.Fill
        contentHost.BackColor = UITheme.LightGrey

        BuildHeader()
        BuildSidebar()

        Me.Controls.Add(contentHost)
        Me.Controls.Add(header)
        Me.Controls.Add(sidebar)
    End Sub

    Private Sub BuildHeader()
        header = New Panel()
        header.Dock = DockStyle.Top
        header.Height = 70
        header.BackColor = UITheme.CardWhite

        lblPageTitle = New Label()
        lblPageTitle.Dock = DockStyle.Fill
        lblPageTitle.AutoSize = False
        lblPageTitle.Padding = New Padding(25, 0, 0, 0)
        lblPageTitle.TextAlign = ContentAlignment.MiddleLeft
        lblPageTitle.Font = UITheme.FontBold(16)
        lblPageTitle.ForeColor = UITheme.Navy

        Dim rightPanel As New Panel()
        rightPanel.Dock = DockStyle.Right
        rightPanel.Width = 440
        rightPanel.Padding = New Padding(0, 0, 25, 0)

        lblUser = New Label()
        lblUser.Dock = DockStyle.Fill
        lblUser.AutoSize = False
        lblUser.TextAlign = ContentAlignment.BottomRight
        lblUser.Font = UITheme.FontBold(11)
        lblUser.ForeColor = UITheme.TextDark
        lblUser.Text = SessionManager.CurrentUserName & "  |  " & SessionManager.CurrentRole

        lblClock = New Label()
        lblClock.Dock = DockStyle.Bottom
        lblClock.Height = 32
        lblClock.AutoSize = False
        lblClock.TextAlign = ContentAlignment.TopRight
        lblClock.Font = UITheme.FontRegular(9)
        lblClock.ForeColor = UITheme.TextMuted

        rightPanel.Controls.Add(lblUser)
        rightPanel.Controls.Add(lblClock)

        header.Controls.Add(lblPageTitle)
        header.Controls.Add(rightPanel)
    End Sub

    Private Sub BuildSidebar()
        sidebar = New Panel()
        sidebar.Dock = DockStyle.Left
        sidebar.Width = 270
        sidebar.BackColor = UITheme.Navy

        ' Menu list (Fill, added first)
        Dim flow As New FlowLayoutPanel()
        flow.Dock = DockStyle.Fill
        flow.FlowDirection = FlowDirection.TopDown
        flow.WrapContents = False
        flow.AutoScroll = True
        flow.BackColor = UITheme.Navy
        flow.Padding = New Padding(0)

        For Each item As NavItem In navItems
            If item.IsAllowedFor(SessionManager.CurrentRole) Then
                Dim btn As Button = MakeNavButton(item)
                navButtons(item.Key) = btn
                flow.Controls.Add(btn)
            End If
        Next

        ' Logout area (Bottom)
        Dim bottom As New Panel()
        bottom.Dock = DockStyle.Bottom
        bottom.Height = 84
        bottom.BackColor = UITheme.Navy

        Dim btnLogout As Button = UITheme.MakeButton("Log Out", UITheme.DangerRed, 220, 48)
        btnLogout.Left = 25
        btnLogout.Top = 18
        AddHandler btnLogout.Click, AddressOf LogoutClick
        bottom.Controls.Add(btnLogout)

        ' Brand area (Top)
        Dim brand As New Panel()
        brand.Dock = DockStyle.Top
        brand.Height = 110
        brand.BackColor = UITheme.Navy

        Dim bizName As String = Convert.ToString(DatabaseManager.ExecuteScalar(
            "SELECT Value FROM Settings WHERE Key = 'BusinessName'"))

        Dim lblBiz As Label = UITheme.MakeCenteredLabel(bizName, 13, True,
                                                        Color.White, 270, 40)
        lblBiz.Top = 22
        Dim lblSub As Label = UITheme.MakeCenteredLabel("Repair Manager", 9, False,
                                                        Color.FromArgb(160, 185, 235), 270, 28)
        lblSub.Top = 66
        brand.Controls.Add(lblBiz)
        brand.Controls.Add(lblSub)

        sidebar.Controls.Add(flow)
        sidebar.Controls.Add(bottom)
        sidebar.Controls.Add(brand)
    End Sub

    Private Function MakeNavButton(item As NavItem) As Button
        Dim btn As New Button()
        btn.Text = item.Text.Replace("&&", "&")
        btn.Tag = item.Key
        btn.Size = New Size(270, 54)
        btn.Margin = New Padding(0)
        btn.Padding = New Padding(24, 0, 0, 0)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.BackColor = UITheme.Navy
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(25, 45, 100)
        btn.ForeColor = Color.White
        btn.Font = UITheme.FontBold(11)
        btn.TextAlign = ContentAlignment.MiddleLeft
        btn.Cursor = Cursors.Hand
        btn.UseMnemonic = False
        AddHandler btn.Click, AddressOf NavButtonClick
        Return btn
    End Function

    ' ---------- Navigation ----------

    Private Sub NavButtonClick(sender As Object, e As EventArgs)
        Navigate(CStr(DirectCast(sender, Button).Tag))
    End Sub

    Private Function FindItem(key As String) As NavItem
        For Each item As NavItem In navItems
            If item.Key = key Then Return item
        Next
        Return Nothing
    End Function

    Private Sub Navigate(key As String)
        Dim item As NavItem = FindItem(key)
        If item Is Nothing Then Return

        ' Permission check (second layer: buttons are already hidden by role)
        If Not item.IsAllowedFor(SessionManager.CurrentRole) Then
            ActivityLogger.Log("AccessDenied", "Tried to open " & item.Text & ".", Nothing, False)
            MessageBox.Show("You do not have permission to open " & item.Text & ".",
                            "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Build the page (third layer: the page checks access itself)
        Dim page As Control
        Try
            Select Case key
                Case "dashboard"
                    page = New HomePage()
                Case "register"
                    page = New RegisterRepairPage()
                Case "repairs"
                    page = New RepairsPage()
                Case "update"
                    page = New UpdateRepairPage()
                Case "payments"
                    page = New PaymentsPage()
                Case "terminated"
                    page = New TerminatedRepairsPage()
                Case Else
                    page = NavConfig.CreatePage(key, item)
            End Select
        Catch ex As UnauthorizedAccessException
            MessageBox.Show(ex.Message, "Access Denied",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End Try

        ' Swap the content
        Dim oldPages As New List(Of Control)
        For Each c As Control In contentHost.Controls
            oldPages.Add(c)
        Next
        contentHost.Controls.Clear()
        For Each c As Control In oldPages
            c.Dispose()
        Next

        page.Dock = DockStyle.Fill
        contentHost.Controls.Add(page)

        lblPageTitle.Text = item.Text.Replace("&&", "&")
        currentKey = key
        HighlightActive(key)
    End Sub

    Private Sub HighlightActive(key As String)
        For Each kv As KeyValuePair(Of String, Button) In navButtons
            If kv.Key = key Then
                kv.Value.BackColor = UITheme.RoyalBlue
                kv.Value.FlatAppearance.MouseOverBackColor = UITheme.RoyalBlue
            Else
                kv.Value.BackColor = UITheme.Navy
                kv.Value.FlatAppearance.MouseOverBackColor = Color.FromArgb(25, 45, 100)
            End If
        Next
    End Sub

    ' ---------- Clock and automatic logout ----------

    Private blockCheckTicks As Integer = 0

    Private Sub ClockTick(sender As Object, e As EventArgs)
        lblClock.Text = DateTime.Now.ToString("dddd, dd MMMM yyyy   HH:mm:ss")

        If Not SessionManager.IsLoggedIn Then Return

        Dim pastCutoff As Boolean = DateTime.Now.TimeOfDay >= EndOfDayService.CloseTime

        ' The day closes automatically at 11:30 PM (while an administrator stays logged in too).
        If pastCutoff Then EndOfDayService.AutoCloseIfDue(SessionManager.IsEmployee)

        If Not SessionManager.IsEmployee Then Return

        Dim reason As String = ""
        If pastCutoff Then
            reason = "The day was closed automatically at 11:30 PM. " &
                     "Employees can log in again from 6:00 AM."
        Else
            ' If the administrator closed the day, log the employee out within ~15 seconds.
            blockCheckTicks += 1
            If blockCheckTicks >= 15 Then
                blockCheckTicks = 0
                reason = EndOfDayService.EmployeeLoginBlockReason()
            End If
        End If

        If reason <> "" Then
            clockTimer.Stop()
            AuthService.Logout("Day closed")
            MessageBox.Show(reason & vbCrLf & vbCrLf & "You have been logged out.",
                            "Day Closed", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        End If
    End Sub

    ' ---------- Logout ----------

    Private Sub LogoutClick(sender As Object, e As EventArgs)
        If MessageBox.Show("Do you want to log out?", "Log Out",
                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            clockTimer.Stop()
            AuthService.Logout("Manual")
            Me.Close()
        End If
    End Sub

    ''' <summary>If the window is closed with the X button, still end the session.</summary>
    Private Sub MainForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        If clockTimer IsNot Nothing Then
            clockTimer.Stop()
            clockTimer.Dispose()
        End If
        If SessionManager.IsLoggedIn Then AuthService.Logout("Window closed")
    End Sub

End Class