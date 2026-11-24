''' <summary>Administrator: review website status, demo mode, test links and tip requests.</summary>
Public Class ReviewSitePage
    Inherits UserControl

    Private lblStatus As Label
    Private chkDemo As CheckBox
    Private txtPort As TextBox
    Private txtBase As TextBox
    Private lblSaveMsg As Label
    Private txtCode As TextBox
    Private lblLink As Label
    Private gridTips As DataGridView
    Private phoneLink As String = ""
    Private pcLink As String = ""
    Private ReadOnly boldFont As Font = UITheme.FontBold(10.5F)

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsAdmin Then
            Throw New UnauthorizedAccessException("Only administrators can manage the review website.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)
        BuildPage()
        LoadValues()
        RefreshStatus()
        LoadTips()
    End Sub

    Private Function Heading(text As String) As Label
        Dim l As Label = UITheme.MakeLabel(text, 14, True, UITheme.Navy)
        l.Margin = New Padding(3, 14, 3, 6)
        Return l
    End Function

    Private Function Note(text As String, height As Integer) As Label
        Dim l As New Label()
        l.AutoSize = False
        l.Width = 860
        l.Height = height
        l.Font = UITheme.FontRegular(10)
        l.ForeColor = UITheme.TextMuted
        l.Text = text
        l.Margin = New Padding(3, 0, 3, 6)
        Return l
    End Function

    Private Function Caption(text As String) As Label
        Dim l As Label = UITheme.MakeLabel(text, 10.5F, True, UITheme.TextDark)
        l.Margin = New Padding(3, 8, 3, 2)
        Return l
    End Function

    Private Sub BuildPage()
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Padding = New Padding(20)

        Dim flow As New FlowLayoutPanel()
        flow.Dock = DockStyle.Fill
        flow.FlowDirection = FlowDirection.TopDown
        flow.WrapContents = False
        flow.AutoScroll = True

        flow.Controls.Add(Heading("Customer Review Website"))
        lblStatus = New Label()
        lblStatus.AutoSize = False
        lblStatus.Width = 860
        lblStatus.Height = 70
        lblStatus.Font = UITheme.FontBold(10.5F)
        lblStatus.Margin = New Padding(3, 0, 3, 6)
        flow.Controls.Add(lblStatus)
        flow.Controls.Add(Note("Customers open the link from their Completed message on a phone connected to the same " &
                               "Wi-Fi as this computer, while this program is running. Reviews and tips arrive in " &
                               "Reviews && Tips automatically.", 56))
        Dim last As Label = DirectCast(flow.Controls(flow.Controls.Count - 1), Label)
        last.Text = last.Text.Replace("&&", "&")

        chkDemo = New CheckBox()
        chkDemo.Text = "Demo mode: simulate the MoMo prompt on the page (no real money)"
        chkDemo.AutoSize = True
        chkDemo.Font = UITheme.FontBold(11)
        chkDemo.Margin = New Padding(3, 8, 3, 4)
        flow.Controls.Add(chkDemo)

        txtPort = UITheme.MakeTextBox(160, 12)
        txtPort.MaxLength = 5
        txtBase = UITheme.MakeTextBox(520, 12)
        txtBase.MaxLength = 120
        flow.Controls.Add(Caption("Port (default 8080)"))
        flow.Controls.Add(txtPort)
        flow.Controls.Add(Caption("Public address (leave empty for now; used when the page is put on real hosting)"))
        flow.Controls.Add(txtBase)

        Dim btnSave As Button = UITheme.MakeButton("Save && Restart Website", UITheme.SuccessGreen, 280, 46)
        btnSave.Text = "Save & Restart Website"
        btnSave.Margin = New Padding(3, 12, 3, 4)
        AddHandler btnSave.Click, AddressOf SaveClick
        flow.Controls.Add(btnSave)

        lblSaveMsg = New Label()
        lblSaveMsg.AutoSize = False
        lblSaveMsg.Width = 860
        lblSaveMsg.Height = 44
        lblSaveMsg.Font = UITheme.FontBold(10.5F)
        flow.Controls.Add(lblSaveMsg)

        flow.Controls.Add(Heading("Try it with a completed repair"))
        txtCode = UITheme.MakeTextBox(220, 12)
        txtCode.MaxLength = 12
        txtCode.PlaceholderText = "PR-001"
        flow.Controls.Add(Caption("Repair code (must be Completed or Collected, with no review yet)"))
        flow.Controls.Add(txtCode)

        Dim btnRow As New FlowLayoutPanel()
        btnRow.AutoSize = True
        btnRow.AutoSizeMode = AutoSizeMode.GrowAndShrink
        btnRow.WrapContents = False
        btnRow.Margin = New Padding(0, 8, 0, 0)
        Dim btnOpen As Button = UITheme.MakeButton("Open on This PC", UITheme.RoyalBlue, 220, 44)
        AddHandler btnOpen.Click, AddressOf OpenClick
        Dim btnCopy As Button = UITheme.MakeButton("Copy Phone Link", UITheme.Navy, 220, 44)
        AddHandler btnCopy.Click, AddressOf CopyClick
        btnRow.Controls.Add(btnOpen)
        btnRow.Controls.Add(btnCopy)
        flow.Controls.Add(btnRow)

        lblLink = New Label()
        lblLink.AutoSize = False
        lblLink.Width = 860
        lblLink.Height = 56
        lblLink.Font = UITheme.FontRegular(10)
        lblLink.Margin = New Padding(3, 8, 3, 6)
        flow.Controls.Add(lblLink)

        flow.Controls.Add(Heading("Recent tip requests"))
        gridTips = New DataGridView()
        gridTips.Width = 860
        gridTips.Height = 260
        UIHelpers.StyleGrid(gridTips)
        gridTips.RowTemplate.Height = 36
        gridTips.ColumnHeadersHeight = 42
        UIHelpers.AddTextColumn(gridTips, "CreatedAt", "When", 120)
        UIHelpers.AddTextColumn(gridTips, "RepairCode", "Repair", 70)
        UIHelpers.AddTextColumn(gridTips, "Employee", "Employee", 110)
        UIHelpers.AddTextColumn(gridTips, "Amount", "Amount", 70, True, "N2")
        UIHelpers.AddTextColumn(gridTips, "Network", "Network", 80)
        UIHelpers.AddTextColumn(gridTips, "Status", "Status", 75)
        UIHelpers.AddTextColumn(gridTips, "Mode", "Mode", 60)
        AddHandler gridTips.CellFormatting,
            Sub(s, e)
                If e.RowIndex >= 0 AndAlso e.Value IsNot Nothing AndAlso
                   gridTips.Columns(e.ColumnIndex).DataPropertyName = "Status" Then
                    Dim v As String = e.Value.ToString()
                    e.CellStyle.ForeColor = If(v = "Paid", UITheme.SuccessGreen,
                                               If(v = "Pending", Color.FromArgb(176, 112, 0), UITheme.TextMuted))
                    e.CellStyle.Font = boldFont
                End If
            End Sub
        flow.Controls.Add(gridTips)

        Dim btnRefresh As Button = UITheme.MakeButton("Refresh", UITheme.Navy, 140, 40)
        btnRefresh.Margin = New Padding(3, 8, 3, 3)
        AddHandler btnRefresh.Click, Sub(s, e)
                                         RefreshStatus()
                                         LoadTips()
                                     End Sub
        flow.Controls.Add(btnRefresh)

        card.Controls.Add(flow)
        Me.Controls.Add(card)
    End Sub

    Private Sub LoadValues()
        chkDemo.Checked = ReviewServer.DemoMode()
        txtPort.Text = ReviewServer.ConfiguredPort().ToString()
        txtBase.Text = AdminToolsService.GetSetting("ReviewBaseUrl")
    End Sub

    Private Sub RefreshStatus()
        If ReviewServer.IsRunning Then
            lblStatus.ForeColor = UITheme.SuccessGreen
            lblStatus.Text = "Website is running." & vbCrLf & "Phones on the same Wi-Fi use: " &
                             ReviewServer.BaseUrl() & "/r/(repair code)?k=(PIN)" & vbCrLf &
                             If(ReviewServer.DemoMode(), "Demo mode is ON: the MoMo prompt is simulated.",
                                "Demo mode is OFF: tips are disabled until a live payment provider is connected.")
        Else
            lblStatus.ForeColor = UITheme.DangerRed
            lblStatus.Text = "Website is NOT running." & vbCrLf &
                             If(ReviewServer.LastError <> "", "Reason: " & ReviewServer.LastError,
                                "Click Save && Restart Website.").Replace("&&", "&") & vbCrLf &
                             "If the port is in use, choose another one, for example 8090."
        End If
    End Sub

    Private Sub LoadTips()
        Try
            gridTips.DataSource = ReviewServer.GetTipRequests()
        Catch
        End Try
    End Sub

    Private Sub SaveClick(sender As Object, e As EventArgs)
        Dim err As String = ReviewServer.SaveConfig(txtPort.Text, chkDemo.Checked, txtBase.Text)
        If err <> "" Then
            lblSaveMsg.ForeColor = UITheme.DangerRed
            lblSaveMsg.Text = err
            Return
        End If
        Dim startErr As String = ReviewServer.StartServer()
        RefreshStatus()
        If startErr <> "" Then
            lblSaveMsg.ForeColor = UITheme.DangerRed
            lblSaveMsg.Text = "Saved, but the website could not start: " & startErr
        Else
            lblSaveMsg.ForeColor = UITheme.SuccessGreen
            lblSaveMsg.Text = "Saved. The website was restarted."
        End If
    End Sub

    Private Function MakeLinks() As Boolean
        Dim err As String = ReviewServer.CreateTestLink(txtCode.Text, phoneLink, pcLink)
        If err <> "" Then
            lblLink.ForeColor = UITheme.DangerRed
            lblLink.Text = err
            Return False
        End If
        lblLink.ForeColor = UITheme.TextDark
        lblLink.Text = "Phone link: " & phoneLink & vbCrLf & "This PC: " & pcLink
        Return True
    End Function

    Private Sub OpenClick(sender As Object, e As EventArgs)
        If Not MakeLinks() Then Return
        If Not ReviewServer.IsRunning Then
            lblLink.ForeColor = UITheme.DangerRed
            lblLink.Text = "The website is not running. Click Save & Restart Website first."
            Return
        End If
        Try
            Process.Start(New ProcessStartInfo(pcLink) With {.UseShellExecute = True})
        Catch ex As Exception
            lblLink.ForeColor = UITheme.DangerRed
            lblLink.Text = "Could not open the browser: " & ex.Message
        End Try
    End Sub

    Private Sub CopyClick(sender As Object, e As EventArgs)
        If Not MakeLinks() Then Return
        Try
            Clipboard.SetText(phoneLink)
            lblLink.Text &= vbCrLf & "Phone link copied. Paste it into WhatsApp or a message to your phone."
        Catch
        End Try
    End Sub

End Class