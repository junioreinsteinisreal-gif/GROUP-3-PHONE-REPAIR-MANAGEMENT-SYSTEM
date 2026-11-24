Imports System.IO

''' <summary>Administrator-only business details and database backup.</summary>
Public Class SettingsPage
    Inherits UserControl

    Private txtName As TextBox
    Private txtPhone As TextBox
    Private txtAddress As TextBox
    Private lblSaveMsg As Label
    Private lblInfo As Label
    Private lblBackupMsg As Label

    Public Sub New()
        SessionManager.RequireLogin()
        If Not SessionManager.IsAdmin Then
            Throw New UnauthorizedAccessException("Only administrators can open settings and backup.")
        End If
        Me.BackColor = UITheme.LightGrey
        Me.Padding = New Padding(10)

        Dim root As New TableLayoutPanel()
        root.Dock = DockStyle.Fill
        root.ColumnCount = 2
        root.RowCount = 1
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        BuildBusinessCard(root)
        BuildBackupCard(root)
        Me.Controls.Add(root)
        RefreshInfo()
    End Sub

    ' ---------- Helpers ----------

    Private Sub FitWidths(flow As FlowLayoutPanel)
        Dim w As Integer = flow.ClientSize.Width - 8
        For Each c As Control In flow.Controls
            If Not (TypeOf c Is Label AndAlso DirectCast(c, Label).AutoSize) Then
                c.Width = Math.Max(100, w - c.Margin.Horizontal)
            End If
        Next
    End Sub

    Private Function AddCard(root As TableLayoutPanel, column As Integer,
                             title As String) As FlowLayoutPanel
        Dim card As New Panel()
        card.Dock = DockStyle.Fill
        card.BackColor = UITheme.CardWhite
        card.Padding = New Padding(25)
        card.Margin = New Padding(5)

        Dim flow As New FlowLayoutPanel()
        flow.Dock = DockStyle.Fill
        flow.FlowDirection = FlowDirection.TopDown
        flow.WrapContents = False
        flow.AutoScroll = True
        AddHandler flow.Resize, Sub(s, e) FitWidths(flow)

        card.Controls.Add(flow)
        root.Controls.Add(card, column, 0)

        Dim lbl As Label = UITheme.MakeLabel(title, 16, True, UITheme.Navy)
        lbl.Margin = New Padding(3, 0, 3, 12)
        flow.Controls.Add(lbl)
        Return flow
    End Function

    Private Function MakeNote(text As String, height As Integer) As Label
        Dim lbl As Label = UITheme.MakeLabel(text, 10, False, UITheme.TextMuted)
        lbl.AutoSize = False
        lbl.Height = height
        lbl.Margin = New Padding(3, 0, 3, 10)
        Return lbl
    End Function

    Private Function FieldLabel(text As String) As Label
        Dim l As Label = UITheme.MakeLabel(text, 11, True, UITheme.TextDark)
        l.Margin = New Padding(3, 8, 3, 2)
        Return l
    End Function

    Private Sub ShowMsg(lbl As Label, text As String, color As Color)
        lbl.ForeColor = color
        lbl.Text = text
    End Sub

    ' ---------- Business details ----------

    Private Sub BuildBusinessCard(root As TableLayoutPanel)
        Dim flow As FlowLayoutPanel = AddCard(root, 0, "Business Details")
        flow.Controls.Add(MakeNote("These details are printed on every receipt.", 28))

        txtName = UITheme.MakeTextBox(300, 12)
        txtName.MaxLength = 60
        txtName.Text = AdminToolsService.GetSetting("BusinessName")
        txtPhone = UITheme.MakeTextBox(300, 12)
        txtPhone.MaxLength = 30
        txtPhone.Text = AdminToolsService.GetSetting("BusinessPhone")
        txtAddress = UITheme.MakeTextBox(300, 12)
        txtAddress.MaxLength = 120
        txtAddress.Text = AdminToolsService.GetSetting("BusinessAddress")

        flow.Controls.Add(FieldLabel("Business name *"))
        flow.Controls.Add(txtName)
        flow.Controls.Add(FieldLabel("Phone number"))
        flow.Controls.Add(txtPhone)
        flow.Controls.Add(FieldLabel("Address"))
        flow.Controls.Add(txtAddress)

        Dim btnSave As Button = UITheme.MakeButton("Save Details", UITheme.SuccessGreen, 300, 50)
        btnSave.Margin = New Padding(3, 18, 3, 6)
        AddHandler btnSave.Click, AddressOf SaveClick
        flow.Controls.Add(btnSave)

        lblSaveMsg = MakeNote("", 70)
        lblSaveMsg.Font = UITheme.FontBold(10.5F)
        flow.Controls.Add(lblSaveMsg)
    End Sub

    Private Sub SaveClick(sender As Object, e As EventArgs)
        Dim err As String = AdminToolsService.SaveBusinessSettings(txtName.Text, txtPhone.Text, txtAddress.Text)
        If err <> "" Then
            ShowMsg(lblSaveMsg, err, UITheme.DangerRed)
        Else
            ShowMsg(lblSaveMsg, "Saved. Receipts use these details now. The sidebar name updates at the next login.",
                    UITheme.SuccessGreen)
        End If
    End Sub

    ' ---------- Backup ----------

    Private Sub BuildBackupCard(root As TableLayoutPanel)
        Dim flow As FlowLayoutPanel = AddCard(root, 1, "Database Backup")

        lblInfo = MakeNote("", 110)
        lblInfo.Font = UITheme.FontRegular(10.5F)
        lblInfo.ForeColor = UITheme.TextDark
        flow.Controls.Add(lblInfo)

        Dim btnNow As Button = UITheme.MakeButton("Create Backup Now", UITheme.Navy, 300, 50)
        AddHandler btnNow.Click, AddressOf BackupNowClick
        Dim btnAs As Button = UITheme.MakeButton("Save Backup As...", UITheme.RoyalBlue, 300, 50)
        AddHandler btnAs.Click, AddressOf BackupAsClick
        Dim btnFolder As Button = UITheme.MakeButton("Open Data Folder", UITheme.TextMuted, 300, 50)
        AddHandler btnFolder.Click,
            Sub(s, ev)
                Try
                    Process.Start("explorer.exe", Path.GetDirectoryName(DatabaseManager.DbPath))
                Catch
                End Try
            End Sub
        flow.Controls.Add(btnNow)
        flow.Controls.Add(btnAs)
        flow.Controls.Add(btnFolder)

        lblBackupMsg = MakeNote("", 60)
        lblBackupMsg.Font = UITheme.FontBold(10.5F)
        lblBackupMsg.Margin = New Padding(3, 10, 3, 10)
        flow.Controls.Add(lblBackupMsg)

        flow.Controls.Add(MakeNote("To restore: close this program, copy a backup file over repair.db in the " &
                                   "data folder, then start the program again. Keep copies on a USB drive " &
                                   "or another computer.", 90))
    End Sub

    Private Sub RefreshInfo()
        Dim size As String = "unknown"
        Try
            Dim bytes As Long = New FileInfo(DatabaseManager.DbPath).Length
            size = If(bytes >= 1048576, (bytes / 1048576.0).ToString("0.0") & " MB",
                      Math.Max(1, CInt(bytes / 1024)).ToString() & " KB")
        Catch
        End Try
        lblInfo.Text = "Live database:" & vbCrLf & DatabaseManager.DbPath & vbCrLf & vbCrLf & "Size: " & size
    End Sub

    Private Sub DoBackup(path As String)
        Dim err As String = AdminToolsService.CreateBackup(path)
        If err <> "" Then
            ShowMsg(lblBackupMsg, err, UITheme.DangerRed)
        Else
            ShowMsg(lblBackupMsg, "Backup saved to: " & path, UITheme.SuccessGreen)
        End If
        RefreshInfo()
    End Sub

    Private Sub BackupNowClick(sender As Object, e As EventArgs)
        Dim folder As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "PhoneRepairBackups")
        DoBackup(Path.Combine(folder, "repair_backup_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".db"))
    End Sub

    Private Sub BackupAsClick(sender As Object, e As EventArgs)
        Using dlg As New SaveFileDialog()
            dlg.Title = "Save database backup"
            dlg.Filter = "SQLite database (*.db)|*.db"
            dlg.FileName = "repair_backup_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".db"
            If dlg.ShowDialog() = DialogResult.OK Then DoBackup(dlg.FileName)
        End Using
    End Sub

End Class