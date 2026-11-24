''' <summary>Temporary screen for features built in later stages.</summary>
Public Class PlaceholderPage
    Inherits UserControl

    Private card As Panel

    Public Sub New(item As NavItem)
        item.EnforceAccess()

        Me.BackColor = UITheme.LightGrey
        card = UITheme.MakeCard(600, 230)

        Dim lblTitle As Label = UITheme.MakeCenteredLabel(item.Text, 22, True,
                                                          UITheme.Navy, card.Width, 50)
        lblTitle.Top = 40

        Dim lblStage As Label = UITheme.MakeCenteredLabel(
            "This screen is built in " & item.StageNote & ".",
            12, False, UITheme.RoyalBlue, card.Width, 32)
        lblStage.Top = 100

        Dim access As String = If(item.AdminOnly, "Administrator only",
                               If(item.AdminAllowed AndAlso item.EmployeeAllowed,
                                  "Administrator and Employee", "Employee only"))
        Dim lblAccess As Label = UITheme.MakeCenteredLabel("Access: " & access,
                                                           10, False, UITheme.TextMuted,
                                                           card.Width, 28)
        lblAccess.Top = 145

        card.Controls.AddRange(New Control() {lblTitle, lblStage, lblAccess})
        Me.Controls.Add(card)

        AddHandler Me.Resize, Sub(s, ev) UITheme.CenterIn(card, Me)
    End Sub

End Class