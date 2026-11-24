''' <summary>Maps menu keys to pages. Anything not built yet shows the placeholder.</summary>
Public Module NavConfig

    Public Function CreatePage(key As String, item As NavItem) As Control
        Select Case key
            Case "employees"
                Return New EmployeeManagementPage()
            Case "myactivity", "activityall"
                Return New ActivityHistoryPage()
            Case "backup"
                Return New SettingsPage()
            Case "audit"
                Return New AdminHistoryPage()
            Case "deleterepair"
                Return New DeleteRepairPage()
            Case "reminders"
                Return New RemindersPage()
            Case "endofday"
                Return New EndOfDayPage()
            Case "monthly"
                Return New MonthlySalesPage()
            Case "feedback"
                Return New FeedbackPage()
            Case "reviewsadmin"
                Return New ReviewsAdminPage()
            Case "notifications"
                Return New NotificationsPage()
            Case "reviewsite"
                Return New ReviewSitePage()
            Case Else
                Return New PlaceholderPage(item)
        End Select
    End Function

End Module