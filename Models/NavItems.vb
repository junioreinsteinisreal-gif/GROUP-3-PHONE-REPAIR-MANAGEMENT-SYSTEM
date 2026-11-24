''' <summary>One sidebar menu entry and who may open it.</summary>
Public Class NavItem

    Public ReadOnly Property Key As String
    Public ReadOnly Property Text As String
    Public ReadOnly Property AdminAllowed As Boolean
    Public ReadOnly Property EmployeeAllowed As Boolean
    Public ReadOnly Property StageNote As String

    Public Sub New(key As String, text As String, adminAllowed As Boolean,
                   employeeAllowed As Boolean, stageNote As String)
        Me.Key = key
        Me.Text = text
        Me.AdminAllowed = adminAllowed
        Me.EmployeeAllowed = employeeAllowed
        Me.StageNote = stageNote
    End Sub

    Public ReadOnly Property AdminOnly As Boolean
        Get
            Return AdminAllowed AndAlso Not EmployeeAllowed
        End Get
    End Property

    Public Function IsAllowedFor(role As String) As Boolean
        If role = "Administrator" Then Return AdminAllowed
        If role = "Employee" Then Return EmployeeAllowed
        Return False
    End Function

    ''' <summary>Pages call this first. Throws if the current user may not open it.</summary>
    Public Sub EnforceAccess()
        SessionManager.RequireLogin()
        If Not IsAllowedFor(SessionManager.CurrentRole) Then
            Throw New UnauthorizedAccessException(
                "You do not have permission to open " & Text & ".")
        End If
    End Sub

End Class