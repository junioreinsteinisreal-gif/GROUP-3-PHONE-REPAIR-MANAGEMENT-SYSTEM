''' <summary>Holds the currently logged-in user and enforces permissions.</summary>
Public Module SessionManager

    Private _role As String = ""
    Private _userId As Integer = 0
    Private _userName As String = ""
    Private _sessionId As Long = 0
    Private _loginTime As DateTime = DateTime.MinValue

    Public ReadOnly Property CurrentRole As String
        Get
            Return _role
        End Get
    End Property

    Public ReadOnly Property CurrentUserId As Integer
        Get
            Return _userId
        End Get
    End Property

    Public ReadOnly Property CurrentUserName As String
        Get
            Return _userName
        End Get
    End Property

    Public ReadOnly Property CurrentSessionId As Long
        Get
            Return _sessionId
        End Get
    End Property

    Public ReadOnly Property LoginTime As DateTime
        Get
            Return _loginTime
        End Get
    End Property

    Public ReadOnly Property IsLoggedIn As Boolean
        Get
            Return _role <> ""
        End Get
    End Property

    Public ReadOnly Property IsAdmin As Boolean
        Get
            Return _role = "Administrator"
        End Get
    End Property

    Public ReadOnly Property IsEmployee As Boolean
        Get
            Return _role = "Employee"
        End Get
    End Property

    Public Sub Start(role As String, userId As Integer,
                     userName As String, sessionId As Long)
        _role = role
        _userId = userId
        _userName = userName
        _sessionId = sessionId
        _loginTime = DateTime.Now
    End Sub

    Public Sub Clear()
        _role = ""
        _userId = 0
        _userName = ""
        _sessionId = 0
        _loginTime = DateTime.MinValue
    End Sub

    ''' <summary>Call at the top of any admin-only form or function.</summary>
    Public Sub RequireAdmin()
        If Not IsAdmin Then
            Throw New UnauthorizedAccessException(
                "Administrator access is required for this action.")
        End If
    End Sub

    Public Sub RequireLogin()
        If Not IsLoggedIn Then
            Throw New UnauthorizedAccessException("You must be logged in.")
        End If
    End Sub

End Module