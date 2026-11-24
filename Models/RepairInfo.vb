''' <summary>One repair with its payment totals, used by the update screen.</summary>
Public Class RepairInfo

    Public Property RepairCode As String = ""
    Public Property CustomerName As String = ""
    Public Property CustomerPhone As String = ""
    Public Property Brand As String = ""
    Public Property Model As String = ""
    Public Property Problem As String = ""
    Public Property Status As String = ""
    Public Property EstimatedCost As Decimal
    Public Property TotalPaid As Decimal
    Public Property DateReceived As Date
    Public Property ExpectedCollection As Date
    Public Property CollectedAt As String = ""
    Public Property EmployeeName As String = ""
    Public Property OwnerId As Integer

    ''' <summary>Nothing is owed on a terminated (Unrepairable) repair.</summary>
    Public ReadOnly Property Balance As Decimal
        Get
            If Status = "Unrepairable" Then Return 0D
            Return EstimatedCost - TotalPaid
        End Get
    End Property

End Class