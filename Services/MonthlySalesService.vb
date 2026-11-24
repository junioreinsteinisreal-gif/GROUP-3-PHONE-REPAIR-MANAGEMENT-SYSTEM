Imports System.Globalization
Imports System.IO
Imports ClosedXML.Excel
Imports Microsoft.Data.Sqlite

Public Class MonthFigures
    Public Property Collected As Double
    Public Property Refunded As Double
    Public Property PaymentCount As Integer
    Public Property Registered As Integer
    Public Property DaysCounted As Integer

    Public ReadOnly Property NetRevenue As Double
        Get
            Return Collected - Refunded
        End Get
    End Property

    Public ReadOnly Property AvgPerDay As Double
        Get
            If DaysCounted <= 0 Then Return 0
            Return NetRevenue / DaysCounted
        End Get
    End Property

    Public ReadOnly Property AvgPerPayment As Double
        Get
            If PaymentCount <= 0 Then Return 0
            Return Collected / PaymentCount
        End Get
    End Property
End Class

''' <summary>Monthly sales figures (administrator only). Payments minus refunds is the revenue.</summary>
Public Module MonthlySalesService

    Private Function Num(sql As String, ParamArray params() As SqliteParameter) As Double
        Dim o As Object = DatabaseManager.ExecuteScalar(sql, params)
        If o Is Nothing OrElse IsDBNull(o) Then Return 0
        Return Convert.ToDouble(o)
    End Function

    Private Function YearMonth(yr As Integer, mo As Integer) As String
        Return yr.ToString("D4", CultureInfo.InvariantCulture) & "-" &
               mo.ToString("D2", CultureInfo.InvariantCulture)
    End Function

    Public Function GetMonthFigures(yr As Integer, mo As Integer) As MonthFigures
        SessionManager.RequireAdmin()
        Dim ym As String = YearMonth(yr, mo)
        Dim f As New MonthFigures()

        f.Collected = Num("SELECT CAST(COALESCE(SUM(Amount), 0) AS REAL) FROM Payments " &
                          "WHERE substr(PaidAt, 1, 7) = @ym", DatabaseManager.P("@ym", ym))
        f.PaymentCount = CInt(Num("SELECT COUNT(*) FROM Payments WHERE substr(PaidAt, 1, 7) = @ym",
                                  DatabaseManager.P("@ym", ym)))
        f.Refunded = Num("SELECT CAST(COALESCE(SUM(Amount), 0) AS REAL) FROM Refunds " &
                         "WHERE substr(RefundedAt, 1, 7) = @ym", DatabaseManager.P("@ym", ym))
        f.Registered = CInt(Num("SELECT COUNT(*) FROM Repairs WHERE substr(CreatedAt, 1, 7) = @ym",
                                DatabaseManager.P("@ym", ym)))

        ' The current month is averaged over the days so far.
        If yr = Date.Today.Year AndAlso mo = Date.Today.Month Then
            f.DaysCounted = Date.Today.Day
        Else
            f.DaysCounted = Date.DaysInMonth(yr, mo)
        End If
        Return f
    End Function

    ''' <summary>Weeks are days 1-7, 8-14, 15-21, 22-28 and 29-end.</summary>
    Public Function GetWeekly(yr As Integer, mo As Integer) As DataTable
        SessionManager.RequireAdmin()
        Dim ym As String = YearMonth(yr, mo)
        Dim totalDays As Integer = Date.DaysInMonth(yr, mo)

        Dim pay As DataTable = DatabaseManager.GetTable(
            "SELECT CAST(substr(PaidAt, 9, 2) AS INTEGER) AS D, COUNT(*) AS Cnt, " &
            "CAST(SUM(Amount) AS REAL) AS Total FROM Payments WHERE substr(PaidAt, 1, 7) = @ym " &
            "GROUP BY substr(PaidAt, 9, 2)", DatabaseManager.P("@ym", ym))
        Dim ref As DataTable = DatabaseManager.GetTable(
            "SELECT CAST(substr(RefundedAt, 9, 2) AS INTEGER) AS D, " &
            "CAST(SUM(Amount) AS REAL) AS Total FROM Refunds WHERE substr(RefundedAt, 1, 7) = @ym " &
            "GROUP BY substr(RefundedAt, 9, 2)", DatabaseManager.P("@ym", ym))

        Dim cnt(4) As Integer
        Dim collected(4) As Double
        Dim refunded(4) As Double
        For Each r As DataRow In pay.Rows
            Dim wk As Integer = Math.Min(4, (Convert.ToInt32(r("D")) - 1) \ 7)
            cnt(wk) += Convert.ToInt32(r("Cnt"))
            collected(wk) += Convert.ToDouble(r("Total"))
        Next
        For Each r As DataRow In ref.Rows
            Dim wk As Integer = Math.Min(4, (Convert.ToInt32(r("D")) - 1) \ 7)
            refunded(wk) += Convert.ToDouble(r("Total"))
        Next

        Dim dt As New DataTable("Weekly")
        dt.Columns.Add("Week", GetType(String))
        dt.Columns.Add("Payments", GetType(Integer))
        dt.Columns.Add("Collected", GetType(Double))
        dt.Columns.Add("Refunded", GetType(Double))
        dt.Columns.Add("Net", GetType(Double))

        Dim mon As String = New Date(yr, mo, 1).ToString("MMM", CultureInfo.InvariantCulture)
        For wk As Integer = 0 To 4
            Dim startDay As Integer = wk * 7 + 1
            If startDay > totalDays Then Exit For
            Dim endDay As Integer = Math.Min(startDay + 6, totalDays)
            Dim rangeText As String = If(startDay = endDay, startDay & " " & mon,
                                         startDay & "-" & endDay & " " & mon)
            dt.Rows.Add("Week " & (wk + 1) & "  (" & rangeText & ")", cnt(wk), collected(wk),
                        refunded(wk), collected(wk) - refunded(wk))
        Next
        Return dt
    End Function

    Public Function GetEmployeeMonth(yr As Integer, mo As Integer) As DataTable
        SessionManager.RequireAdmin()
        Return DatabaseManager.GetTable(
            "SELECT t.Employee, t.Registered, t.Payments, t.Collected, t.Refunded, " &
            "CAST(t.Collected - t.Refunded AS REAL) AS Net, " &
            "CASE WHEN t.Payments > 0 THEN CAST(t.Collected / t.Payments AS REAL) ELSE 0.0 END AS AvgPayment " &
            "FROM (SELECT e.FullName AS Employee, " &
            "(SELECT COUNT(*) FROM Repairs r WHERE r.EmployeeId = e.EmployeeId " &
            " AND substr(r.CreatedAt, 1, 7) = @ym) AS Registered, " &
            "(SELECT COUNT(*) FROM Payments p WHERE p.EmployeeId = e.EmployeeId " &
            " AND substr(p.PaidAt, 1, 7) = @ym) AS Payments, " &
            "CAST((SELECT COALESCE(SUM(p.Amount), 0) FROM Payments p WHERE p.EmployeeId = e.EmployeeId " &
            " AND substr(p.PaidAt, 1, 7) = @ym) AS REAL) AS Collected, " &
            "CAST((SELECT COALESCE(SUM(f.Amount), 0) FROM Refunds f WHERE f.EmployeeId = e.EmployeeId " &
            " AND substr(f.RefundedAt, 1, 7) = @ym) AS REAL) AS Refunded " &
            "FROM Employees e) t WHERE t.Payments > 0 OR t.Refunded > 0 OR t.Registered > 0 " &
            "ORDER BY Net DESC, t.Employee",
            DatabaseManager.P("@ym", YearMonth(yr, mo)))
    End Function

    ''' <summary>Creates the monthly workbook. Returns its file path.</summary>
    Public Function BuildMonthReport(yr As Integer, mo As Integer) As String
        SessionManager.RequireAdmin()
        Dim f As MonthFigures = GetMonthFigures(yr, mo)
        Dim ym As String = YearMonth(yr, mo)

        Dim summary As New DataTable("Summary")
        summary.Columns.Add("Item", GetType(String))
        summary.Columns.Add("Value", GetType(String))
        summary.Rows.Add("Business", AdminToolsService.GetSetting("BusinessName"))
        summary.Rows.Add("Month", New Date(yr, mo, 1).ToString("MMMM yyyy", CultureInfo.InvariantCulture))
        summary.Rows.Add("Report created", DatabaseManager.NowText() & " by " & SessionManager.CurrentUserName)
        summary.Rows.Add("", "")
        summary.Rows.Add("Net revenue (after refunds)", UIHelpers.Money(f.NetRevenue))
        summary.Rows.Add("Payments collected", UIHelpers.Money(f.Collected))
        summary.Rows.Add("Number of payments", f.PaymentCount.ToString())
        summary.Rows.Add("Refunds given", UIHelpers.Money(f.Refunded))
        summary.Rows.Add("Average revenue per day (" & f.DaysCounted & " days)", UIHelpers.Money(f.AvgPerDay))
        summary.Rows.Add("Average payment", UIHelpers.Money(f.AvgPerPayment))
        summary.Rows.Add("Repairs registered", f.Registered.ToString())

        Dim filePath As String = Path.Combine(EndOfDayService.ReportsFolder(),
                                              "MonthlySales_" & ym & ".xlsx")
        Using wb As New XLWorkbook()
            EndOfDayService.AddSheet(wb, summary, "Summary")
            EndOfDayService.AddSheet(wb, GetWeekly(yr, mo), "Weekly")
            EndOfDayService.AddSheet(wb, GetEmployeeMonth(yr, mo), "Employees")
            wb.SaveAs(filePath)
        End Using
        Return filePath
    End Function

End Module