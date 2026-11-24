Imports System.Drawing.Printing

''' <summary>Draws a receipt and sends it to preview or a printer.</summary>
Public Class ReceiptPrinter

    Private ReadOnly data As ReceiptData
    Private ReadOnly doc As PrintDocument

    Public Sub New(receipt As ReceiptData, thermal As Boolean)
        data = receipt
        doc = New PrintDocument()
        doc.DocumentName = "Receipt " & data.ReceiptNumber
        If thermal Then
            doc.DefaultPageSettings.PaperSize = New PaperSize("Receipt 80mm", 315, 650)
            doc.DefaultPageSettings.Margins = New Margins(10, 10, 10, 10)
        End If
        AddHandler doc.PrintPage, AddressOf OnPrintPage
        AddHandler doc.EndPrint, AddressOf OnEndPrint
    End Sub

    Public Sub Preview(owner As IWin32Window)
        Using dlg As New PrintPreviewDialog()
            dlg.Document = doc
            dlg.Width = 760
            dlg.Height = 860
            dlg.ShowDialog(owner)
        End Using
    End Sub

    Public Sub PrintNow(owner As IWin32Window)
        Using pd As New PrintDialog()
            pd.Document = doc
            pd.UseEXDialog = True
            If pd.ShowDialog(owner) = DialogResult.OK Then doc.Print()
        End Using
    End Sub

    ''' <summary>Only a real print counts, not a preview.</summary>
    Private Sub OnEndPrint(sender As Object, e As PrintEventArgs)
        Try
            Dim pd As PrintDocument = DirectCast(sender, PrintDocument)
            If pd.PrintController IsNot Nothing AndAlso Not pd.PrintController.IsPreview Then
                ReceiptService.RecordPrint(data)
            End If
        Catch
        End Try
    End Sub

    Private Sub OnPrintPage(sender As Object, e As PrintPageEventArgs)
        Dim g As Graphics = e.Graphics
        Dim left As Single = e.PageSettings.Margins.Left
        Dim w As Single = 300
        Dim y As Single = e.PageSettings.Margins.Top

        Dim fTitle As New Font("Segoe UI", 13, FontStyle.Bold)
        Dim fBold As New Font("Segoe UI", 9, FontStyle.Bold)
        Dim fBody As New Font("Segoe UI", 9, FontStyle.Regular)
        Dim fSmall As New Font("Segoe UI", 8, FontStyle.Regular)
        Dim centered As New StringFormat() With {.Alignment = StringAlignment.Center}
        Dim rightAlign As New StringFormat() With {.Alignment = StringAlignment.Far}
        Dim dash As New Pen(Color.Black) With {.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash}

        Dim centerLine = Sub(text As String, f As Font)
                             Dim h As Single = g.MeasureString(text, f, CInt(w)).Height
                             g.DrawString(text, f, Brushes.Black, New RectangleF(left, y, w, h), centered)
                             y += h
                         End Sub

        Dim row = Sub(label As String, value As String, f As Font)
                      Dim lw As Single = 112
                      Dim h1 As Single = g.MeasureString(label, f, CInt(lw)).Height
                      Dim h2 As Single = g.MeasureString(value, f, CInt(w - lw)).Height
                      Dim h As Single = Math.Max(h1, h2)
                      g.DrawString(label, f, Brushes.Black, New RectangleF(left, y, lw, h))
                      g.DrawString(value, f, Brushes.Black, New RectangleF(left + lw, y, w - lw, h), rightAlign)
                      y += h + 2
                  End Sub

        Dim divider = Sub()
                          g.DrawLine(dash, left, y + 4, left + w, y + 4)
                          y += 10
                      End Sub

        ' ----- Header -----
        centerLine(If(data.BusinessName.Trim() = "", "Phone Repair Shop", data.BusinessName), fTitle)
        If data.BusinessAddress.Trim() <> "" Then centerLine(data.BusinessAddress, fSmall)
        If data.BusinessPhone.Trim() <> "" Then centerLine("Tel: " & data.BusinessPhone, fSmall)
        divider()
        centerLine("PAYMENT RECEIPT", fBold)
        If data.PrintCount > 0 Then centerLine("*** DUPLICATE COPY ***", fBold)
        y += 4
        row("Receipt No:", data.ReceiptNumber, fBold)
        row("Date / Time:", data.PaidAt, fBody)
        divider()

        ' ----- Customer and device -----
        row("Customer:", data.CustomerName, fBody)
        row("Phone:", data.CustomerPhone, fBody)
        row("Repair code:", data.RepairCode, fBold)
        row("Device:", data.Brand & " " & data.Model, fBody)
        divider()

        ' ----- Money -----
        row("Estimated cost:", UIHelpers.Money(CDbl(data.EstimatedCost)), fBody)
        row("Previously paid:", UIHelpers.Money(CDbl(data.PreviousPaid)), fBody)
        row("This payment:", UIHelpers.Money(CDbl(data.Amount)), fBold)
        row("Total paid:", UIHelpers.Money(CDbl(data.TotalPaid)), fBody)
        row("Balance due:", UIHelpers.Money(CDbl(data.Balance)), fBold)
        row("Payment method:", data.Method, fBody)
        divider()

        row("Served by:", data.EmployeeName, fBody)
        y += 6
        centerLine("Thank you for your business!", fBody)
        centerLine("Keep this receipt and bring your repair code on collection.", fSmall)

        e.HasMorePages = False
    End Sub

End Class