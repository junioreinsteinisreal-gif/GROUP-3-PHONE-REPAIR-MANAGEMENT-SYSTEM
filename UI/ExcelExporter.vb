Imports ClosedXML.Excel

''' <summary>Exports what a grid is showing (visible columns, current sort) to .xlsx.</summary>
Public Module ExcelExporter

    Public Function ExportGrid(grid As DataGridView, baseName As String,
                               sheetName As String) As Boolean
        Dim src As DataTable = TryCast(grid.DataSource, DataTable)
        If src Is Nothing OrElse src.Rows.Count = 0 Then
            MessageBox.Show("There is nothing to export. Run a search first.",
                            "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If

        Dim savePath As String = ""
        Using dlg As New SaveFileDialog()
            dlg.Title = "Export to Excel"
            dlg.Filter = "Excel workbook (*.xlsx)|*.xlsx"
            dlg.FileName = baseName & "_" & DateTime.Now.ToString("yyyyMMdd_HHmm") & ".xlsx"
            If dlg.ShowDialog() <> DialogResult.OK Then Return False
            savePath = dlg.FileName
        End Using

        Try
            Dim outTable As New DataTable("ExportData")
            Dim fields As New List(Of String)

            For Each col As DataGridViewColumn In grid.Columns
                If Not col.Visible OrElse col.DataPropertyName = "" OrElse
                   Not src.Columns.Contains(col.DataPropertyName) Then Continue For

                Dim header As String = col.HeaderText.Replace("&&", "&").Replace(vbCr, " ").
                                       Replace(vbLf, " ").Trim()
                If header = "" Then header = col.DataPropertyName
                Dim finalName As String = header
                Dim n As Integer = 2
                While outTable.Columns.Contains(finalName)
                    finalName = header & " " & n
                    n += 1
                End While
                outTable.Columns.Add(finalName, src.Columns(col.DataPropertyName).DataType)
                fields.Add(col.DataPropertyName)
            Next

            For Each drv As DataRowView In src.DefaultView
                Dim vals(fields.Count - 1) As Object
                For i As Integer = 0 To fields.Count - 1
                    vals(i) = drv(fields(i))
                Next
                outTable.Rows.Add(vals)
            Next

            Using wb As New XLWorkbook()
                Dim ws = wb.Worksheets.Add(outTable, sheetName)
                ws.Columns().AdjustToContents()
                wb.SaveAs(savePath)
            End Using

            MessageBox.Show("Exported " & outTable.Rows.Count & " rows to:" & vbCrLf & savePath,
                            "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return True
        Catch ioEx As System.IO.IOException
            MessageBox.Show("The file could not be written. If it is open in Excel, close it and try again.",
                            "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch exc As Exception
            MessageBox.Show("The export failed: " & exc.Message,
                            "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return False
    End Function

End Module