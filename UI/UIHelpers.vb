Imports System.Globalization

''' <summary>Shared money formatting and table styling.</summary>
Public Module UIHelpers

    ' "GH" followed by the cedi sign
    Public ReadOnly CurrencySymbol As String = "GH" & ChrW(&H20B5)

    Public Function Money(amount As Double) As String
        Return CurrencySymbol & " " & amount.ToString("N2", CultureInfo.InvariantCulture)
    End Function

    ''' <summary>Gives a DataGridView the app's modern look.</summary>
    Public Sub StyleGrid(grid As DataGridView)
        grid.AutoGenerateColumns = False
        grid.ReadOnly = True
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.AllowUserToResizeRows = False
        grid.MultiSelect = False
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.RowHeadersVisible = False
        grid.BorderStyle = BorderStyle.None
        grid.BackgroundColor = UITheme.CardWhite
        grid.GridColor = Color.FromArgb(215, 220, 230)
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal

        grid.EnableHeadersVisualStyles = False
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        grid.ColumnHeadersHeight = 58
        grid.ColumnHeadersDefaultCellStyle.BackColor = UITheme.Navy
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = UITheme.Navy
        grid.ColumnHeadersDefaultCellStyle.Font = UITheme.FontBold(10.5F)
        grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        grid.ColumnHeadersDefaultCellStyle.Padding = New Padding(8, 0, 8, 0)

        grid.RowTemplate.Height = 44
        grid.DefaultCellStyle.Font = UITheme.FontRegular(10.5F)
        grid.DefaultCellStyle.BackColor = UITheme.CardWhite
        grid.DefaultCellStyle.ForeColor = UITheme.TextDark
        grid.DefaultCellStyle.SelectionBackColor = UITheme.RoyalBlue
        grid.DefaultCellStyle.SelectionForeColor = Color.White
        grid.DefaultCellStyle.Padding = New Padding(8, 0, 8, 0)
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(243, 246, 252)

        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    ''' <summary>Adds a bound text column. dataName is the DataTable column name.</summary>
    Public Sub AddTextColumn(grid As DataGridView, dataName As String, header As String,
                             weight As Single,
                             Optional rightAlign As Boolean = False,
                             Optional format As String = "")
        Dim col As New DataGridViewTextBoxColumn()
        col.Name = dataName
        col.DataPropertyName = dataName
        col.HeaderText = header
        col.FillWeight = weight
        col.MinimumWidth = 80
        If rightAlign Then
            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        End If
        If format <> "" Then col.DefaultCellStyle.Format = format
        grid.Columns.Add(col)
    End Sub

End Module