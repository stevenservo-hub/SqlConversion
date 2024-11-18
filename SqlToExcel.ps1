try {
    $server = "your_server"
    $database = "your_database"
    $query = "SELECT * FROM your_table"
    $excelFile = "C:\example.xlsx"
    
    $connectionString = "Server=$server;Database=$database;Integrated Security=True;"
    $dataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $dataAdapter.SelectCommand = New-Object System.Data.SqlClient.SqlCommand($query, $connectionString)
    $dataSet = New-Object System.Data.DataSet
    $dataAdapter.Fill($dataSet)
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    
    $row = 1
    foreach ($column in $dataSet.Tables[0].Columns) {
        $worksheet.Cells.Item($row, $column.Ordinal + 1) = $column.ColumnName
    }
    
    $row = 2
    $dataRows = $dataSet.Tables[0].Rows
    $columns = $dataSet.Tables[0].Columns

    $dataRows | ForEach-Object -Parallel {
        param ($dataRow, $columns, $row)
        $result = @()
        foreach ($column in $columns) {
            $result += [PSCustomObject]@{ Row = $row; Column = $column.Ordinal + 1; Value = $dataRow[$column] }
        }
        return $result
    } -ArgumentList $_, $columns, $row++ | ForEach-Object {
        $worksheet.Cells.Item($_.Row, $_.Column) = $_.Value
    }

    $workbook.SaveAs($excelFile)
    $excel.Quit()
}
catch {
    write-error "error: $_"
}




























