# Load the Stopwatch class
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$dllpath = "C:\Users\salvira1\Downloads\FilepathR\EPPlus.dll"
$excelfile = "C:\Users\salvira1\Downloads\FilepathR\TenMRecords.xlsx"
$refrenceAssembly = "C:\Users\salvira1\Downloads\FilepathR\System.Text.Encoding.CodePages.dll"
Import-Module -Name "PSSQLite"
$DataSource = "C:\Users\salvira1\Downloads\FilepathR\SalesData.SQLite"
$dataTable = New-Object System.Data.DataTable

# ADD DLLS with Deps
Add-Type -Path $dllpath -ReferencedAssemblies $refrenceAssembly 
$excelPackage = New-Object OfficeOpenXml.ExcelPackage -ArgumentList (New-Object IO.FileInfo $excelfile)

# Read Excel
$worksheet = $excelPackage.Workbook.Worksheets[1]

#Create SQL TAble and Datatable
$columns = $worksheet.Dimension.End.Column
$tableExe = ""; 
$createTableQuery = "CREATE TABLE IF NOT EXISTS SalesData ("
$dataType = " VARCHAR(2000),"

for ($c = 1; $c -le $columns; $c++) {
   $Coloumvalue = $worksheet.Cells[1, $c].Text
   $dataTable.Columns.Add($Coloumvalue, [System.Type]::GetType("System.String"))
   # Enclose column names with double quotes to handle spaces or special characters
   $tableExe = $tableExe + "`"$Coloumvalue`"" + $dataType
}
$tableExe  = $tableExe.TrimEnd(",")
# Construct the final create table query
$createTableQuery = $createTableQuery + $tableExe + " )"
Invoke-SqliteQuery -Query $createTableQuery -DataSource $DataSource


# Read data from Excel and add to DataTable
$rows = $worksheet.Dimension.End.Row
for ($r = 2; $r -le $rows; $r++) {
    $row = $dataTable.NewRow()
    for ($c = 1; $c -le $columns; $c++) {
        $row[$worksheet.Cells[1, $c].Text] = $worksheet.Cells[$r, $c].Text
    }
    $dataTable.Rows.Add($row)
}

Invoke-SQLiteBulkCopy -DataSource $DataSource -DataTable $dataTable -Table "SalesData" -Force

# Stop the stopwatch
$stopwatch.Stop()

# Output the elapsed time
Write-Output "Execution Time: $($stopwatch.Elapsed.TotalSeconds) seconds"

Invoke-SqliteQuery -DataSource $DataSource -Query "SELECT Count(*) FROM SalesData"



