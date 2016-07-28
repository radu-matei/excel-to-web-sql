##############################################################################
#.SYNOPSIS
# Creates a CREATE TABLE script with the names and (bsaic) data types of each 
# column of an Excel file
#
#.DESCRIPTION
# This function takes an Excel file and based on the header of the table and
# the number format of the cells creates a basic CREATE TABLE script for each
# sheet in the file.
#
#.PARAMETER $excelFilePath
# The path to the Excel file
#
#.PARAMETER $headerRow
# The row that contains the header of the file (int). Default is 1
#
#.PARAMETER $exportToFile
# If true (default), the function will create a separate .sql file 
# containing the script to create the table. Else, it will simply
# display the script to the console.
#
#.EXAMPLE
# CreateSqlTableFrom-Excel -excelFilePath "pathToExcelFile.xlsx"
#
#.EXAMPLE
# CreateSqlTableFrom-Excel -excelFilePath $path -exportToFile $false  - this script takes the file from $path
# and only displays the result to the console
##############################################################################


function CreateSqlTableFrom-Excel
{
    Param([string]$excelFilePath, [int]$headerRow = 1, [bool]$exportToFile = $true, [int]$sqlStringLength = 50)

    $formatDictionary = @{"@" = "nvarchar(" + $sqlStringLength + ")"; 
                          "0" = "int"; 
                          "0,000" = "float"; 
                          "0,00" = "float"; 
                          "0,0" = "float"; 
                          "General" = "nvarchar(" + $sqlStringLength + ")";
                         }

    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($path)
    
    for($i = 1; $i -le $workbook.Sheets.Count; $i++)
    {
        $sqlCreateTableStatement = ""
        $sheet = $workbook.Sheets.Item($i)
        
        $sqlCreateTableStatement += "CREATE TABLE " + $sheet.Name + "`n("

        for ($column = 1; $column -le $sheet.UsedRange.Columns.Count; $column++)
        {
            $sqlCreateTableStatement += "`n" + 
                                        $sheet.Cells($headerRow, $column).Text + 
                                        " " + 
                                        $formatDictionary[$sheet.Cells($headerRow, $column).NumberFormat]
        }
        
        $sqlCreateTableStatement += "`n)"

        if($exportToFile)
        {
            $scriptFileName = "CREATE " + $sheet.Name + ".sql"
            $sqlCreateTableStatement | Out-File $scriptFileName
        }

        else 
        {
            $sqlCreateTableStatement
        }
    }
    
    $workbook.Close()
    $excel.Quit()
    $excel = $null
    [GC]::Collect()
}