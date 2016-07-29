##############################################################################
#.SYNOPSIS
# Gets an Excel file and for each sheet creates a JSON file with objects from
# the sheet.
#
#.DESCRIPTION
# This function assumes that there is a row in the sheet (the default is row 1)
# that contains the header of the table and gets the column names from there.
# Be aware of various data types and the formatting form the Excel file!
#
#
#.PARAMETER $ExcelFilePath
# The path to the Excel file
#
#.PARAMETER $HeaderRow
# The row that contains the header of the file (int). Default is 1
#
#.PARAMETER $SuppressFileCreation
# Switch variable that when passed, will suppress the creation of JSON files
# for each workbook in the sheet
#
#.EXAMPLE
# ExportExcelTo-Json -ExcelFilePath "pathToExcelFile.xlsx"
#
#.EXAMPLE
# ExportExcelTo-Json -ExcelFilePath $path -SuppressFileCreation  - this script takes the file from $path
# and only displays the result to the console
##############################################################################

function ExportExcelTo-Json
{
    Param(
         [string]$ExcelFilePath, 
         [int]$HeaderRow = 1, 
         [switch]$SuppressFileCreation)

    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($ExcelFilePath)

    for($i = 1; $i -le $workbook.Sheets.Count; $i++)
    {
        $sheet = $workbook.Sheets.Item($i)


        $sheetArray = @()

        for($row = $HeaderRow + 1; $row -le $sheet.UsedRange.Rows.Count; $row++)
        {
            $rowObject = New-Object -TypeName PSObject

            for($column = 1; $column -le $sheet.UsedRange.Columns.Count; $column++)
            {
                $rowObject | Add-Member -MemberType NoteProperty `
                                          -Name $sheet.Cells($headerRow, $column).Text `
                                          -Value $sheet.Cells($row, $column).Value2 `
            }

            $sheetArray += $rowObject
        }

        if(!$SuppressFileCreation)
        {
            ConvertTo-Json $sheetArray | Out-File "$($sheet.Name).json"
        }

        else 
        {
            ConvertTo-Json $sheetArray
        }
    }

    $workbook.Close()
    $excel.Quit()
    $excel = $null
    [GC]::Collect()
    
}
