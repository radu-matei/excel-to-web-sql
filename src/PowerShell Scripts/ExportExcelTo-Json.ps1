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
#.PARAMETER $excelFilePath
# The path to the Excel file
#
#.PARAMETER $headerRow
# The row that contains the header of the file (int). Default is 1
#
#.PARAMETER $exportToFile
# If true (default), the function will export the data from each sheet into a 
# separate JSON file with the same name as the sheet
#
#.EXAMPLE
# ExportExcelTo-Json -excelFilePath "pathToExcelFile.xlsx"
#
#.EXAMPLE
# ExportExcelTo-Json -excelFilePath $path -exportToFile $false  - this script takes the file from $path
# and only displays the result to the console
##############################################################################

function ExportExcelTo-Json
{
    Param([string]$excelFilePath, [int]$headerRow = (1), [bool]$exportToFile = ($true))

    $excel = New-Object -ComObject Excel.Application

    for($i = 1; $i -le $workbook.Sheets.Count; $i++)
    {
        $sheet = $workbook.Sheets.Item($i)


        $sheetArray = @()

        for($row = $headerRow + 1; $row -le $sheet.UsedRange.Rows.Count; $row++)
        {
            $sheetObject = New-Object -TypeName PSObject

            for($column = 1; $column -le $sheet.UsedRange.Columns.Count; $column++)
            {
                $sheetObject | Add-Member -MemberType NoteProperty -Name $sheet.Cells($headerRow, $column).Text -Value $sheet.Cells($row, $column).Value2
            }

            $sheetArray += $sheetObject
        }

        if($exportToFile)
        {
            ConvertTo-Json $sheetArray | Out-File "$($sheet.Name).json"
        }

        else 
        {
            $sheetArray
        }
    }
}