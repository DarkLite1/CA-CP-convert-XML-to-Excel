#Requires -Version 7

function Add-RowToWorkbookHC {
    <#
        .SYNOPSIS
            Write the rows of one XML file to the Excel file.

        .DESCRIPTION
            Returns the number of rows written. Throws when the rows do not fit
            anymore, without having written anything.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [HashTable]$Workbook,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [Array]$Rows
    )

    if (-not $Rows.Count) { return 0 }

    if (-not (Test-RowLimitHC -Workbook $Workbook -Rows $Rows)) {
        throw 'XML file could not be added to Excel, max row limit reached'
    }

    foreach ($row in $Rows) {
        $sheet = $Workbook.Sheet[$row.Key]

        $row.Cells.GetEnumerator().ForEach(
            {
                $address = '{0}{1}' -f $_.Key, $sheet.RowNumber

                $sheet.Worksheet.Cells[$address].Value = $_.Value
            }
        )

        $sheet.RowNumber++
    }

    $Rows.Count
}
