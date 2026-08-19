#Requires -Version 7

function Test-RowLimitHC {
    <#
        .SYNOPSIS
            Check up front whether all rows of one XML file still fit.

        .DESCRIPTION
            An XML file is added completely or not at all. Checking before
            writing avoids ending up with half a file in the Excel file when
            the maximum number of rows of a worksheet is reached.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [HashTable]$Workbook,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [Array]$Rows,
        [int]$MaxRowNumber = 1048576
    )

    $requiredRows = @{}

    foreach ($row in $Rows) {
        $requiredRows[$row.Key] = 1 + $requiredRows[$row.Key]
    }

    foreach ($item in $requiredRows.GetEnumerator()) {
        $sheet = $Workbook.Sheet[$item.Key]

        if (($sheet.RowNumber + $item.Value - 1) -ge $MaxRowNumber) {
            return $false
        }
    }

    $true
}
