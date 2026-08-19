#Requires -Version 7

function Get-ColumnCellHC {
    <#
        .SYNOPSIS
            Return the cell addresses of one row: A1, B1, C1, ... AA1, AB1, ...

        .DESCRIPTION
            Used to write the table header row without hard coding column
            letters, so a worksheet with more than 26 columns keeps working.

        .EXAMPLE
            Get-ColumnCellHC -RowNumber 2 -RequiredColumns 3   # A2, B2, C2
    #>
    param (
        [Parameter(Mandatory)]
        [double]$RowNumber,
        [double]$RequiredColumns = 50
    )

    $alphaToZulu = [char[]](65..90)

    for ($i = 0; $i -lt $RequiredColumns; $i++) {

        $start = $null

        $result = [math]::Floor($i / 26)

        if ($result -ge 1) {
            $start = $alphaToZulu[$result - 1]

            $index = $i - (26 * $result)
        }
        else {
            $index = $i
        }

        '{0}{1}{2}' -f $start, $alphaToZulu[$index], $RowNumber
    }
}
