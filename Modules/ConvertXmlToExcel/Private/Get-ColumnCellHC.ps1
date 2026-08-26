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

<#
    Column letter to column number, remembered after the first conversion.

    A worksheet has at most a few dozen columns but a run writes millions of
    cells, so every lookup after the first few is a hash hit.
#>
$script:ColumnNumberCache = @{}

function Get-ColumnNumberHC {
    <#
        .SYNOPSIS
            Convert a column letter to its column number: A is 1, Z is 26,
            AA is 27.

        .DESCRIPTION
            EPPlus can resolve a column letter itself, but only as part of
            parsing a full cell address like 'AB1234'. Writing a cell by row
            and column number avoids that parsing, and this is what turns the
            letters used in the worksheet definitions into numbers.

        .EXAMPLE
            Get-ColumnNumberHC -ColumnLetter 'AB'   # 28
    #>
    param (
        [Parameter(Mandatory)]
        [String]$ColumnLetter
    )

    $columnNumber = $script:ColumnNumberCache[$ColumnLetter]

    if ($columnNumber) { return $columnNumber }

    $columnNumber = 0

    foreach ($letter in $ColumnLetter.ToUpperInvariant().ToCharArray()) {
        if ($letter -lt 'A' -or $letter -gt 'Z') {
            throw "Column '$ColumnLetter' is not a valid column letter"
        }

        $columnNumber = ($columnNumber * 26) + ([int]$letter - 64)
    }

    $script:ColumnNumberCache[$ColumnLetter] = $columnNumber

    $columnNumber
}