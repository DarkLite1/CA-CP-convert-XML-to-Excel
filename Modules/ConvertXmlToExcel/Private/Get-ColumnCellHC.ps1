#Requires -Version 7

function Get-ColumnCellHC {
    <#
        .SYNOPSIS
            Return the cell addresses of one row: A1, B1, C1, ... AA1, AB1, ...

        .DESCRIPTION
            Used to write the table header row without hard coding column
            letters, so a worksheet with more than 26 columns keeps working.

            The other direction, a column letter to a column number, is not
            here: it is done where it is needed, inside Add-RowToWorkbookHC,
            against a cache of its own. That conversion runs once per cell of
            every row written, which is the hottest path of the whole script,
            so it is worth the few lines it takes to keep the call out of it.
            A copy used to live here as well and was called by nothing.

        .EXAMPLE
            Get-ColumnCellHC -RowNumber 2 -RequiredColumns 3   # A2, B2, C2
    #>
    param (
        [Parameter(Mandatory)]
        [int]$RowNumber,
        [int]$RequiredColumns = 50
    )

    <#
        Counted in base 26, from the last letter to the first, because a column
        letter is not quite a number in base 26: there is no zero digit, so 'Z'
        is followed by 'AA' and not by 'BA'. Subtracting one from what is left
        after each letter is what accounts for that.

        Worked out this way rather than with a letter for the first position and
        a letter for the second, which had nothing to put in the first position
        beyond column 'ZZ' and silently produced an address of two letters again
        from there on. No worksheet here comes near 702 columns, but a wrong
        address is written without complaint, so it is worth not being able to
        happen.
    #>
    for ($i = 0; $i -lt $RequiredColumns; $i++) {
        $columnLetter = ''
        $remaining = $i

        do {
            <#
                Both casts to [int] are needed. Dividing two whole numbers
                gives a fractional number in PowerShell, and [math]::Floor
                hands one back as well, so without them what is left over is
                no longer a whole number and [char] refuses it outright:
                'Cannot convert value "65" to type "System.Char"'.
            #>
            $columnLetter = [char][int](65 + ($remaining % 26)) + $columnLetter

            $remaining = [int][math]::Floor($remaining / 26) - 1
        } while ($remaining -ge 0)

        '{0}{1}' -f $columnLetter, $RowNumber
    }
}