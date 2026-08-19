#Requires -Version 7

function ConvertTo-CorrectTypeHC {
    <#
        .SYNOPSIS
            Convert an XML string value to a boolean, a double or a string.

        .DESCRIPTION
            XML values are always strings. This turns 'true'/'false' into a
            boolean and numeric text into a double, so numbers sort and total
            correctly in Excel. Anything else is returned unchanged. Empty and
            null values return nothing.

        .EXAMPLE
            ConvertTo-CorrectTypeHC -Value 'true'    # returns $true
            ConvertTo-CorrectTypeHC -Value '12.5'    # returns 12.5 (double)
            ConvertTo-CorrectTypeHC -Value 'ABC'     # returns 'ABC'
    #>
    Param(
        $Value
    )

    switch ($Value) {
        '' { break }
        $null { break }
        'true' { $true; break }
        'false' { $false; break }
        Default {
            try {
                [double]$_
            }
            catch {
                if ($Error.Count) { $Error.RemoveAt(0) }
                $Value
            }
        }
    }
}
