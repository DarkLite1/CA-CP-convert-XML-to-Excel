#Requires -Version 7

function ConvertTo-DateTimeHC {
    <#
        .SYNOPSIS
            Convert an XML string value to a date.

        .DESCRIPTION
            Returns $null when the value is empty or not a valid date, so a
            missing or malformed date never stops the script and can be handled
            by the caller.

        .EXAMPLE
            ConvertTo-DateTimeHC -Value '2024-08-16T09:30:25+02:00'
            ConvertTo-DateTimeHC -Value ''      # returns $null
    #>
    param (
        $Value
    )

    if (-not $Value) { return $null }

    try {
        [DateTime]$Value
    }
    catch {
        if ($Error.Count) { $Error.RemoveAt(0) }
        $null
    }
}
