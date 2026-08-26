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

            This function runs once per cell, so it is the hottest code path in
            the whole script. It uses '[double]::TryParse' instead of a
            '[double]' cast inside a try/catch, because the cast raises a .NET
            exception for every value that is not a number. Most cells hold
            text (material names, units, license plates, vendors), so the cast
            spent the bulk of its time throwing and catching exceptions.

            The parse settings match what the '[double]' cast did before:
            InvariantCulture, so '12.5' is read the same way on a machine with
            any regional setting, and Float + AllowThousands, which is the
            style PowerShell itself uses for a string to double conversion.

        .PARAMETER Value
            The text to convert, usually a leaf value coming from the XML file.

        .EXAMPLE
            ConvertTo-CorrectTypeHC -Value 'true'    # returns $true
            ConvertTo-CorrectTypeHC -Value '12.5'    # returns 12.5 (double)
            ConvertTo-CorrectTypeHC -Value 'ABC'     # returns 'ABC'
    #>
    Param(
        $Value
    )

    if ($null -eq $Value) { return }

    $stringValue = [String]$Value

    if ($stringValue.Length -eq 0) { return }

    <#
        A plain '-eq' comparison, which is case insensitive in PowerShell just
        like the 'switch' statement it replaces, so 'TRUE' and 'True' are still
        recognized.
    #>
    if ($stringValue -eq 'true') { return $true }
    if ($stringValue -eq 'false') { return $false }

    $doubleValue = [double]0

    $isDouble = [double]::TryParse(
        $stringValue,
        [System.Globalization.NumberStyles]::Float -bor
        [System.Globalization.NumberStyles]::AllowThousands,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$doubleValue
    )

    if ($isDouble) { return $doubleValue }

    <#
        The original value, not the string built above, so a non string input
        is handed back exactly as it came in.
    #>
    $Value
}