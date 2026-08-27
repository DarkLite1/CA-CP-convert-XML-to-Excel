#Requires -Version 7

function ConvertTo-TextHC {
    <#
        .SYNOPSIS
            Keep an XML value as text.

        .DESCRIPTION
            The counterpart of ConvertTo-CorrectTypeHC, for values that only
            look like a number but are not one: identifiers, codes, reference
            numbers, license plates and free text.

            Handing those to ConvertTo-CorrectTypeHC damages them in two ways
            that cannot be undone once the Excel file is written:

            - Leading zeros are lost. An order number '0000123456' coming from
              the source system becomes 123456 and no longer matches the system
              it came from.
            - Precision is lost above 2^53. An 18 digit identifier becomes
              1.23456789012346E+17, so the last digits are simply gone.

            Empty and null values return nothing, exactly like
            ConvertTo-CorrectTypeHC does, so an empty element leaves an empty
            cell instead of a cell holding an empty string.

            The value is returned as a [String] on purpose. An XML value is
            text to begin with, so this is what it already was; saying it here
            makes the intent visible at the call site and stops Excel from
            guessing.

        .PARAMETER Value
            The text to keep, usually a leaf value coming from the XML file.

        .EXAMPLE
            ConvertTo-TextHC -Value '0000123456'   # returns '0000123456'

        .EXAMPLE
            ConvertTo-TextHC -Value ''             # returns nothing
    #>
    Param(
        $Value
    )

    if ($null -eq $Value) { return }

    $stringValue = [String]$Value

    if ($stringValue.Length -eq 0) { return }

    $stringValue
}