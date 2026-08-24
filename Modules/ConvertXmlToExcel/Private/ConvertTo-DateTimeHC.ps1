#Requires -Version 7

function ConvertTo-DateTimeHC {
    <#
        .SYNOPSIS
            Convert an XML string value to a date.

        .DESCRIPTION
            Returns $null when the value is empty or not a valid date, so a
            missing or malformed date never stops the script and can be handled
            by the caller.

            The date is returned exactly as it is written in the XML file, with
            the time zone offset removed and never re-calculated. A value of
            '2025-11-01T00:30:00+01:00' is always the 1st of November at 00:30,
            on every machine.

            This matters for the month an XML record is filed under. A plain
            '[DateTime]' cast converts the value to the local time of the
            machine running the script: on a server set to UTC that same value
            becomes the 31st of October at 23:30, which would write the record
            to the Excel file of the previous month. Parsing as a
            [DateTimeOffset] and taking its '.DateTime' keeps the wall clock
            time of the plant that produced the file, so the result no longer
            depends on the server's time zone or on daylight saving time.

        .PARAMETER Value
            The text to convert, usually an ISO 8601 value coming from the XML
            file.

        .EXAMPLE
            ConvertTo-DateTimeHC -Value '2024-08-16T09:30:25+02:00'
            # 2024-08-16 09:30:25, on a machine in any time zone

        .EXAMPLE
            ConvertTo-DateTimeHC -Value ''
            # returns $null
    #>
    param (
        $Value
    )

    if (-not $Value) { return $null }

    $stringValue = [String]$Value

    if ([String]::IsNullOrWhiteSpace($stringValue)) { return $null }

    $dateTimeOffset = [DateTimeOffset]::MinValue

    <#
        InvariantCulture so a date is read the same way on a machine with a
        different regional setting. AssumeLocal only applies to values without
        an offset, and because '.DateTime' is used the assumed offset is
        discarded again, so those values are also returned as written.
    #>
    $isDate = [DateTimeOffset]::TryParse(
        $stringValue,
        [CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::AssumeLocal,
        [ref]$dateTimeOffset
    )

    if (-not $isDate) { return $null }

    $dateTimeOffset.DateTime
}