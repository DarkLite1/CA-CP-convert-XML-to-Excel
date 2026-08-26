#Requires -Version 7

<#
    The month key of a year and month, remembered after the first conversion.

    A run holds records for one or two months, so this never grows past a
    handful of entries while it saves a date format call per record.
#>
$script:MonthKeyCache = @{}

function Get-MonthKeyHC {
    <#
        .SYNOPSIS
            Convert a date to the key used for grouping and for the Excel file
            name, for example '202408 August'.

        .DESCRIPTION
            This key is used in two places that must always agree: the
            orchestrator uses it to decide which Excel file to write to, and the
            row builders use it to decide which records belong in that file.
            Keeping both on this single function means they can never drift
            apart.

            The month name is written with InvariantCulture, so it is English
            on every machine. Without it the name follows the regional setting
            of the server: the same August would be '202408 augustus' on a
            Dutch machine and '202408 août' on a French one. Because this key
            becomes the name of the Excel file, a server whose culture changed,
            or a second server with another regional setting, would not
            recognise the file of the running month and would quietly start an
            empty one next to it.

            This matches ConvertTo-DateTimeHC, which already reads dates with
            InvariantCulture so that a date means the same thing everywhere.
            Writing them follows the same rule.

            The result is cached per year and month. Formatting a date is not
            free, it walks the format string and reads the month names of the
            culture, and this ran once per record in a file. Every record of
            the same month produces the same key by definition, so it is worked
            out once.

        .EXAMPLE
            Get-MonthKeyHC -Date ([DateTime]'2024-08-16')   # '202408 August'
    #>
    param (
        [Parameter(Mandatory)]
        [DateTime]$Date
    )

    $cacheKey = ($Date.Year * 100) + $Date.Month

    $monthKey = $script:MonthKeyCache[$cacheKey]

    if ($monthKey) { return $monthKey }

    $monthKey = $Date.ToString('yyyyMM MMMM', [CultureInfo]::InvariantCulture)

    $script:MonthKeyCache[$cacheKey] = $monthKey

    $monthKey
}

function Get-MonthRangeHC {
    <#
        .SYNOPSIS
            Turn a month key back into the first moment of that month and the
            first moment of the next one.

        .DESCRIPTION
            Used by the row builders to decide whether a record belongs in the
            month being written. They compared month keys before, which meant
            formatting the date of every single record into a string and then
            comparing two strings, once per delivery, alarm and sequence
            parameter in the file. The range is worked out once per call and
            every record after that costs two date comparisons.

            'End' is the first moment of the NEXT month, so a record belongs to
            the month when it is at or after 'Start' and before 'End'. That
            leaves no gap at the end of the month, however precise the
            timestamp is.

            Only the leading 'yyyyMM' of the key is read. The month name that
            follows is there to make the Excel file name readable and adds
            nothing here, which also keeps this working whatever language the
            name was written in.

        .EXAMPLE
            Get-MonthRangeHC -MonthKey '202408 August'
            # Start : 2024-08-01 00:00:00
            # End   : 2024-09-01 00:00:00
    #>
    param (
        [Parameter(Mandatory)]
        [String]$MonthKey
    )

    if ($MonthKey.Length -lt 6) {
        throw "Month key '$MonthKey' does not start with 'yyyyMM'"
    }

    $year = 0
    $month = 0

    <#
        InvariantCulture here too, so the digits are read the same way on a
        machine with any regional setting.
    #>
    $numberStyle = [System.Globalization.NumberStyles]::None
    $culture = [CultureInfo]::InvariantCulture

    if (
        (-not [int]::TryParse($MonthKey.Substring(0, 4), $numberStyle, $culture, [ref]$year)) -or
        (-not [int]::TryParse($MonthKey.Substring(4, 2), $numberStyle, $culture, [ref]$month)) -or
        ($month -lt 1) -or ($month -gt 12)
    ) {
        throw "Month key '$MonthKey' does not start with 'yyyyMM'"
    }

    $start = [DateTime]::new($year, $month, 1)

    [PSCustomObject]@{
        Start = $start
        End   = $start.AddMonths(1)
    }
}