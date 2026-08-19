#Requires -Version 7

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

        .EXAMPLE
            Get-MonthKeyHC -Date ([DateTime]'2024-08-16')   # '202408 August'
    #>
    param (
        [Parameter(Mandatory)]
        [DateTime]$Date
    )

    $Date.ToString('yyyyMM MMMM')
}
