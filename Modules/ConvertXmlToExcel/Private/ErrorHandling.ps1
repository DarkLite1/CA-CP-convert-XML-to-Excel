#Requires -Version 7

<#
    .SYNOPSIS
        Collect script level errors without throwing.

    .DESCRIPTION
        Taken from the Permission-matrix project so this repository has no
        dependency on an internal error handling module.
#>

function Add-ErrorHC {
    <#
    .SYNOPSIS
        Append a structured error record to a system-error accumulator.

    .DESCRIPTION
        Records rather than throws: the record is appended to SystemErrors and
        the caller decides how to proceed.

    .PARAMETER SystemErrors
        A [ref] to a collection exposing .Add(), for example
        [System.Collections.Generic.List[object]]. An array created with @() is
        fixed-size and causes a terminating error. Prefer a generic List over an
        ArrayList: ArrayList.Add() returns the insertion index, which would leak
        onto the pipeline.

    .NOTES
        Type is restricted to 'FatalError'/'Warning' because callers decide
        whether to halt by testing Type -eq 'FatalError'. A typo such as 'Fatal'
        would never match and would silently downgrade the error to advisory.

    .EXAMPLE
        $errors = [System.Collections.Generic.List[object]]::new()
        Add-ErrorHC -Type 'FatalError' -Name 'Bad row' -Message 'Missing path.' -Category 'Matrix' -SystemErrors ([ref]$errors)
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('FatalError', 'Warning')]
        [string]$Type,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Message,
        [Parameter()][string]$Description = '',
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    $SystemErrors.Value.Add(
        [PSCustomObject]@{
            DateTime    = Get-Date
            Type        = $Type
            Name        = $Name
            Message     = $Message
            Description = $Description
            Category    = $Category
        }
    )
}
