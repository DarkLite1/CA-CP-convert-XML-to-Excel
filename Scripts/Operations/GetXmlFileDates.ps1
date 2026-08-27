#Requires -Version 7

<#
    .SYNOPSIS
        Return every month found in one XML file.

    .DESCRIPTION
        Run by path from 'Invoke-ConvertXmlToExcel', once per XML file, possibly
        in a parallel runspace. It imports the ConvertXmlToExcel module by path
        and uses its helpers to read the dates.

        Replaces the three original scripts 'Get loading date.ps1',
        'Get alarm raised date.ps1' and 'Get file creation date.ps1', which each
        returned a single date per XML file taken from the first or last record.
        That was wrong for files holding records of more than one month, because
        the whole file was then written to the Excel file of that one date.

        This script returns all months found in the file, so the orchestrator
        can send the file to every Excel file it has records for.

    .PARAMETER XmlFile
        The XML file to read.

    .PARAMETER Type
        'Batch'    : date taken from delivery.deliveryHeader.load_start_date
        'Alarm'    : date taken from alarm.raised
        'Sequence' : date taken from batchComputerHeader.file_created_on

    .PARAMETER ModulePath
        Full path to 'ConvertXmlToExcel.psm1'.

    .OUTPUTS
        File               : the XML file
        MonthKeys          : every month found, for example @('202408 August')
        RecordCount        : the number of records looked at
        RecordsWithoutDate : records that hold no usable date
        Error              : $null when the file could be read
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory)]
    [System.IO.FileSystemInfo]$XmlFile,
    [Parameter(Mandatory)]
    [ValidateSet('Batch', 'Alarm', 'Sequence')]
    [String]$Type,
    [Parameter(Mandatory)]
    [String]$ModulePath
)

try {
    $ErrorActionPreference = 'Stop'

    $result = [PSCustomObject]@{
        File               = $XmlFile
        MonthKeys          = @()
        RecordCount        = 0
        RecordsWithoutDate = 0
        Error              = $null
    }

    #region Load the module's private helper functions
    <#
        The helpers are private to the module, so importing the module would
        not expose them. This script is run by path, possibly in its own
        runspace, so it dot-sources the helper files it needs directly.

        Only the helpers this script actually calls are loaded, not every file
        in 'Private'. ForEach-Object -Parallel has no per-runspace setup step,
        so this whole script body runs once per XML file: loading all of them
        meant parsing every helper file again for every file processed. Worse,
        'ExcelWorkbook.ps1' carries a '#Requires -Modules ImportExcel', which
        pulled ImportExcel into runspaces that never write a single cell.

        Loaded EVERY time, without checking whether the functions are already
        there. When the orchestrator runs this script sequentially it does so
        from inside the module, so the module's own private copies are in scope
        here and any such check would find them and skip the loading. The
        functions would then resolve to the module's copies, whose script scope
        is the module's and not this one, and 'Get-MonthKeyHC' reads its cache
        from that scope. Dot-sourcing unconditionally puts the copies this
        script owns in front of them, which is what keeps it running against
        the files it was pointed at.
    #>
    $moduleRoot = Split-Path $ModulePath -Parent

    foreach (
        $helperName in @(
            'ConvertTo-DateTimeHC'
            'Get-MonthKeyHC'
            'Get-XmlFileMonthHC'
            'Get-XmlPropertyPathHC'
        )
    ) {
        . (Join-Path $moduleRoot "Private\$helperName.ps1")
    }
    #endregion

    Write-Verbose "File '$XmlFile': get dates of type '$Type'"

    <#
        The file is streamed, not loaded into an XmlDocument. The export step
        loads the document it needs itself, so building one here only to read
        one element per record was work thrown away.
    #>
    $dates = Get-XmlFileMonthHC -Path $XmlFile.FullName -Type $Type

    $result.MonthKeys = $dates.MonthKeys
    $result.RecordCount = $dates.RecordCount
    $result.RecordsWithoutDate = $dates.RecordsWithoutDate

    #region No date at all
    if (-not $result.MonthKeys.Count) {
        throw "No date found in XML property '$(Get-XmlPropertyPathHC -Type $Type)'"
    }
    #endregion

    #region Some records have no date
    if ($result.RecordsWithoutDate) {
        throw "$($result.RecordsWithoutDate) of $($result.RecordCount) records have no date in XML property '$(Get-XmlPropertyPathHC -Type $Type)'. These records cannot be placed in a monthly Excel file."
    }
    #endregion

    Write-Verbose "File '$XmlFile': found month(s) '$($result.MonthKeys -join "', '")'"
}
catch {
    $result.Error = $_

    if (-not (Test-Path $XmlFile.FullName -PathType Leaf)) {
        $result.Error = 'File removed during script execution'
    }

    Write-Warning "File '$XmlFile': Error retrieving dates: $($result.Error)"

    if ($Error.Count) { $Error.RemoveAt(0) }
}
finally {
    $result
}