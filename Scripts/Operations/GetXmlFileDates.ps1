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
    #>
    $moduleRoot = Split-Path $ModulePath -Parent

    Get-ChildItem -LiteralPath (Join-Path $moduleRoot 'Private') -Filter '*.ps1' -File |
    ForEach-Object { . $_.FullName }
    #endregion

    Write-Verbose "File '$XmlFile': get dates of type '$Type'"

    $xmlDocument = Get-XmlDocumentHC -Path $XmlFile.FullName

    $dates = Get-XmlFileMonthHC -Xml $xmlDocument -Type $Type

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
