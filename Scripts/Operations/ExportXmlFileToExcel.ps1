#Requires -Version 7
#Requires -Modules ImportExcel

<#
    .SYNOPSIS
        Write the records of one month to one Excel file.

    .DESCRIPTION
        Run by path from 'Invoke-ConvertXmlToExcel', once per (month, Excel
        file), possibly in a parallel runspace. It imports the ConvertXmlToExcel
        module by path and uses its helpers to build and write the rows.

        Only the records belonging to 'MonthKey' are written. An XML file
        holding records of two months is handed to this script twice, once for
        each month, with a different 'ExcelFilePath' each time.

        An XML file is added completely or not at all. When the rows do not fit
        anymore because a worksheet reached the maximum number of rows, nothing
        is written for that file and an error is returned.

        This script does not move XML files to the archive folder. Because one
        XML file can be written to more than one Excel file, the file may only
        be archived after every Excel file has been handled. The orchestrator
        does that once all Excel files are done.

    .PARAMETER XmlFiles
        The XML files that hold records for 'MonthKey'.

    .PARAMETER ExcelFilePath
        The Excel file for 'MonthKey'.

    .PARAMETER Type
        'Batch', 'Alarm' or 'Sequence'.

    .PARAMETER MonthKey
        The month to write, for example '202408 August'. Only records of this
        month are written.

    .PARAMETER ModulePath
        Full path to 'ConvertXmlToExcel.psm1'.
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory)]
    [System.IO.FileSystemInfo[]]$XmlFiles,
    [Parameter(Mandatory)]
    [String]$ExcelFilePath,
    [Parameter(Mandatory)]
    [ValidateSet('Batch', 'Alarm', 'Sequence')]
    [String]$Type,
    [Parameter(Mandatory)]
    [String]$MonthKey,
    [Parameter(Mandatory)]
    [String]$ModulePath
)

Begin {
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

    $workbook = Open-ExcelWorkbookHC -Path $ExcelFilePath -Type $Type
}

Process {
    Write-Verbose "Excel file '$ExcelFilePath' month '$MonthKey'"

    $newRowAdded = $false

    $i = 0

    $fileNamesInExcelFile = Get-FileNameInWorkbookHC -Workbook $workbook

    foreach ($file in $XmlFiles | Sort-Object Name) {
        try {
            $result = [PSCustomObject]@{
                DateTime       = Get-Date
                File           = $file
                ExcelFilePath  = $ExcelFilePath
                MonthKey       = $MonthKey
                AddedToSheet   = $false
                AlreadyInSheet = $false
                RowsAdded      = 0
                Error          = $null
            }

            $i++

            #region Do not add file when already added
            if ($fileNamesInExcelFile[$file.Name]) {
                Write-Verbose "XML file $i/$($XmlFiles.Count) '$($file.Name)': already added to Excel file"

                $result.AlreadyInSheet = $true

                Continue
            }
            #endregion

            #region Read file
            Write-Verbose "XML file $i/$($XmlFiles.Count) '$($file.Name)': read XML file"

            $xmlDocument = Get-XmlDocumentHC -Path $file.FullName
            #endregion

            #region Get the rows for this month only
            $params = @{
                Xml      = $xmlDocument
                Type     = $Type
                MonthKey = $MonthKey
                FileName = $file.Name
                AddedOn  = $result.DateTime
            }
            $rows = @(Get-XmlRowHC @params)
            #endregion

            #region Nothing to add for this month
            if (-not $rows.Count) {
                Write-Verbose "XML file $i/$($XmlFiles.Count) '$($file.Name)': no records for month '$MonthKey'"

                Continue
            }
            #endregion

            Write-Verbose "XML file $i/$($XmlFiles.Count) '$($file.Name)': add $($rows.Count) rows"

            #region Add rows to the Excel file
            try {
                $params = @{
                    Workbook = $workbook
                    Rows     = $rows
                }
                $result.RowsAdded = Add-RowToWorkbookHC @params
            }
            catch {
                Remove-FileFromWorkbookHC -Workbook $workbook -FileName $file.Name
                throw $_
            }
            #endregion

            $result.AddedToSheet = $true
            $newRowAdded = $true
        }
        catch {
            $result.Error = $_
            $result.AddedToSheet = $false
            Write-Warning "XML file $i/$($XmlFiles.Count) '$($file.Name)': Failure: $_"
            if ($Error.Count) { $Error.RemoveAt(0) }
        }
        finally {
            $result
        }
    }

    if ($newRowAdded) {
        Format-ExcelWorkbookHC -Workbook $workbook
    }

    $workbook.Package.Dispose()
}
