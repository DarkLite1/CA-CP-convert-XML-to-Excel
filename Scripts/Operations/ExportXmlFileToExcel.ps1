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

        Every XML file handed in returns one result, which reports the outcome
        for this month with one of 'AddedToSheet', 'AlreadyInSheet' or
        'NothingToAdd', or with 'Error'. 'NothingToAdd' means the file was read
        without a problem and holds no rows for this month.

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

        Loaded EVERY time, without checking whether the functions are already
        there. When the orchestrator runs this script sequentially it does so
        from inside the module, so the module's own private copies are in scope
        here and any such check would find them and skip the loading. The
        functions would then resolve to the module's copies, whose script scope
        is the module's and not this one. Dot-sourcing unconditionally puts the
        copies this script owns in front of them.
    #>
    $moduleRoot = Split-Path $ModulePath -Parent

    Get-ChildItem -LiteralPath (Join-Path $moduleRoot 'Private') -Filter '*.ps1' -File |
    ForEach-Object { . $_.FullName }
    #endregion

    function New-ResultHC {
        <#
            .SYNOPSIS
                One result for one XML file for this month.
        #>
        param (
            [Parameter(Mandatory)]$File,
            <#
                Untyped on purpose: a [String] parameter that is not given
                arrives as an empty string, and an empty string in 'Error' is
                still an answer to a question nobody asked. Left alone it is
                $null, which is what 'no error' looks like everywhere else.
            #>
            $ErrorMessage
        )

        [PSCustomObject]@{
            DateTime       = Get-Date
            File           = $File
            ExcelFilePath  = $ExcelFilePath
            MonthKey       = $MonthKey
            AddedToSheet   = $false
            AlreadyInSheet = $false
            NothingToAdd   = $false
            RowsAdded      = 0
            Error          = $ErrorMessage
        }
    }
}

Process {
    Write-Verbose "Excel file '$ExcelFilePath' month '$MonthKey'"

    #region Open the Excel file
    <#
        A workbook that cannot be opened is reported per file, like any other
        failure, instead of being thrown. This script runs inside the caller's
        ForEach-Object, so a throw here does not fail this month alone: it ends
        the whole pipeline, the caller never gets the results of the months that
        did succeed, and not one XML file of the entire run is archived.
    #>
    try {
        $workbook = Open-ExcelWorkbookHC -Path $ExcelFilePath -Type $Type
    }
    catch {
        Write-Warning "Excel file '$ExcelFilePath': Failed opening: $_"

        $openError = "Failed opening Excel file: $_"

        if ($Error.Count) { $Error.RemoveAt(0) }

        foreach ($file in $XmlFiles) {
            New-ResultHC -File $file -ErrorMessage $openError
        }

        return
    }
    #endregion

    $newRowAdded = $false

    $i = 0

    <#
        The results are collected instead of being handed back one at a time.
        Nothing is on disk until the workbook is saved at the end, so a file
        cannot be reported as added before that save is known to have worked.
        Handing a result back straight away told the caller a file was written
        while it still only lived in memory, and the caller archives on that
        answer.
    #>
    $results = [System.Collections.Generic.List[Object]]::new()

    try {
        $fileNamesInExcelFile = Get-FileNameInWorkbookHC -Workbook $workbook

        foreach ($file in $XmlFiles | Sort-Object Name) {
            try {
                $result = New-ResultHC -File $file

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
                <#
                    A month key can be found for a record whose children produce
                    no rows at all: a delivery loaded in this month that holds
                    no batches, or a batch computer created in this month that
                    holds no sequence parameters. The file was read correctly and
                    there is simply nothing to write, so this is a success and
                    not a gap.

                    'NothingToAdd' says so explicitly. Without it the
                    orchestrator sees a month that was neither added nor already
                    present, refuses to archive the file, and reports it again on
                    every following run, because nothing about the file will ever
                    change.
                #>
                if (-not $rows.Count) {
                    Write-Verbose "XML file $i/$($XmlFiles.Count) '$($file.Name)': no records for month '$MonthKey'"

                    $result.NothingToAdd = $true

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
                $result.RowsAdded = 0
                Write-Warning "XML file $i/$($XmlFiles.Count) '$($file.Name)': Failure: $_"
                if ($Error.Count) { $Error.RemoveAt(0) }
            }
            finally {
                $results.Add($result)
            }
        }

        #region Save the Excel file
        <#
            The save is what puts the rows on disk, so a save that fails means
            nothing was written, however well the rows themselves went. Every
            file that was added in this run is turned back into a failure here,
            so the caller leaves those XML files in the folder and picks them up
            again on the next run.

            A file that was already in the workbook keeps its answer: it was put
            there by an earlier run and is still on disk. A file that had nothing
            to add keeps its answer too, because there was nothing to save.
        #>
        if ($newRowAdded) {
            try {
                Format-ExcelWorkbookHC -Workbook $workbook
            }
            catch {
                Write-Warning "Excel file '$ExcelFilePath': Failed saving: $_"

                $saveError = "Failed saving the Excel file, the rows of this file were not written: $_"

                if ($Error.Count) { $Error.RemoveAt(0) }

                foreach ($result in $results) {
                    if (-not $result.AddedToSheet) { Continue }

                    $result.AddedToSheet = $false
                    $result.RowsAdded = 0
                    $result.Error = $saveError
                }
            }
        }
        #endregion
    }
    catch {
        <#
            Whatever is left: a helper that failed in a way this script does not
            know about. It is turned into a result per file rather than thrown,
            because this script runs inside the caller's ForEach-Object and a
            throw there ends the whole pipeline. The caller would lose the
            results of every month that did succeed and would archive nothing at
            all, for a problem in one workbook.
        #>
        Write-Warning "Excel file '$ExcelFilePath': Failure: $_"

        $unexpectedError = "Failed writing the Excel file: $_"

        if ($Error.Count) { $Error.RemoveAt(0) }

        $answeredFor = @{}

        foreach ($result in $results) {
            $answeredFor[$result.File.Name] = $true

            <#
                Nothing was saved, so a file reported as added has to be a
                failure here as well.
            #>
            if ($result.AddedToSheet) {
                $result.AddedToSheet = $false
                $result.RowsAdded = 0
                $result.Error = $unexpectedError
            }
        }

        foreach ($file in $XmlFiles) {
            if ($answeredFor[$file.Name]) { Continue }

            $results.Add((New-ResultHC -File $file -ErrorMessage $unexpectedError))
        }
    }
    finally {
        <#
            Always, so a failure never leaves the Excel file open. The package
            holds the file until it is disposed, and in a parallel run the next
            month waiting for that same workbook would find it locked by this
            one.
        #>
        $workbook.Package.Dispose()
    }

    $results
}