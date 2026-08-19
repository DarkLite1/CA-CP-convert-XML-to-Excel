#Requires -Version 7
#Requires -Modules ImportExcel

function Invoke-ConvertXmlToExcel {
    <#
    .SYNOPSIS
        Read XML files in a folder and export their content to Excel files, one
        Excel file per month.

    .DESCRIPTION
        Main orchestrator for the conversion. Replaces the three original
        'Main.ps1' scripts of the batch, alarm and sequence repositories. The
        type of XML file is set with the property 'Type' in the .json input
        file; everything else is the same for all three types.

        Requires the internal modules 'Toolbox.EventLog' and 'Toolbox.HTML' at
        run time for event logging and mail. They are intentionally not declared
        with '#Requires' here so the module still imports (and its private
        helpers stay unit-testable) on machines where those internal modules are
        not installed.

        One XML file can hold records of more than one month (for example a
        batch file with a delivery loaded today and another loaded tomorrow,
        where tomorrow is the first day of the next month). Such a file is
        written to more than one Excel file, each getting only the records of
        its own month.

        An XML file is only moved to the archive folder when all of its records
        were written without a single error. When something goes wrong the file
        stays in place and is reported in the summary mail, so it can be looked
        at and retried on the next run.

    .PARAMETER ScriptName
        Name used for the event log source and the log file names.

    .PARAMETER ImportFile
        The .json file with the script settings.

    .PARAMETER ScriptPath
        Hashtable with the full paths to the operation scripts. Must contain:
        - GetXmlFileDates
        - ExportToExcel

    .PARAMETER ModulePath
        Full path to 'ConvertXmlToExcel.psm1', passed on to the operation
        scripts so they can import the module in their own runspace.

    .EXAMPLE
        Invoke-ConvertXmlToExcel -ScriptName 'CP Batch' -ImportFile 'C:\cfg.json' `
            -ScriptPath @{
                GetXmlFileDates = 'C:\repo\Scripts\Operations\GetXmlFileDates.ps1'
                ExportToExcel   = 'C:\repo\Scripts\Operations\ExportXmlFileToExcel.ps1'
            } `
            -ModulePath 'C:\repo\Modules\ConvertXmlToExcel\ConvertXmlToExcel.psm1'
    #>
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [String]$ScriptName,
        [Parameter(Mandatory)]
        [String]$ImportFile,
        [Parameter(Mandatory)]
        [HashTable]$ScriptPath,
        [Parameter(Mandatory)]
        [String]$ModulePath,
        [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\TMS CP\$ScriptName",
        [String[]]$ScriptAdmin = @(
            $env:POWERSHELL_SCRIPT_ADMIN,
            $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
        )
    )

    Begin {
        try {
            Get-ScriptRuntimeHC -Start
            Import-EventLogParamsHC -Source $ScriptName
            Write-EventLog @EventStartParams
            $Error.Clear()

            #region Test path exists
            $scriptPathItem = @{}

            $ScriptPath.GetEnumerator().ForEach(
                {
                    try {
                        $key = $_.Key
                        $value = $_.Value

                        $params = @{
                            Path        = $value
                            ErrorAction = 'Stop'
                        }
                        $scriptPathItem[$key] = (Get-Item @params).FullName
                    }
                    catch {
                        throw "ScriptPath.$key '$value' not found"
                    }
                }
            )

            try {
                $modulePathItem = (Get-Item -Path $ModulePath -EA Stop).FullName
            }
            catch {
                throw "ModulePath '$ModulePath' not found"
            }
            #endregion

            #region Create log folder
            try {
                $logParams = @{
                    LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                    Name         = $ScriptName
                    Date         = 'ScriptStartTime'
                    NoFormatting = $true
                }
                $logFile = New-LogFileNameHC @LogParams
            }
            Catch {
                throw "Failed creating the log folder '$LogFolder': $_"
            }
            #endregion

            #region Import .json file
            Write-Verbose "Import .json file '$ImportFile'"

            $importFileContent = Get-Content $ImportFile -Raw -EA Stop -Encoding UTF8 |
            ConvertFrom-Json
            #endregion

            #region Test .json file properties
            Write-Verbose 'Test .json file properties'

            try {
                @(
                    'MaxConcurrentJobs', 'SendMail', 'Path', 'ExcelFileName', 'Type'
                ).where(
                    { -not $importFileContent.$_ }
                ).foreach(
                    { throw "Property '$_' not found" }
                )

                #region Test Type
                if ($importFileContent.Type -notMatch '^Batch$|^Alarm$|^Sequence$') {
                    throw "Property 'Type' with value '$($importFileContent.Type)' is not valid. Accepted values are 'Batch', 'Alarm' or 'Sequence'"
                }
                #endregion

                @(
                    'GetXmlFileDates', 'CreateExcelFile'
                ).foreach(
                    {
                        $key = $_
                        $value = $importFileContent.MaxConcurrentJobs.$_
                        if (-not $value) {
                            throw "Property 'MaxConcurrentJobs.$key' not found"
                        }

                        #region Test integer value
                        try {
                            $null = [int]$value
                        }
                        catch {
                            throw "Property 'MaxConcurrentJobs.$key' needs to be a number, the value '$value' is not supported."
                        }
                        #endregion
                    }
                )

                @(
                    'XmlFiles', 'ExcelFiles'
                ).foreach(
                    {
                        if (-not $importFileContent.Path.$_) {
                            throw "Property 'Path.$_' not found"
                        }
                        if (-not (Test-Path -LiteralPath $importFileContent.Path.$_ -PathType Container)) {
                            throw "Folder '$($importFileContent.Path.$_)' for 'Path.$_' not found"
                        }
                    }
                )

                #region Test SendMail
                @('To', 'When').Where(
                    { -not $importFileContent.SendMail.$_ }
                ).foreach(
                    { throw "Property 'SendMail.$_' not found" }
                )

                if ($importFileContent.SendMail.When -notMatch '^Never$|^Always$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                    throw "Property 'SendMail.When' with value '$($importFileContent.SendMail.When)' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
                }
                #endregion

                #region Test ExcelFileName
                if (-not ($importFileContent.ExcelFileName -like '*{0}*')) {
                    throw "Property 'ExcelFileName' with value '$($importFileContent.ExcelFileName)' is missing the date placeholder '{0}'"
                }
                if (-not ($importFileContent.ExcelFileName -like '*.xlsx')) {
                    throw "Property 'ExcelFileName' with value '$($importFileContent.ExcelFileName)' is not ending with file extension '.xlsx'"
                }
                #endregion
            }
            catch {
                throw "Input file '$ImportFile': $_"
            }
            #endregion
        }
        catch {
            Write-Warning $_
            Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
            Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
            Write-EventLog @EventEndParams; Exit 1
        }
    }

    Process {
        try {
            $XmlFilesFolder = $importFileContent.Path.XmlFiles
            $ExcelFilesFolder = $importFileContent.Path.ExcelFiles
            $ExcelFileName = $importFileContent.ExcelFileName
            $Type = $importFileContent.Type

            #region Get XML files
            $M = "Get xml files in folder '$XmlFilesFolder'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $xmlFiles = Get-ChildItem $XmlFilesFolder -Filter '*.xml' -File |
            Sort-Object CreationTime

            $M = "Found $($xmlFiles.Count) xml files of type '$Type'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            #endregion

            #region Get the months in each XML file
            $scriptBlock = {
                $file = $_

                #region Declare variables for code running in parallel
                if (-not $MaxConcurrentJobs) {
                    $scriptPathItem = $using:scriptPathItem
                    $modulePathItem = $using:modulePathItem
                    $Type = $using:Type
                    $EventVerboseParams = $using:EventVerboseParams
                }
                #endregion

                #region Start job
                $invokeParams = @{
                    FilePath     = $scriptPathItem.GetXmlFileDates
                    ArgumentList = $file, $Type, $modulePathItem
                }

                $params = $invokeParams.ArgumentList
                & $invokeParams.FilePath @params
                #endregion
            }

            #region Run code serial or parallel
            $MaxConcurrentJobs = $importFileContent.MaxConcurrentJobs.GetXmlFileDates

            $foreachParams = if ($MaxConcurrentJobs -eq 1) {
                @{
                    Process = $scriptBlock
                }
            }
            else {
                @{
                    Parallel      = $scriptBlock
                    ThrottleLimit = $MaxConcurrentJobs
                }
            }

            Write-Verbose "Get dates from $($xmlFiles.Count) files"

            $fileDates = @($xmlFiles | ForEach-Object @foreachParams)

            Write-Verbose "Retrieved dates from $($fileDates.Count) files"
            #endregion
            #endregion

            #region Group the XML files per month
            <#
                One XML file can be in more than one group, because it can hold
                records of more than one month. Each group is written to its own
                Excel file and only gets the records of its own month.
            #>
            $filesPerMonth = @{}

            foreach ($fileDate in $fileDates.where({ -not $_.Error })) {
                foreach ($monthKey in $fileDate.MonthKeys) {
                    if (-not $filesPerMonth.ContainsKey($monthKey)) {
                        $filesPerMonth[$monthKey] = [System.Collections.Generic.List[Object]]::new()
                    }
                    $filesPerMonth[$monthKey].Add($fileDate.File)
                }
            }

            $groups = @(
                $filesPerMonth.GetEnumerator() | Sort-Object Name | ForEach-Object {
                    @{
                        MonthKey      = $_.Key
                        Files         = @($_.Value)
                        ExcelFilePath = Join-Path $ExcelFilesFolder (
                            $ExcelFileName -f $_.Key
                        )
                    }
                }
            )

            $M = "Found $($groups.Count) different months in $($xmlFiles.Count) XML files"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $fanOutFiles = @(
                $fileDates.where({ (-not $_.Error) -and ($_.MonthKeys.Count -gt 1) })
            )

            if ($fanOutFiles.Count) {
                $M = "$($fanOutFiles.Count) XML file(s) hold records of more than one month and are written to more than one Excel file"
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            }
            #endregion

            #region Create Excel files
            $scriptBlock = {
                $group = $_

                #region Declare variables for code running in parallel
                if (-not $MaxConcurrentJobs) {
                    $scriptPathItem = $using:scriptPathItem
                    $modulePathItem = $using:modulePathItem
                    $Type = $using:Type
                    $EventVerboseParams = $using:EventVerboseParams
                }
                #endregion

                #region Start job
                $invokeParams = @{
                    FilePath     = $scriptPathItem.ExportToExcel
                    ArgumentList = $group.Files, $group.ExcelFilePath, $Type,
                    $group.MonthKey, $modulePathItem
                }

                $params = $invokeParams.ArgumentList
                & $invokeParams.FilePath @params
                #endregion
            }

            #region Run code serial or parallel
            $MaxConcurrentJobs = $importFileContent.MaxConcurrentJobs.CreateExcelFile

            $foreachParams = if ($MaxConcurrentJobs -eq 1) {
                @{
                    Process = $scriptBlock
                }
            }
            else {
                @{
                    Parallel      = $scriptBlock
                    ThrottleLimit = $MaxConcurrentJobs
                }
            }

            Write-Verbose 'Start script to export data to Excel'

            $exportExcelResults = @($groups | ForEach-Object @foreachParams)
            #endregion
            #endregion

            #region Archive XML files
            <#
                Archiving is done here and not in the export script, because one XML
                file can be written to more than one Excel file. The file may only
                be moved once every Excel file it belongs to has been handled.

                A file with an error is never archived. It stays in the folder and
                is reported in the mail, so a manual action can be taken. When the
                problem is solved the next run picks the file up again. The Excel
                files it was already added to are not written twice, because the
                file name is checked before adding.
            #>
            $archiveFolder = Join-Path $XmlFilesFolder 'Archive'

            $archivedFiles = @()
            $notArchivedFiles = @()

            foreach ($fileDate in $fileDates) {
                $fileName = $fileDate.File.Name

                $fileResults = @(
                    $exportExcelResults.where({ $_.File.Name -eq $fileName })
                )

                $reason = $null

                if ($fileDate.Error) {
                    $reason = "No date could be read from the XML file: $($fileDate.Error)"
                }
                elseif ($fileResults.where({ $_.Error })) {
                    $reason = @(
                        $fileResults.where({ $_.Error }).ForEach(
                            { "'$(Split-Path $_.ExcelFilePath -Leaf)': $($_.Error)" }
                        )
                    ) -join ' | '
                }
                elseif (
                    $fileResults.where(
                        { $_.AddedToSheet -or $_.AlreadyInSheet }
                    ).Count -ne $fileDate.MonthKeys.Count
                ) {
                    $reason = 'Not all months of this XML file were written to an Excel file'
                }

                if ($reason) {
                    $notArchivedFiles += [PSCustomObject]@{
                        File   = $fileDate.File
                        Reason = $reason
                    }

                    Write-Warning "File '$fileName' not archived: $reason"

                    Continue
                }

                #region Move file to the archive folder
                try {
                    $null = New-Item -Path $archiveFolder -ItemType Directory -EA Ignore

                    if (Test-Path -LiteralPath $fileDate.File.FullName -PathType Leaf) {
                        Write-Verbose "Move file '$fileName' to '$archiveFolder'"

                        Move-Item -LiteralPath $fileDate.File.FullName -Destination $archiveFolder -EA Stop
                    }

                    $archivedFiles += $fileDate.File
                }
                catch {
                    $notArchivedFiles += [PSCustomObject]@{
                        File   = $fileDate.File
                        Reason = "Failed moving file to archive folder '$archiveFolder': $_"
                    }
                    $Error.RemoveAt(0)
                }
                #endregion
            }
            #endregion

            Write-Verbose 'All done'
        }
        catch {
            Write-Warning $_
            Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
            Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
            Write-EventLog @EventEndParams; Exit 1
        }
    }

    End {
        try {
            #region Get results
            $collection = @{
                ExportXmlToExcel = @{
                    Results        = $exportExcelResults
                    AddedToSheet   = $exportExcelResults.where({ $_.AddedToSheet })
                    AlreadyInSheet = $exportExcelResults.where({ $_.AlreadyInSheet })
                    Errors         = $exportExcelResults.where({ $_.Error })
                }
                FileDates        = @{
                    Errors = $fileDates.where({ $_.Error })
                    FanOut = $fanOutFiles
                }
                Archive          = @{
                    Archived    = $archivedFiles
                    NotArchived = $notArchivedFiles
                }
                System           = @{
                    Errors = $Error.Exception.Message.where({ $_ })
                }
            }
            #endregion

            #region Count
            $count = @{
                ExportXmlToExcel = @{
                    Results        = $collection.ExportXmlToExcel.Results.Count
                    Errors         = $collection.ExportXmlToExcel.Errors.Count
                    AddedToSheet   = $collection.ExportXmlToExcel.AddedToSheet.Count
                    AlreadyInSheet = $collection.ExportXmlToExcel.AlreadyInSheet.Count
                }
                FileDates        = @{
                    Errors = $collection.FileDates.Errors.Count
                    FanOut = $collection.FileDates.FanOut.Count
                }
                Archive          = @{
                    Archived    = $collection.Archive.Archived.Count
                    NotArchived = $collection.Archive.NotArchived.Count
                }
                System           = @{
                    Errors = $collection.System.Errors.Count
                }
                Total            = @{
                    XmlFiles    = $xmlFiles.Count
                    ExcelFiles  = $groups.Count
                }
            }

            $count.Total.Errors = $count.ExportXmlToExcel.Errors +
            $count.FileDates.Errors + $count.System.Errors
            #endregion

            $mailParams = @{}

            $excelParams = @{
                Path          = "$logFile - Log.xlsx"
                AutoNameRange = $true
                Append        = $true
                AutoSize      = $true
                FreezeTopRow  = $true
                Verbose       = $false
            }

            #region Export results to Excel
            if ($count.ExportXmlToExcel.Results) {
                $excelParams.WorksheetName = $excelParams.TableName = 'Overview'

                $M = "Export {0} rows to Excel sheet '{1}'" -f
                $count.ExportXmlToExcel.Results, $excelParams.WorksheetName
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                $collection.ExportXmlToExcel.Results |
                Select-Object -Property DateTime, ExcelFilePath, MonthKey, @{
                    Name       = 'File'
                    Expression = { $_.File.Name }
                },
                AlreadyInSheet, AddedToSheet, RowsAdded, Error |
                Export-Excel @excelParams

                $mailParams.Attachments = $excelParams.Path
            }
            #endregion

            #region Export files that were not archived to Excel
            if ($count.Archive.NotArchived) {
                $excelParams.WorksheetName = $excelParams.TableName = 'NotArchived'

                $M = "Export {0} rows to Excel sheet '{1}'" -f
                $count.Archive.NotArchived, $excelParams.WorksheetName
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                $collection.Archive.NotArchived | Select-Object -Property @{
                    Name       = 'File'
                    Expression = { $_.File.FullName }
                }, Reason | Export-Excel @excelParams

                $mailParams.Attachments = $excelParams.Path
            }
            #endregion

            #region Export errors to Excel
            if ($count.System.Errors) {
                $excelParams.WorksheetName = $excelParams.TableName = 'SystemErrors'

                $M = "Export {0} rows to Excel sheet '{1}'" -f
                $count.System.Errors, $excelParams.WorksheetName
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                $collection.System.Errors | Export-Excel @excelParams

                $mailParams.Attachments = $excelParams.Path
            }

            if ($count.FileDates.Errors) {
                $excelParams.WorksheetName = $excelParams.TableName = 'DateErrors'

                $M = "Export {0} rows to Excel sheet '{1}'" -f
                $count.FileDates.Errors, $excelParams.WorksheetName
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                $collection.FileDates.Errors | Select-Object -Property @{
                    Name       = 'File'
                    Expression = { $_.File.FullName }
                }, RecordCount, RecordsWithoutDate, Error | Export-Excel @excelParams

                $mailParams.Attachments = $excelParams.Path
            }
            #endregion

            #region Mail subject and priority
            $mailParams += @{
                Priority = 'Normal'
                Subject  = @(
                    '{0} file{1} added' -f
                    $count.ExportXmlToExcel.AddedToSheet,
                    $(if ($count.ExportXmlToExcel.AddedToSheet -ne 1) { 's' })
                )
            }

            if ($count.Archive.NotArchived) {
                $mailParams.Subject += '{0} file{1} not archived' -f
                $count.Archive.NotArchived,
                $(if ($count.Archive.NotArchived -ne 1) { 's' })
            }

            if ($count.Total.Errors) {
                $mailParams.Subject += '{0} error{1}' -f
                $count.Total.Errors,
                $(if ($count.Total.Errors -ne 1) { 's' })

                $mailParams.Priority = 'High'
            }

            $mailParams.Subject = $mailParams.Subject -join ', '
            #endregion

            #region Check to send mail to user
            Write-Verbose 'Check if mail needs to be send'

            $sendMailToUser = $false

            if (
                (
                    $importFileContent.SendMail.When -eq 'Always'
                ) -or
                (
                    ($importFileContent.SendMail.When -eq 'OnlyOnError') -and
                    ($count.Total.Errors)
                ) -or
                (
                    ($importFileContent.SendMail.When -eq 'OnlyOnErrorOrAction') -and
                    (
                        ($count.Total.Errors) -or
                        ($count.ExportXmlToExcel.AddedToSheet) -or
                        ($count.Archive.NotArchived)
                    )
                )
            ) {
                $sendMailToUser = $true
            }
            #endregion

            #region Create HTML table
            Write-Verbose 'Create HTML table'

            $htmlTable = "
            <table>
                <tr>
                    <th>Quantity</th>
                    <th>Description</th>
                </tr>
                $(
                    '<tr>' +
                        "<td>$($count.Total.XmlFiles)</td>" +
                        '<td>Total XML file{0}</td>' -f
                        $(if ($count.Total.XmlFiles -ne 1) {'s'}) +
                    '</tr>'
                )
                $(
                    if ($count.ExportXmlToExcel.AddedToSheet) {
                        '<tr>' +
                            "<td>$($count.ExportXmlToExcel.AddedToSheet)</td>" +
                            '<td>XML file{0} added to Excel:{1}</td>' -f $(
                                if ($count.ExportXmlToExcel.AddedToSheet -ne 1) {'s'}
                                ),
                                $(
                                    ($collection.ExportXmlToExcel.AddedToSheet | Group-Object 'ExcelFilePath' | ForEach-Object {
                                        '<br>- {0}: {1}' -f
                                            $_.Count, (Split-Path $_.Name -Leaf)
                                    }) -join ''
                                ) +
                        '</tr>'
                    }
                )
                $(
                    if ($count.FileDates.FanOut) {
                        '<tr>' +
                            "<td>$($count.FileDates.FanOut)</td>" +
                            '<td>XML file{0} with records in more than one month, written to more than one Excel file</td>' -f $(
                                if ($count.FileDates.FanOut -ne 1) {'s'}
                            ) +
                        '</tr>'
                    }
                )
                $(
                    if ($count.ExportXmlToExcel.AlreadyInSheet) {
                        '<tr>' +
                            "<td>$($count.ExportXmlToExcel.AlreadyInSheet)</td>" +
                            '<td>XML file{0} already in Excel</td>' -f $(
                                if ($count.ExportXmlToExcel.AlreadyInSheet -ne 1) {'s'}
                                ) +
                        '</tr>'
                    }
                )
                $(
                    if ($count.Archive.Archived) {
                        '<tr>' +
                            "<td>$($count.Archive.Archived)</td>" +
                            '<td>XML file{0} moved to the archive folder</td>' -f $(
                                if ($count.Archive.Archived -ne 1) {'s'}
                            ) +
                        '</tr>'
                    }
                )
                $(
                    if ($count.Archive.NotArchived) {
                        '<tr style="background-color: #f78474">' +
                            "<td>$($count.Archive.NotArchived)</td>" +
                            '<td><b>XML file{0} NOT archived, manual action required:</b>{1}</td>' -f $(
                                if ($count.Archive.NotArchived -ne 1) {'s'}
                                ),
                                $(
                                    ($collection.Archive.NotArchived | ForEach-Object {
                                        '<br>- {0}: {1}' -f $_.File.Name, $_.Reason
                                    }) -join ''
                                ) +
                        '</tr>'
                    }
                )
                $(
                    if ($count.System.Errors) {
                        '<tr style="background-color: #f78474">' +
                            "<td>$($count.System.Errors)</td>" +
                            '<td>System error{0}</td>' -f $(
                                if ($count.System.Errors -ne 1) {'s'}
                            ) +
                        '</tr>'
                    }
                )
                $(
                    if ($count.FileDates.Errors) {
                        '<tr style="background-color: #f78474">' +
                            "<td>$($count.FileDates.Errors)</td>" +
                            '<td>Date error{0}</td>' -f $(
                                if ($count.FileDates.Errors -ne 1) {'s'}
                            ) +
                        '</tr>'
                    }
                )
                $(
                    if ($count.ExportXmlToExcel.Errors) {
                        '<tr style="background-color: #f78474">' +
                            "<td>$($count.ExportXmlToExcel.Errors)</td>" +
                            '<td>Export XML file to Excel error{0}</td>' -f $(
                                if ($count.ExportXmlToExcel.Errors -ne 1) {'s'}
                            ) +
                        '</tr>'
                    }
                )
            </table>

            <p>Folder locations:</p>
            <table>
                <tr>
                    <th>XML files</th>
                    <td>$('<a href="{0}">{0}</a>' -f $importFileContent.Path.XmlFiles)</td>
                </tr>
                <tr>
                    <th>Excel files</th>
                    <td>$('<a href="{0}">{0}</a>' -f $importFileContent.Path.ExcelFiles)</td>
                </tr>
            </table>
            "
            #endregion

            #region Send mail
            $mailParams += @{
                To             = $importFileContent.SendMail.To
                Message        = "
                    <p>Summary of XML file conversion (type '$($importFileContent.Type)'):</p> $htmlTable"
                LogFolder      = $LogParams.LogFolder
                Header         = $ScriptName
                EventLogSource = $ScriptName
                Save           = $LogFile + ' - Mail.html'
                ErrorAction    = 'Stop'
            }

            if ($mailParams.Attachments) {
                $mailParams.Message +=
                "<p><i>* Check the attachment for details</i></p>"
            }

            Get-ScriptRuntimeHC -Stop

            if ($sendMailToUser) {
                Write-Verbose 'Send e-mail to the user'

                if ($count.Total.Errors) {
                    $mailParams.Bcc = $ScriptAdmin
                }
                Send-MailHC @mailParams
            }
            else {
                Write-Verbose 'Send no e-mail to the user'

                if ($count.Total.Errors) {
                    Write-Verbose 'Send e-mail to admin only with errors'

                    $mailParams.To = $ScriptAdmin
                    Send-MailHC @mailParams
                }
            }
            #endregion
        }
        catch {
            Write-Warning $_
            Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
            Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
            Exit 1
        }
        finally {
            Write-EventLog @EventEndParams
        }
    }
}
