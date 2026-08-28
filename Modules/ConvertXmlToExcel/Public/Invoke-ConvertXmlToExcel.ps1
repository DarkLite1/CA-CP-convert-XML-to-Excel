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
        type of XML file is set with the property 'Type' in the configuration
        file; everything else is the same for all three types.

        One XML file can hold records of more than one month (for example a
        batch file with a delivery loaded today and another loaded tomorrow,
        where tomorrow is the first day of the next month). Such a file is
        written to more than one Excel file, each getting only the records of
        its own month.

        An XML file is only moved to the archive folder when all of its records
        were written without a single error. When something goes wrong the file
        stays in place and is reported in the summary mail, so it can be looked
        at and retried on the next run.

        Event logging, log files and mail are done by the module's own helper
        functions, taken from the Permission-matrix project. ImportExcel is the
        only external module dependency; mail uses the MailKit and MimeKit
        assemblies whose paths come from the configuration file.

    .PARAMETER ConfigurationJsonFile
        The .json configuration file, based on a file in 'Examples'.

    .PARAMETER ScriptPath
        Hashtable with the full paths to the operation scripts. Must contain
        'GetXmlFileDates' and 'ExportToExcel'.

    .PARAMETER ModulePath
        Full path to 'ConvertXmlToExcel.psm1', passed on to the operation
        scripts so they can load the private helpers in their own runspace.

    .PARAMETER SystemErrors
        A [ref] to a List[object] collecting the script level errors of the run.
        The caller reads it to decide the exit code.

    .EXAMPLE
        $errors = [System.Collections.Generic.List[object]]::new()

        Invoke-ConvertXmlToExcel -ConfigurationJsonFile 'C:\Batch.json' `
            -ScriptPath @{
                GetXmlFileDates = 'C:\repo\Scripts\Operations\GetXmlFileDates.ps1'
                ExportToExcel   = 'C:\repo\Scripts\Operations\ExportXmlFileToExcel.ps1'
            } `
            -ModulePath 'C:\repo\Modules\ConvertXmlToExcel\ConvertXmlToExcel.psm1' `
            -SystemErrors ([ref]$errors)
    #>
    [CmdLetBinding()]
    Param (
        [Parameter(Mandatory)]
        [String]$ConfigurationJsonFile,
        [Parameter(Mandatory)]
        [HashTable]$ScriptPath,
        [Parameter(Mandatory)]
        [String]$ModulePath,
        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    $scriptStartTime = Get-Date

    $config = $null
    $datedLogFolder = $null
    $logFileAttachments = @()
    $fileDates = @()
    $xmlFiles = @()
    $groups = @()

    <#
        'Archived', 'NotArchived' and 'Duplicates' are appended to one item at
        a time in the archive loop below. A PowerShell array cannot grow, so
        '$array += $item' allocates a new array and copies every item into it,
        which turns filling one into an operation that costs more the more
        files there are. A List grows in place.

        The others are assigned once from a pipeline and stay arrays.
    #>
    $collection = @{
        Results        = @()
        Added          = @()
        AlreadyInSheet = @()
        FanOut         = @()
        Archived       = [System.Collections.Generic.List[Object]]::new()
        NotArchived    = [System.Collections.Generic.List[Object]]::new()
        Duplicates     = [System.Collections.Generic.List[Object]]::new()
    }

    <#
        The archive result per XML file name, so the log file can show what
        happened to each file in its own row.
    #>
    $archiveOutcome = @{}

    <#
        The results of each XML file, looked up by file name. Filled after the
        export runs.

        Declared here and not where it is filled, because the overview at the
        end of this function runs after the catch block and reads it even when
        the export threw before it could be filled. An empty hash then gives
        the same answer the old full scan of an empty result list gave.
    #>
    $resultsByFileName = @{}

    try {
        #region Import the configuration file
        Write-Verbose "Import configuration file '$ConfigurationJsonFile'"

        try {
            $config = Get-Content $ConfigurationJsonFile -Raw -Encoding UTF8 -EA Stop |
            ConvertFrom-Json -EA Stop
        }
        catch {
            throw "Configuration file '$ConfigurationJsonFile' could not be read: $_"
        }
        #endregion

        #region Test the operation scripts and the module exist
        $scriptPathItem = @{}

        $ScriptPath.GetEnumerator().ForEach(
            {
                $key = $_.Key
                $value = $_.Value

                try {
                    $scriptPathItem[$key] = (Get-Item -Path $value -EA Stop).FullName
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

        #region Test the configuration file properties
        Write-Verbose 'Test configuration file properties'

        try {
            @(
                'Type', 'MaxConcurrentJobs', 'Path', 'ExcelFileName', 'Settings'
            ).where(
                { -not $config.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            if ($config.Type -notMatch '^Batch$|^Alarm$|^Sequence$') {
                throw "Property 'Type' with value '$($config.Type)' is not valid. Accepted values are 'Batch', 'Alarm' or 'Sequence'"
            }

            @('GetXmlFileDates', 'CreateExcelFile').foreach(
                {
                    $key = $_
                    $value = $config.MaxConcurrentJobs.$_

                    if (-not $value) {
                        throw "Property 'MaxConcurrentJobs.$key' not found"
                    }

                    try {
                        $null = [int]$value
                    }
                    catch {
                        throw "Property 'MaxConcurrentJobs.$key' needs to be a number, the value '$value' is not supported."
                    }
                }
            )

            @('XmlFiles', 'ExcelFiles').foreach(
                {
                    if (-not $config.Path.$_) {
                        throw "Property 'Path.$_' not found"
                    }
                    if (-not (Test-Path -LiteralPath $config.Path.$_ -PathType Container)) {
                        throw "Folder '$($config.Path.$_)' for 'Path.$_' not found"
                    }
                }
            )

            if ($config.ExcelFileName -notlike '*{0}*') {
                throw "Property 'ExcelFileName' with value '$($config.ExcelFileName)' is missing the date placeholder '{0}'"
            }
            if ($config.ExcelFileName -notlike '*.xlsx') {
                throw "Property 'ExcelFileName' with value '$($config.ExcelFileName)' is not ending with file extension '.xlsx'"
            }

            #region Test SkipFilesModifiedWithinSeconds
            <#
                Optional, so a configuration file written before this existed
                keeps working and gets the default.
            #>
            if ($null -ne $config.SkipFilesModifiedWithinSeconds) {
                try {
                    $skipSeconds = [int]$config.SkipFilesModifiedWithinSeconds
                }
                catch {
                    throw "Property 'SkipFilesModifiedWithinSeconds' needs to be a number, the value '$($config.SkipFilesModifiedWithinSeconds)' is not supported."
                }

                if ($skipSeconds -lt 0) {
                    throw "Property 'SkipFilesModifiedWithinSeconds' with value '$skipSeconds' cannot be negative."
                }
            }
            #endregion

            if (-not $config.Settings.ScriptName) {
                throw "Property 'Settings.ScriptName' not found"
            }

            #region Test Settings.SendMail
            $sendMail = $config.Settings.SendMail

            if (-not $sendMail) {
                throw "Property 'Settings.SendMail' not found"
            }
            if (-not $sendMail.From) {
                throw "Property 'Settings.SendMail.From' not found"
            }
            if ((-not $sendMail.To) -and (-not $sendMail.Bcc)) {
                throw "Property 'Settings.SendMail.To' or 'Settings.SendMail.Bcc' is required"
            }
            if ($sendMail.When -notMatch '^Never$|^Always$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                throw "Property 'Settings.SendMail.When' with value '$($sendMail.When)' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
            }
            if (-not $sendMail.Smtp.ServerName) {
                throw "Property 'Settings.SendMail.Smtp.ServerName' not found"
            }
            if (-not $sendMail.Smtp.Port) {
                throw "Property 'Settings.SendMail.Smtp.Port' not found"
            }
            if (-not $sendMail.AssemblyPath.MailKit) {
                throw "Property 'Settings.SendMail.AssemblyPath.MailKit' not found"
            }
            if (-not $sendMail.AssemblyPath.MimeKit) {
                throw "Property 'Settings.SendMail.AssemblyPath.MimeKit' not found"
            }
            #endregion

            if (-not $config.Settings.SaveLogFiles.Where.Folder) {
                throw "Property 'Settings.SaveLogFiles.Where.Folder' not found"
            }
        }
        catch {
            throw "Configuration file '$ConfigurationJsonFile': $_"
        }
        #endregion

        #region Create the dated log folder
        try {
            $null = New-Item -Path $config.Settings.SaveLogFiles.Where.Folder `
                -ItemType 'Directory' -Force -EA Stop

            $datedLogFolder = Get-DatedLogFolderPathHC `
                -LogFolder $config.Settings.SaveLogFiles.Where.Folder `
                -ScriptStartTime $scriptStartTime `
                -JsonFileName $config.Settings.ScriptName
        }
        catch {
            throw "Failed creating the log folder '$($config.Settings.SaveLogFiles.Where.Folder)': $_"
        }
        #endregion

        $Type = $config.Type
        $XmlFilesFolder = $config.Path.XmlFiles
        $ExcelFilesFolder = $config.Path.ExcelFiles
        $ExcelFileName = $config.ExcelFileName

        $skipFilesModifiedWithinSeconds = if (
            $null -ne $config.SkipFilesModifiedWithinSeconds
        ) {
            [int]$config.SkipFilesModifiedWithinSeconds
        }
        else { 5 }

        #region The values both stages hand to their script block
        <#
            A flat object holding nothing but strings: the two script paths, the
            module path and the type. Flat on purpose, because a script block
            running in another runspace should not have to reach through an
            object to find what it needs, and because everything it does need is
            then visible in one place instead of being collected from whatever
            happens to be in scope.
        #>
        $runContext = [PSCustomObject]@{
            Type            = $Type
            ModulePath      = $modulePathItem
            GetXmlFileDates = $scriptPathItem.GetXmlFileDates
            ExportToExcel   = $scriptPathItem.ExportToExcel
        }
        #endregion

        #region Get the XML files
        Write-Verbose "Get xml files in folder '$XmlFilesFolder'"

        <#
            A file that is still being copied into the folder is already there
            to be found, but only part of it is on disk, so reading it fails on
            an XML document that has no end. That is not a broken file and not
            something to report: it is a run that looked too early. The file is
            left where it is and converts on the next run, whole.

            The moment it was last written to is what decides. Anything touched
            more recently than the configured number of seconds is passed over
            without a word, which is why nothing counts it as an error or names
            it in the mail. Reading the file to find out would mean opening the
            very file that may still be busy.
        #>
        $writtenBefore = (Get-Date).AddSeconds(-$skipFilesModifiedWithinSeconds)

        $allXmlFiles = @(
            Get-ChildItem $XmlFilesFolder -Filter '*.xml' -File | Sort-Object CreationTime
        )

        $xmlFiles = @(
            $allXmlFiles.where({ $_.LastWriteTime -lt $writtenBefore })
        )

        if ($allXmlFiles.Count -ne $xmlFiles.Count) {
            Write-Verbose "Skipped $($allXmlFiles.Count - $xmlFiles.Count) xml files still being written to, they are converted on the next run"
        }

        Write-Verbose "Found $($xmlFiles.Count) xml files of type '$Type'"
        #endregion

        #region Get the months in each XML file
        <#
            One script block for both ways of running. Invoke-WithOptionalParallelismHC
            runs it as a plain loop when the throttle is 1 and in parallel
            runspaces above that, and hands it everything it needs as arguments.

            The block therefore holds no '$using:' at all. It cannot: the helper
            rebuilds it with [scriptblock]::Create() inside each runspace, which
            leaves it without the scope '$using:' would bind to. Everything from
            here travels through the DTO instead, which is also what makes the
            two ways of running the same code rather than two blocks that have
            to be kept in step by hand.
        #>
        $fileDates = @(
            Invoke-WithOptionalParallelismHC `
                -InputObject $xmlFiles `
                -ThrottleLimit ([int]$config.MaxConcurrentJobs.GetXmlFileDates) `
                -ArgumentList $runContext `
                -ScriptBlock {
                param ($XmlFile, $Context)

                & $Context.GetXmlFileDates -XmlFile $XmlFile `
                    -Type $Context.Type -ModulePath $Context.ModulePath
            }
        )

        Write-Verbose "Retrieved dates from $($fileDates.Count) files"
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

        <#
            One flat object per month, holding everything the export script
            needs and nothing else. This is what the script block below is
            handed, one at a time.
        #>
        $groups = @(
            $filesPerMonth.GetEnumerator() | Sort-Object Name | ForEach-Object {
                [PSCustomObject]@{
                    MonthKey      = $_.Key
                    Files         = @($_.Value)
                    ExcelFilePath = Join-Path $ExcelFilesFolder ($ExcelFileName -f $_.Key)
                }
            }
        )

        $collection.FanOut = @(
            $fileDates.where({ (-not $_.Error) -and ($_.MonthKeys.Count -gt 1) })
        )

        Write-Verbose "Found $($groups.Count) different months in $($xmlFiles.Count) XML files"
        #endregion

        #region Create the Excel files
        <#
            The same shape as the stage above: one block, everything it needs
            passed in, no '$using:'. The block carries no try/catch either. The
            export script reports its own failures per file, so there is nothing
            left here to catch that it would not already have answered for.
        #>
        $collection.Results = @(
            Invoke-WithOptionalParallelismHC `
                -InputObject $groups `
                -ThrottleLimit ([int]$config.MaxConcurrentJobs.CreateExcelFile) `
                -ArgumentList $runContext `
                -ScriptBlock {
                param ($Group, $Context)

                & $Context.ExportToExcel -XmlFiles $Group.Files `
                    -ExcelFilePath $Group.ExcelFilePath -Type $Context.Type `
                    -MonthKey $Group.MonthKey -ModulePath $Context.ModulePath
            }
        )
        #endregion

        #region Archive the XML files
        <#
            Archiving is done here and not in the export script, because one XML
            file can be written to more than one Excel file. The file may only
            be moved once every Excel file it belongs to has been handled.

            A file with an error is never archived. It stays in the folder and
            is reported in the mail, so a manual action can be taken. When the
            problem is solved the next run picks the file up again. The Excel
            files it was already added to are not written twice, because the
            file name is checked before adding.

            'NothingToAdd' counts as a handled month, just like a month that was
            added or that was already in the workbook. A month key is read from
            the date of a record while the rows come from the children of that
            record, so a record without children gives a month with no rows: a
            delivery loaded this month holding no batches, or a batch computer
            created this month holding no sequence parameters. Treating that as
            a missing month left the file in the folder and reported it on every
            run, forever, because nothing about the file would ever change.
        #>
        $archiveFolder = Join-Path $XmlFilesFolder 'Archive'

        <#
            Both this loop and the overview further down need the results of
            one file, and both scanned the complete result list for every file
            to find them. With one entry per file per month that scan grows
            with the number of files, and it was repeated once per file, so the
            work grew with the square of the number of files. Walking the
            results once up front and grouping them by name turns every one of
            those scans into a single lookup.
        #>
        foreach ($result in $collection.Results) {
            $resultFileName = $result.File.Name

            if (-not $resultsByFileName.ContainsKey($resultFileName)) {
                $resultsByFileName[$resultFileName] = [System.Collections.Generic.List[Object]]::new()
            }

            $resultsByFileName[$resultFileName].Add($result)
        }

        foreach ($fileDate in $fileDates) {
            $fileName = $fileDate.File.Name

            $fileResults = @($resultsByFileName[$fileName])

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
                    { $_.AddedToSheet -or $_.AlreadyInSheet -or $_.NothingToAdd }
                ).Count -ne $fileDate.MonthKeys.Count
            ) {
                $reason = 'Not all months of this XML file were written to an Excel file'
            }

            if ($reason) {
                $collection.NotArchived.Add(
                    [PSCustomObject]@{
                        File   = $fileDate.File
                        Reason = $reason
                    }
                )

                $archiveOutcome[$fileName] = @{ Archived = $false; Reason = $reason }

                Write-Warning "File '$fileName' not archived: $reason"

                Continue
            }

            #region Move the file to the archive folder
            $moveResult = Move-ToArchiveFolderHC -File $fileDate.File `
                -ArchiveFolder $archiveFolder -ScriptStartTime $scriptStartTime

            $archiveOutcome[$fileName] = $moveResult

            if ($moveResult.Archived) {
                $collection.Archived.Add($fileDate.File)

                if ($moveResult.IsDuplicate) {
                    $collection.Duplicates.Add(
                        [PSCustomObject]@{
                            File       = $fileDate.File
                            ArchivedAs = $moveResult.ArchivedAs
                        }
                    )

                    Write-Verbose "File '$fileName' archived as duplicate '$($moveResult.ArchivedAs)'"
                }
            }
            else {
                $collection.NotArchived.Add(
                    [PSCustomObject]@{
                        File   = $fileDate.File
                        Reason = $moveResult.Reason
                    }
                )

                Write-Warning "File '$fileName' not archived: $($moveResult.Reason)"
            }
            #endregion
        }
        #endregion

        $collection.Added = @($collection.Results.where({ $_.AddedToSheet }))
        $collection.AlreadyInSheet = @($collection.Results.where({ $_.AlreadyInSheet }))

        Write-Verbose 'All done'
    }
    catch {
        Add-ErrorHC -Type 'FatalError' -Name 'Conversion failure' `
            -Message "$_" -Category 'Convert' -SystemErrors $SystemErrors

        Write-Warning $_

        if ($Error.Count) { $Error.RemoveAt(0) }
    }

    #region Collect the date and export failures as system errors
    foreach ($fileDate in $fileDates.where({ $_.Error })) {
        Add-ErrorHC -Type 'Warning' -Name 'Date failure' `
            -Message "File '$($fileDate.File.Name)': $($fileDate.Error)" `
            -Category 'Date' -SystemErrors $SystemErrors
    }

    foreach ($result in $collection.Results.where({ $_.Error })) {
        Add-ErrorHC -Type 'Warning' -Name 'Export failure' `
            -Message "File '$($result.File.Name)' month '$($result.MonthKey)': $($result.Error)" `
            -Category 'Export' -SystemErrors $SystemErrors
    }

    <#
        A file that stays behind always needs a manual action, so it counts as
        an error. Without this the mail could report zero errors while files
        were piling up in the input folder.
    #>
    foreach ($item in $collection.NotArchived) {
        Add-ErrorHC -Type 'Warning' -Name 'Not archived' `
            -Message "File '$($item.File.Name)': $($item.Reason)" `
            -Category 'Archive' -SystemErrors $SystemErrors
    }
    #endregion

    #region Count
    $count = @{
        TotalXmlFiles  = $xmlFiles.Count
        TotalExcelFile = $groups.Count
        <#
            Counted per XML file, not per result. A file that holds records of
            two months produces two results, but it is still one file, so
            'Added' must never be higher than 'TotalXmlFiles'. The quantity per
            Excel file is shown in the mail, where the total per workbook does
            add up to more than the quantity of files.
        #>
        Added          = @($collection.Added.File.Name | Sort-Object -Unique).Count
        AlreadyInSheet = @($collection.AlreadyInSheet.File.Name | Sort-Object -Unique).Count
        FanOut         = $collection.FanOut.Count
        Archived       = $collection.Archived.Count
        Duplicates     = $collection.Duplicates.Count
        NotArchived    = $collection.NotArchived.Count
        Errors         = $SystemErrors.Value.Count
    }
    #endregion

    #region Build the overview
    <#
        One row per XML file per month, plus a row for a file that never got as
        far as an Excel file. The Error column holds whatever went wrong for
        that row: reading the date, writing to Excel, or archiving. That way a
        problem is always on the line of the file it belongs to, instead of in
        a separate sheet that has to be matched up by hand.
    #>
    $overview = [System.Collections.Generic.List[Object]]::new()

    foreach ($fileDate in $fileDates) {
        $fileName = $fileDate.File.Name

        $archive = $archiveOutcome[$fileName]

        <#
            A real boolean, so Excel shows TRUE or FALSE and the column can be
            filtered like the other yes or no columns. The name the file
            received is kept apart in 'ArchivedAs', which only differs from the
            file name when the file was archived as a duplicate.
        #>
        $isArchived = [Boolean]($archive -and $archive.Archived)

        $archivedAs = if ($isArchived) { $archive.ArchivedAs }

        $archiveError = if ($archive -and (-not $archive.Archived)) { $archive.Reason }

        $fileResults = @($resultsByFileName[$fileName])

        if (-not $fileResults) {
            #region The file never reached an Excel file
            $overview.Add(
                [PSCustomObject]@{
                    DateTime       = $scriptStartTime
                    File           = $fileName
                    MonthKey       = $null
                    ExcelFile      = $null
                    AlreadyInSheet = $false
                    AddedToSheet   = $false
                    NothingToAdd   = $false
                    RowsAdded      = 0
                    Archived       = $isArchived
                    ArchivedAs     = $archivedAs
                    Error          = @(
                        $(if ($fileDate.Error) { "No date could be read from the XML file: $($fileDate.Error)" })
                        $archiveError
                    ).where({ $_ }) -join ' | '
                }
            )
            #endregion

            Continue
        }

        foreach ($result in $fileResults) {
            $overview.Add(
                [PSCustomObject]@{
                    DateTime       = $result.DateTime
                    File           = $fileName
                    MonthKey       = $result.MonthKey
                    ExcelFile      = $(if ($result.ExcelFilePath) { Split-Path $result.ExcelFilePath -Leaf })
                    AlreadyInSheet = $result.AlreadyInSheet
                    AddedToSheet   = $result.AddedToSheet
                    <#
                        A real boolean here too, so a row that added nothing can
                        be told apart from a row that failed: the file was read
                        correctly and holds no records for this month.
                    #>
                    NothingToAdd   = [Boolean]$result.NothingToAdd
                    RowsAdded      = $result.RowsAdded
                    Archived       = $isArchived
                    ArchivedAs     = $archivedAs
                    Error          = @(
                        $result.Error
                        $archiveError
                    ).where({ $_ }) -join ' | '
                }
            )
        }
    }
    #endregion

    #region Write the log files
    if ($datedLogFolder) {
        try {
            if ($overview) {
                $logFileAttachments += Out-LogFileHC `
                    -DataToExport $overview `
                    -PartialPath (Join-Path $datedLogFolder 'Overview') `
                    -FileExtensions '.xlsx'
            }

            if ($SystemErrors.Value.Count) {
                $logFileAttachments += Out-LogFileHC `
                    -DataToExport ($SystemErrors.Value | Select-Object *) `
                    -PartialPath (Join-Path $datedLogFolder 'Errors') `
                    -FileExtensions '.json'
            }
        }
        catch {
            Write-Warning "Failed writing the log files: $_"
            if ($Error.Count) { $Error.RemoveAt(0) }
        }
    }
    #endregion

    #region Build and send the summary mail
    if ($config -and $config.Settings.SendMail) {
        try {
            $mail = New-SummaryMailHC -Count $count -Collection $collection `
                -Type $config.Type -Path @{
                XmlFiles   = $config.Path.XmlFiles
                ExcelFiles = $config.Path.ExcelFiles
            }

            #region Check if mail needs to be sent
            $when = $config.Settings.SendMail.When

            $sendMailToUser = (
                ($when -eq 'Always') -or
                (($when -eq 'OnlyOnError') -and $count.Errors) -or
                (
                    ($when -eq 'OnlyOnErrorOrAction') -and
                    ($count.Errors -or $count.Added -or $count.NotArchived)
                )
            )
            #endregion

            #region Save the mail body next to the log files
            if ($datedLogFolder) {
                $savedMailBody = Save-MailBodyToLogHC `
                    -MailParams @{ Subject = $mail.Subject; Body = $mail.Body } `
                    -LogFolder $datedLogFolder

                Write-Verbose "Saved the mail body to '$savedMailBody'"
            }
            #endregion

            if ($sendMailToUser) {
                Write-Verbose 'Send e-mail to the user'

                #region SMTP credential
                $credential = $null

                $smtpUserName = Get-StringValueHC $config.Settings.SendMail.Smtp.UserName
                $smtpPassword = Get-StringValueHC $config.Settings.SendMail.Smtp.Password

                if ($smtpUserName -and $smtpPassword) {
                    $credential = [PSCredential]::new(
                        $smtpUserName,
                        (ConvertTo-SecureString $smtpPassword -AsPlainText -Force)
                    )
                }
                #endregion

                $mailParams = @{
                    MailKitAssemblyPath = Get-StringValueHC $config.Settings.SendMail.AssemblyPath.MailKit
                    MimeKitAssemblyPath = Get-StringValueHC $config.Settings.SendMail.AssemblyPath.MimeKit
                    SmtpServerName      = Get-StringValueHC $config.Settings.SendMail.Smtp.ServerName
                    SmtpPort            = [int](Get-StringValueHC $config.Settings.SendMail.Smtp.Port)
                    SmtpConnectionType  = Get-StringOrDefaultHC (
                        Get-StringValueHC $config.Settings.SendMail.Smtp.ConnectionType
                    ) 'None'
                    From                = Get-StringValueHC $config.Settings.SendMail.From
                    FromDisplayName     = Get-StringValueHC $config.Settings.SendMail.FromDisplayName
                    To                  = Get-MailRecipientListHC -SendMailSettings $config.Settings.SendMail
                    Bcc                 = @($config.Settings.SendMail.Bcc)
                    Subject             = $mail.Subject
                    Body                = $mail.Body
                    Priority            = $(if ($count.Errors) { 'High' } else { 'Normal' })
                    Attachments         = $logFileAttachments
                    Credential          = $credential
                }

                <#
                    Assign first and splat the variable. Splatting the call
                    directly as '@(Remove-BlankValueHC ...)' would splat an
                    ARRAY, which binds the values positionally instead of by
                    name.
                #>
                $cleanMailParams = Remove-BlankValueHC -Hashtable $mailParams

                Send-MailKitMessageHC @cleanMailParams
            }
            else {
                Write-Verbose 'Send no e-mail to the user'
            }
        }
        catch {
            Add-ErrorHC -Type 'Warning' -Name 'Mail failure' `
                -Message "Failed sending the summary mail: $_" `
                -Category 'Mail' -SystemErrors $SystemErrors

            Write-Warning "Failed sending the summary mail: $_"

            if ($Error.Count) { $Error.RemoveAt(0) }
        }
    }
    #endregion

    #region Write to the Windows event log
    if ($config -and $config.Settings.SaveInEventLog) {
        try {
            $eventLogData = [System.Collections.Generic.List[Object]]::new()

            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script started'
                    DateTime  = $scriptStartTime
                    EntryType = 'Information'
                    EventID   = '100'
                }
            )

            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = "$($count.Added) file(s) added, $($count.Archived) archived, $($count.NotArchived) not archived, $($count.Errors) error(s)"
                    DateTime  = Get-Date
                    EntryType = $(if ($count.Errors) { 'Warning' } else { 'Information' })
                    EventID   = '4'
                }
            )

            Write-EventLogSafeHC `
                -EventLogData $eventLogData `
                -ScriptName $config.Settings.ScriptName `
                -Settings $config.Settings `
                -SystemErrors $SystemErrors
        }
        catch {
            Write-Warning "Failed writing to the event log: $_"
            if ($Error.Count) { $Error.RemoveAt(0) }
        }
    }
    #endregion

    #region Remove the old log files
    if ($config -and $config.Settings.SaveLogFiles.DeleteLogsAfterDays) {
        Remove-OldLogsHC `
            -LogFolder $config.Settings.SaveLogFiles.Where.Folder `
            -RetentionDays ([int]$config.Settings.SaveLogFiles.DeleteLogsAfterDays) `
            -SystemErrors $SystemErrors
    }
    #endregion
}