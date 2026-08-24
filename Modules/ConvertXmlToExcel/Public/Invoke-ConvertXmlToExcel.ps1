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

    $collection = @{
        Results        = @()
        Added          = @()
        AlreadyInSheet = @()
        FanOut         = @()
        Archived       = @()
        NotArchived    = @()
    }

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

        #region Get the XML files
        Write-Verbose "Get xml files in folder '$XmlFilesFolder'"

        $xmlFiles = @(
            Get-ChildItem $XmlFilesFolder -Filter '*.xml' -File | Sort-Object CreationTime
        )

        Write-Verbose "Found $($xmlFiles.Count) xml files of type '$Type'"
        #endregion

        #region Get the months in each XML file
        $scriptBlock = {
            $file = $_

            #region Declare variables for code running in parallel
            if (-not $MaxConcurrentJobs) {
                $scriptPathItem = $using:scriptPathItem
                $modulePathItem = $using:modulePathItem
                $Type = $using:Type
            }
            #endregion

            & $scriptPathItem.GetXmlFileDates -XmlFile $file -Type $Type -ModulePath $modulePathItem
        }

        $MaxConcurrentJobs = [int]$config.MaxConcurrentJobs.GetXmlFileDates

        $foreachParams = if ($MaxConcurrentJobs -eq 1) {
            @{ Process = $scriptBlock }
        }
        else {
            @{ Parallel = $scriptBlock; ThrottleLimit = $MaxConcurrentJobs }
        }

        $fileDates = @($xmlFiles | ForEach-Object @foreachParams)

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

        $groups = @(
            $filesPerMonth.GetEnumerator() | Sort-Object Name | ForEach-Object {
                @{
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
        $scriptBlock = {
            $group = $_

            #region Declare variables for code running in parallel
            if (-not $MaxConcurrentJobs) {
                $scriptPathItem = $using:scriptPathItem
                $modulePathItem = $using:modulePathItem
                $Type = $using:Type
            }
            #endregion

            & $scriptPathItem.ExportToExcel -XmlFiles $group.Files -ExcelFilePath $group.ExcelFilePath -Type $Type -MonthKey $group.MonthKey -ModulePath $modulePathItem
        }

        $MaxConcurrentJobs = [int]$config.MaxConcurrentJobs.CreateExcelFile

        $foreachParams = if ($MaxConcurrentJobs -eq 1) {
            @{ Process = $scriptBlock }
        }
        else {
            @{ Parallel = $scriptBlock; ThrottleLimit = $MaxConcurrentJobs }
        }

        $collection.Results = @($groups | ForEach-Object @foreachParams)
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
        #>
        $archiveFolder = Join-Path $XmlFilesFolder 'Archive'

        foreach ($fileDate in $fileDates) {
            $fileName = $fileDate.File.Name

            $fileResults = @(
                $collection.Results.where({ $_.File.Name -eq $fileName })
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
                $fileResults.where({ $_.AddedToSheet -or $_.AlreadyInSheet }).Count -ne
                $fileDate.MonthKeys.Count
            ) {
                $reason = 'Not all months of this XML file were written to an Excel file'
            }

            if ($reason) {
                $collection.NotArchived += [PSCustomObject]@{
                    File   = $fileDate.File
                    Reason = $reason
                }

                Write-Warning "File '$fileName' not archived: $reason"

                Continue
            }

            #region Move the file to the archive folder
            try {
                $null = New-Item -Path $archiveFolder -ItemType Directory -EA Ignore

                if (Test-Path -LiteralPath $fileDate.File.FullName -PathType Leaf) {
                    Write-Verbose "Move file '$fileName' to '$archiveFolder'"

                    Move-Item -LiteralPath $fileDate.File.FullName -Destination $archiveFolder -EA Stop
                }

                $collection.Archived += $fileDate.File
            }
            catch {
                $collection.NotArchived += [PSCustomObject]@{
                    File   = $fileDate.File
                    Reason = "Failed moving file to archive folder '$archiveFolder': $_"
                }
                if ($Error.Count) { $Error.RemoveAt(0) }
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
        NotArchived    = $collection.NotArchived.Count
        Errors         = $SystemErrors.Value.Count
    }
    #endregion

    #region Write the log files
    if ($datedLogFolder) {
        try {
            if ($collection.Results) {
                $logFileAttachments += Out-LogFileHC `
                    -DataToExport (
                        $collection.Results | Select-Object -Property DateTime,
                        ExcelFilePath, MonthKey,
                        @{ Name = 'File'; Expression = { $_.File.Name } },
                        AlreadyInSheet, AddedToSheet, RowsAdded, Error
                    ) `
                    -PartialPath (Join-Path $datedLogFolder 'Overview') `
                    -FileExtensions '.xlsx'
            }

            if ($collection.NotArchived) {
                $logFileAttachments += Out-LogFileHC `
                    -DataToExport (
                        $collection.NotArchived | Select-Object -Property @{
                            Name = 'File'; Expression = { $_.File.FullName }
                        }, Reason
                    ) `
                    -PartialPath (Join-Path $datedLogFolder 'NotArchived') `
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