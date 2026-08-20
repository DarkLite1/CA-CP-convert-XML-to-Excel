#Requires -Version 7

<#
    .SYNOPSIS
        Build the subject and HTML body of the summary e-mail.

    .DESCRIPTION
        The Permission-matrix HTML builders are specific to permission matrices,
        so the body for this script is built here. Plain table based HTML with
        inline styles, so it renders the same in Outlook and in a browser.
#>

function New-SummaryMailHC {
    <#
        .SYNOPSIS
            Build the subject line and HTML body for the summary e-mail.

        .DESCRIPTION
            Summarizes one run: how many XML files were found, how many were
            added to which Excel file, how many spanned more than one month, how
            many were already present, and how many were archived. Files that
            could not be archived are highlighted, because they need a manual
            action before the next run.

        .PARAMETER Count
            The hashtable of counters built by Invoke-ConvertXmlToExcel.

        .PARAMETER Collection
            The hashtable of result collections built by Invoke-ConvertXmlToExcel.

        .PARAMETER Type
            'Batch', 'Alarm' or 'Sequence', named in the intro line.

        .PARAMETER Path
            Hashtable with the 'XmlFiles' and 'ExcelFiles' folders, linked in
            the body so they can be opened from the mail.

        .OUTPUTS
            A hashtable with the keys 'Subject' and 'Body'.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [HashTable]$Count,
        [Parameter(Mandatory)]
        [HashTable]$Collection,
        [Parameter(Mandatory)]
        [String]$Type,
        [Parameter(Mandatory)]
        [HashTable]$Path
    )

    #region Subject
    $subjectParts = @(
        '{0} file{1} added' -f $Count.Added, $(if ($Count.Added -ne 1) { 's' })
    )

    if ($Count.NotArchived) {
        $subjectParts += '{0} file{1} not archived' -f
        $Count.NotArchived, $(if ($Count.NotArchived -ne 1) { 's' })
    }

    if ($Count.Errors) {
        $subjectParts += '{0} error{1}' -f
        $Count.Errors, $(if ($Count.Errors -ne 1) { 's' })
    }

    $subject = $subjectParts -join ', '
    #endregion

    #region Rows
    $rows = [System.Text.StringBuilder]::new()

    $null = $rows.Append(
        ('<tr><td>{0}</td><td>Total XML file{1}</td></tr>' -f
        $Count.TotalXmlFiles, $(if ($Count.TotalXmlFiles -ne 1) { 's' }))
    )

    if ($Count.Added) {
        $perFile = (
            $Collection.Added | Group-Object 'ExcelFilePath' | ForEach-Object {
                '<br>- {0}: {1}' -f $_.Count, (Split-Path $_.Name -Leaf)
            }
        ) -join ''

        $null = $rows.Append(
            ('<tr><td>{0}</td><td>XML file{1} added to Excel:{2}</td></tr>' -f
            $Count.Added, $(if ($Count.Added -ne 1) { 's' }), $perFile)
        )
    }

    if ($Count.FanOut) {
        $null = $rows.Append(
            ('<tr><td>{0}</td><td>XML file{1} with records in more than one month, written to more than one Excel file</td></tr>' -f
            $Count.FanOut, $(if ($Count.FanOut -ne 1) { 's' }))
        )
    }

    if ($Count.AlreadyInSheet) {
        $null = $rows.Append(
            ('<tr><td>{0}</td><td>XML file{1} already in Excel</td></tr>' -f
            $Count.AlreadyInSheet, $(if ($Count.AlreadyInSheet -ne 1) { 's' }))
        )
    }

    if ($Count.Archived) {
        $null = $rows.Append(
            ('<tr><td>{0}</td><td>XML file{1} moved to the archive folder</td></tr>' -f
            $Count.Archived, $(if ($Count.Archived -ne 1) { 's' }))
        )
    }

    if ($Count.NotArchived) {
        $detail = (
            $Collection.NotArchived | ForEach-Object {
                '<br>- {0}: {1}' -f $_.File.Name, $_.Reason
            }
        ) -join ''

        $null = $rows.Append(
            ('<tr style="background-color:#fee2e2"><td>{0}</td><td><b>XML file{1} NOT archived, manual action required:</b>{2}</td></tr>' -f
            $Count.NotArchived, $(if ($Count.NotArchived -ne 1) { 's' }), $detail)
        )
    }

    if ($Count.Errors) {
        $null = $rows.Append(
            ('<tr style="background-color:#fee2e2"><td>{0}</td><td>Error{1}, see the attached log file for the details</td></tr>' -f
            $Count.Errors, $(if ($Count.Errors -ne 1) { 's' }))
        )
    }
    #endregion

    $body = @"
<p>Summary of the XML to Excel conversion (type '$Type'):</p>
<table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;">
    <tr>
        <th style="text-align:left;padding:4px 8px;">Quantity</th>
        <th style="text-align:left;padding:4px 8px;">Description</th>
    </tr>
    $($rows.ToString())
</table>
<p>Folder locations:</p>
<table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;">
    <tr>
        <th style="text-align:left;padding:4px 8px;">XML files</th>
        <td style="padding:4px 8px;"><a href="$($Path.XmlFiles)">$($Path.XmlFiles)</a></td>
    </tr>
    <tr>
        <th style="text-align:left;padding:4px 8px;">Excel files</th>
        <td style="padding:4px 8px;"><a href="$($Path.ExcelFiles)">$($Path.ExcelFiles)</a></td>
    </tr>
</table>
"@

    @{
        Subject = $subject
        Body    = $body
    }
}
