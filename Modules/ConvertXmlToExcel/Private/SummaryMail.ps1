#Requires -Version 7

<#
    .SYNOPSIS
        Build the subject and HTML body of the summary e-mail.

    .DESCRIPTION
        The HTML is deliberately old fashioned: tables with the 'border',
        'cellpadding' and 'cellspacing' attributes, inline styles on every cell
        and no CSS classes, no stylesheet and no layout that depends on float,
        flex or grid. The classic Outlook client renders HTML with the Word
        engine, which ignores most of that. Everything below is what the Word
        engine does support.
#>

function New-SummaryMailHC {
    <#
        .SYNOPSIS
            Build the subject line and HTML body for the summary e-mail.

        .DESCRIPTION
            The body opens with the Excel files that were written, most recent
            month first, so the newest work is visible without scrolling. The
            totals follow at the bottom, because they are a conclusion rather
            than something to read past on the way to the detail.

            Files that could not be archived are listed with their reason,
            because they need a manual action before the next run.

        .PARAMETER Count
            The hashtable of counters built by Invoke-ConvertXmlToExcel.

        .PARAMETER Collection
            The hashtable of result collections built by Invoke-ConvertXmlToExcel.

        .PARAMETER Type
            'Batch', 'Alarm' or 'Sequence', named in the intro line.

        .PARAMETER Path
            Hashtable with the 'XmlFiles' and 'ExcelFiles' folders, linked at
            the bottom so they can be opened from the mail.

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

    #region Styles reused by every table
    $tableStyle = 'border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:10pt;'
    $headerStyle = 'background-color:#f2f2f2;font-weight:bold;text-align:left;'
    $cellStyle = 'font-family:Segoe UI,Arial,sans-serif;font-size:10pt;'
    $numberStyle = "$cellStyle text-align:right;"
    #endregion

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

    $body = [System.Text.StringBuilder]::new()

    $null = $body.Append(
        ("<p style='$cellStyle'>Summary of the XML to Excel conversion (type '{0}'):</p>" -f $Type)
    )

    #region Excel files, most recent month first
    if ($Collection.Added) {
        <#
            Grouped per Excel file, so one row per workbook instead of one row
            per XML file. The month key starts with 'yyyyMM', so sorting the key
            descending puts the most recent month on top.
        #>
        $perExcelFile = $Collection.Added | Group-Object 'ExcelFilePath' |
        ForEach-Object {
            [PSCustomObject]@{
                MonthKey  = @($_.Group.MonthKey)[0]
                ExcelFile = Split-Path $_.Name -Leaf
                XmlFiles  = $_.Count
                Rows      = ($_.Group | Measure-Object -Property 'RowsAdded' -Sum).Sum
            }
        } | Sort-Object 'MonthKey' -Descending

        $null = $body.Append(
            "<table border='1' cellpadding='5' cellspacing='0' style='$tableStyle'>" +
            "<tr>" +
            "<th style='$cellStyle$headerStyle'>Month</th>" +
            "<th style='$cellStyle$headerStyle'>Excel file</th>" +
            "<th style='$cellStyle$headerStyle'>XML files</th>" +
            "<th style='$cellStyle$headerStyle'>Rows</th>" +
            "</tr>"
        )

        foreach ($row in $perExcelFile) {
            $null = $body.Append(
                ("<tr><td style='$cellStyle'>{0}</td><td style='$cellStyle'>{1}</td><td style='$numberStyle'>{2}</td><td style='$numberStyle'>{3}</td></tr>" -f
                $row.MonthKey, $row.ExcelFile, $row.XmlFiles, $row.Rows)
            )
        }

        $null = $body.Append('</table>')
    }
    else {
        $null = $body.Append(
            "<p style='$cellStyle'>No XML file was added to an Excel file.</p>"
        )
    }
    #endregion

    #region Files that need a manual action
    if ($Collection.NotArchived) {
        $null = $body.Append(
            "<p style='$cellStyle'><b>XML files NOT archived, manual action required:</b></p>" +
            "<table border='1' cellpadding='5' cellspacing='0' style='$tableStyle'>" +
            "<tr>" +
            "<th style='$cellStyle$headerStyle'>XML file</th>" +
            "<th style='$cellStyle$headerStyle'>Reason</th>" +
            "</tr>"
        )

        foreach ($item in $Collection.NotArchived) {
            $null = $body.Append(
                ("<tr><td style='$cellStyle' bgcolor='#fee2e2'>{0}</td><td style='$cellStyle' bgcolor='#fee2e2'>{1}</td></tr>" -f
                $item.File.Name, $item.Reason)
            )
        }

        $null = $body.Append('</table>')
    }
    #endregion

    #region Totals, after the detail
    $totals = [System.Collections.Generic.List[Object]]::new()

    $totals.Add(@{ Name = 'Total XML files'; Value = $Count.TotalXmlFiles })
    $totals.Add(@{ Name = 'Added to Excel'; Value = $Count.Added })

    if ($Count.AlreadyInSheet) {
        $totals.Add(@{ Name = 'Already in Excel'; Value = $Count.AlreadyInSheet })
    }
    if ($Count.FanOut) {
        $totals.Add(@{
                Name  = 'Spread over more than one month'
                Value = $Count.FanOut
            })
    }

    $totals.Add(@{ Name = 'Moved to the archive folder'; Value = $Count.Archived })

    if ($Count.NotArchived) {
        $totals.Add(@{ Name = 'NOT archived'; Value = $Count.NotArchived })
    }
    if ($Count.Errors) {
        $totals.Add(@{ Name = 'Errors'; Value = $Count.Errors })
    }

    $null = $body.Append(
        "<p style='$cellStyle'><b>Summary</b></p>" +
        "<table border='1' cellpadding='5' cellspacing='0' style='$tableStyle'>"
    )

    foreach ($total in $totals) {
        $highlight = if ($total.Name -match 'NOT archived|Errors') {
            " bgcolor='#fee2e2'"
        }

        $null = $body.Append(
            ("<tr><td style='$cellStyle$headerStyle'$highlight>{0}</td><td style='$numberStyle'$highlight>{1}</td></tr>" -f
            $total.Name, $total.Value)
        )
    }

    $null = $body.Append('</table>')
    #endregion

    #region Folders
    $null = $body.Append(
        "<p style='$cellStyle'>Folders:</p>" +
        "<table border='1' cellpadding='5' cellspacing='0' style='$tableStyle'>" +
        ("<tr><td style='$cellStyle$headerStyle'>XML files</td><td style='$cellStyle'><a href='{0}'>{0}</a></td></tr>" -f $Path.XmlFiles) +
        ("<tr><td style='$cellStyle$headerStyle'>Excel files</td><td style='$cellStyle'><a href='{0}'>{0}</a></td></tr>" -f $Path.ExcelFiles) +
        '</table>'
    )
    #endregion

    @{
        Subject = $subject
        Body    = $body.ToString()
    }
}