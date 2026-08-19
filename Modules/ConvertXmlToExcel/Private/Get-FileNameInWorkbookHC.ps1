#Requires -Version 7

function Get-FileNameInWorkbookHC {
    <#
        .SYNOPSIS
            The names of the XML files already present in the Excel file.

        .DESCRIPTION
            Used to make sure the same XML file is never added twice to the
            same Excel file.

            One XML file can appear in more than one Excel file, because it can
            hold records for more than one month. Within a single Excel file
            though, an XML file is either fully added or not added at all, so
            checking the file name is enough and there is no need to compare
            each row.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [HashTable]$Workbook
    )

    $fileNames = @{}

    $worksheet = $Workbook.Sheet[$Workbook.PlantKey].Worksheet

    $lastRow = $worksheet.Dimension.End.Row

    if ($lastRow -lt 3) { return $fileNames }

    $worksheet.Cells["A3:A$lastRow"].Value | Sort-Object -Unique |
    ForEach-Object {
        if ($_) { $fileNames[$_] = $true }
    }

    $fileNames
}
