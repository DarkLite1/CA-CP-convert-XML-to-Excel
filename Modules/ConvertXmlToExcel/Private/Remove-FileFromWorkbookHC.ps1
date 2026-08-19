#Requires -Version 7

function Remove-FileFromWorkbookHC {
    <#
        .SYNOPSIS
            Remove every row of one XML file from every worksheet.

        .DESCRIPTION
            Safety net for the case where writing the rows of a file fails
            halfway. A file is either fully in the Excel file or not at all.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [HashTable]$Workbook,
        [Parameter(Mandatory)]
        [String]$FileName
    )

    foreach ($key in $Workbook.Sheet.Keys) {
        $sheet = $Workbook.Sheet[$key]

        $lastRow = $sheet.Worksheet.Dimension.End.Row

        if ($lastRow -lt 3) { Continue }

        $rowsToRemove = $sheet.Worksheet.Cells["A3:A$lastRow"].Where(
            { $_.Value -eq $FileName }
        ).Start.Row | Sort-Object -Descending -Unique

        foreach ($rowNumber in $rowsToRemove) {
            Write-Verbose "Remove row '$rowNumber' in sheet '$($sheet.Worksheet.Name)' for file '$FileName'"
            $sheet.Worksheet.DeleteRow($rowNumber)
        }

        $sheet.RowNumber = $sheet.Worksheet.Dimension.End.Row + 1
    }
}
