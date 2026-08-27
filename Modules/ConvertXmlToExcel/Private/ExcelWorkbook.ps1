#Requires -Version 7
#Requires -Modules ImportExcel

<#
    .SYNOPSIS
        Open, fill and format the Excel files.

    .DESCRIPTION
        These functions are type independent. They take the worksheet layout
        from Get-WorksheetDefinitionHC and the rows from Get-XmlRowHC.
#>

<#
    The maximum number of rows a worksheet can hold in the .xlsx format. Used
    by Test-RowLimitHC to refuse an XML file that would not fit completely.
#>
$script:ExcelMaxRowNumber = 1048576

<#
    The number format of every date column. Applied once per worksheet at
    column level by Set-DateColumnFormatHC, not per cell.
#>
$script:ExcelDateFormat = 'dd/mm/yyyy hh:mm:ss'

<#
    How many data rows the column auto sizing looks at. Auto sizing measures
    the rendered text of every cell in its range, so letting it loose on a
    worksheet that grows towards a million rows makes every run slower than the
    one before it, while re-measuring data that has not changed since it was
    written. A sample of the first rows gives the same widths in practice at a
    cost that no longer depends on the size of the workbook.
#>
$script:AutoFitSampleRowCount = 200

function Open-ExcelWorkbookHC {
    <#
        .SYNOPSIS
            Open an existing Excel file or create a new one with all worksheets
            and headers in place.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$Path,
        [Parameter(Mandatory)]
        [ValidateSet('Batch', 'Alarm', 'Sequence')]
        [String]$Type
    )

    $definitions = @(Get-WorksheetDefinitionHC -Type $Type)

    $workbook = @{
        Path        = $Path
        Created     = $false
        PlantKey    = $definitions[0].Key
        Definitions = $definitions
        Sheet       = @{}
    }

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        #region Open all worksheets
        try {
            $workbook.Package = Open-ExcelPackage -Path $Path -ErrorAction Stop

            foreach ($definition in $definitions) {
                $worksheet = $workbook.Package.Workbook.Worksheets[$definition.Name]

                if (-not $worksheet) {
                    throw "Worksheet '$($definition.Name)' not found"
                }

                <#
                    A worksheet that holds nothing at all has no 'Dimension',
                    and one that lost its header rows reports a last row below
                    the two rows every worksheet of this script starts with.

                    Both are refused instead of being worked around. Reading
                    'Dimension.End.Row + 1' on an empty worksheet gives row 1,
                    so the next write would land on top of the header rows and
                    the damage would only show up in the Excel file, long after
                    the run reported success. A workbook in that state was
                    emptied or edited by hand, which is a manual action, so it
                    is named here as one.
                #>
                if (
                    (-not $worksheet.Dimension) -or
                    ($worksheet.Dimension.End.Row -lt 2)
                ) {
                    throw "Worksheet '$($definition.Name)' is empty, its header rows are missing"
                }

                $workbook.Sheet[$definition.Key] = @{
                    Definition = $definition
                    Worksheet  = $worksheet
                    RowNumber  = $worksheet.Dimension.End.Row + 1
                }

                <#
                    Also applied to a workbook created by an earlier version,
                    which only had the format on the cells written at the time.
                    Setting it on the column costs the same on an empty
                    worksheet as on a full one, so it is done on every open.
                #>
                Set-DateColumnFormatHC -Worksheet $worksheet -Definition $definition
            }
        }
        catch {
            throw "Failed opening Excel file '$Path': $_"
        }
        #endregion
    }
    else {
        #region Create all worksheets
        $workbook.Created = $true
        $workbook.Package = Open-ExcelPackage -Path $Path -Create

        foreach ($definition in $definitions) {
            $params = @{
                ExcelPackage  = $workbook.Package
                WorksheetName = $definition.Name
            }
            $worksheet = Add-Worksheet @params

            $workbook.Sheet[$definition.Key] = @{
                Definition = $definition
                Worksheet  = $worksheet
                RowNumber  = 3
            }

            #region Merge cells
            foreach ($range in $definition.MergeColumns) {
                $worksheet.Cells[$range].Merge = $true
            }
            #endregion

            #region Add first row header and format
            foreach ($item in $definition.FirstHeaderRow.GetEnumerator()) {
                $worksheet.Cells[$item.Key].Value = $item.Value

                $params = @{
                    Worksheet = $worksheet
                    Cell      = $item.Key
                    Type      = $item.Value
                }
                Format-WorksheetHeaderHC @params
            }
            #endregion

            #region Add table headers
            $i = 0

            $params = @{
                RowNumber       = 2
                RequiredColumns = $definition.TableHeaderRow.Count
            }

            foreach ($address in (Get-ColumnCellHC @params)) {
                $worksheet.Cells[$address].Value = $definition.TableHeaderRow[$i]
                $i++
            }
            #endregion

            Set-DateColumnFormatHC -Worksheet $worksheet -Definition $definition
        }
        #endregion
    }

    $workbook
}

function Set-DateColumnFormatHC {
    <#
        .SYNOPSIS
            Give the date columns of a worksheet their number format.

        .DESCRIPTION
            The format is set on the COLUMN, not on a range of cells, and only
            when the workbook is opened.

            Formatting the range 'H3:H<lastRow>' after every write, as was done
            before, re-styles every row that was already in the workbook, so the
            cost grows with the workbook and the same historical rows are styled
            again on every run. A column carries a single style entry, so this
            costs the same whether the worksheet holds three rows or a million,
            and rows written later inherit it automatically.

            Rows 1 and 2 hold the headers and are inside the column too, but a
            date format has no effect on text, so they are unaffected.

        .PARAMETER Worksheet
            The EPPlus worksheet.

        .PARAMETER Definition
            The worksheet definition from Get-WorksheetDefinitionHC, holding
            'DateColumns' as column letters.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Worksheet,
        [Parameter(Mandatory)]
        $Definition
    )

    if (-not $Definition.DateColumns) { return }

    foreach ($column in $Definition.DateColumns) {
        try {
            <#
                The definition holds column letters and the column indexer
                needs a number. Letting EPPlus resolve the address keeps the
                letters in the definition, where they are readable, and keeps
                a letter to number conversion out of this module.
            #>
            $columnNumber = $Worksheet.Cells["${column}1"].Start.Column

            $Worksheet.Column($columnNumber).Style.NumberFormat.Format = $script:ExcelDateFormat
        }
        catch {
            throw "Failed formatting date column '$column' of worksheet '$($Worksheet.Name)': $_"
        }
    }
}

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

    <#
        A worksheet with no cells at all has no 'Dimension', so it is read
        before the last row and not through it. Nothing was ever added to such
        a worksheet, which is exactly what an empty result says.
    #>
    $dimension = $worksheet.Dimension

    if (-not $dimension) { return $fileNames }

    $lastRow = $dimension.End.Row

    if ($lastRow -lt 3) { return $fileNames }

    $worksheet.Cells["A3:A$lastRow"].Value | Sort-Object -Unique |
    ForEach-Object {
        if ($_) { $fileNames[$_] = $true }
    }

    $fileNames
}

function Test-RowLimitHC {
    <#
        .SYNOPSIS
            Check up front whether all rows of one XML file still fit.

        .DESCRIPTION
            An XML file is added completely or not at all. Checking before
            writing avoids ending up with half a file in the Excel file when
            the maximum number of rows of a worksheet is reached.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [HashTable]$Workbook,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [Array]$Rows,
        [int]$MaxRowNumber = $script:ExcelMaxRowNumber
    )

    $requiredRows = @{}

    foreach ($row in $Rows) {
        $requiredRows[$row.Key] = 1 + $requiredRows[$row.Key]
    }

    foreach ($item in $requiredRows.GetEnumerator()) {
        $sheet = $Workbook.Sheet[$item.Key]

        <#
            '-ge' and not '-gt': a worksheet cannot hold a row AT the maximum
            row number either, so writing has to stop one row earlier.
        #>
        if (($sheet.RowNumber + $item.Value - 1) -ge $MaxRowNumber) {
            return $false
        }
    }

    $true
}

function Add-RowToWorkbookHC {
    <#
        .SYNOPSIS
            Write the rows of one XML file to the Excel file.

        .DESCRIPTION
            Returns the number of rows written. Throws when the rows do not fit
            anymore, without having written anything.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [HashTable]$Workbook,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [Array]$Rows
    )

    if (-not $Rows.Count) { return 0 }

    if (-not (Test-RowLimitHC -Workbook $Workbook -Rows $Rows)) {
        throw 'XML file could not be added to Excel, max row limit reached'
    }

    <#
        Column letter to column number, built as the letters are met and reused
        for every following row. A worksheet has at most a few dozen columns
        while this call writes every row of an XML file, so after the first row
        every lookup is a hash hit.

        Kept local on purpose. A script scoped cache would have to be declared
        in some file, and this function would then quietly stop working when it
        is loaded on its own, which is exactly what the unit test does.
    #>
    $columnNumberByLetter = @{}

    foreach ($row in $Rows) {
        $sheet = $Workbook.Sheet[$row.Key]

        <#
            Pulled out of the cell loop. Both were resolved again for every
            single cell before, and 'Worksheet' is a hash lookup.
        #>
        $worksheet = $sheet.Worksheet
        $rowNumber = $sheet.RowNumber

        <#
            'SetValue' writes straight to the cell store. The previous
            '$worksheet.Cells[$address].Value = ...' built an address string
            like 'AB1234', handed it back to EPPlus to be parsed into a row and
            a column again, and allocated an ExcelRange object to hold the
            result. That is three throwaway objects per cell, on a worksheet
            with up to 68 columns, for every row of every file.

            A 'foreach' statement replaces '.GetEnumerator().ForEach({ })'
            because the method form invokes its script block once per cell,
            which carries its own per-call cost on this path.
        #>
        foreach ($cell in $row.Cells.GetEnumerator()) {
            $columnLetter = $cell.Key

            $columnNumber = $columnNumberByLetter[$columnLetter]

            if (-not $columnNumber) {
                #region Column letter to column number: A is 1, Z is 26, AA is 27
                $columnNumber = 0

                foreach ($letter in [Char[]]$columnLetter.ToUpperInvariant()) {
                    if ($letter -lt 'A' -or $letter -gt 'Z') {
                        throw "Column '$columnLetter' is not a valid column letter"
                    }

                    $columnNumber = ($columnNumber * 26) + ([int]$letter - 64)
                }
                #endregion

                $columnNumberByLetter[$columnLetter] = $columnNumber
            }

            $worksheet.SetValue($rowNumber, $columnNumber, $cell.Value)
        }

        $sheet.RowNumber++
    }

    $Rows.Count
}

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

        <#
            Removing rows can empty a worksheet completely, and a worksheet
            with no cells has no 'Dimension' at all, so it is tested for before
            a row number is read from it.
        #>
        $dimension = $sheet.Worksheet.Dimension

        if (-not $dimension) {
            $sheet.RowNumber = 3
            Continue
        }

        $lastRow = $dimension.End.Row

        if ($lastRow -lt 3) { Continue }

        $rowsToRemove = $sheet.Worksheet.Cells["A3:A$lastRow"].Where(
            { $_.Value -eq $FileName }
        ).Start.Row | Sort-Object -Descending -Unique

        foreach ($rowNumber in $rowsToRemove) {
            Write-Verbose "Remove row '$rowNumber' in sheet '$($sheet.Worksheet.Name)' for file '$FileName'"
            $sheet.Worksheet.DeleteRow($rowNumber)
        }

        <#
            The header rows survive the deletion, so the dimension is normally
            still there and the next write lands right below the rows that are
            left. The fallback is the first data row, which is where writing
            starts on a worksheet that holds nothing.
        #>
        $dimension = $sheet.Worksheet.Dimension

        $sheet.RowNumber = if ($dimension) { $dimension.End.Row + 1 } else { 3 }
    }
}

function Format-ExcelWorkbookHC {
    <#
        .SYNOPSIS
            Format the date columns, the tables and the frozen rows, then save.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [HashTable]$Workbook
    )

    try {
        foreach ($definition in $Workbook.Definitions) {
            $sheet = $Workbook.Sheet[$definition.Key]

            try {
                $sheetName = $sheet.Worksheet.Name

                Write-Verbose "Format worksheet '$sheetName'"

                <#
                    The date columns are not formatted here anymore. They get
                    their number format once per open, at column level, in
                    Set-DateColumnFormatHC. Rows written in this run inherit it,
                    so there is nothing left to do per run.
                #>

                #region Format tables
                if ($definition.TableName) {
                    try {
                        $tableName = $definition.TableName

                        $range = 'A2:{0}' -f $sheet.Worksheet.Dimension.End.Address

                        if ($table = $sheet.Worksheet.Tables[$tableName]) {
                            $table.TableXml.table.ref = $range

                            <#
                                The filter of a table carries its own range and
                                Excel expects the two to say the same thing.
                                Growing only the table left the filter on the
                                range of the run before, so the rows added last
                                fell outside it and Excel could refuse to open
                                the workbook without repairing it first, which
                                is how a month's file is lost.

                                Tested for rather than assumed: a table without
                                a header row has no filter to grow.
                            #>
                            $autoFilter = $table.TableXml.table.autoFilter

                            if ($autoFilter) {
                                $autoFilter.ref = $range
                            }
                        }
                        else {
                            $table = $sheet.Worksheet.Tables.Add(
                                $sheet.Worksheet.Cells[$range], $tableName
                            )
                        }

                        $table.TableStyle = 'Medium6'

                        <#
                            Auto sizing the columns is cosmetic. When it fails
                            the rows are still correct, so it may not stop the
                            Excel file from being saved. If saving would fail
                            here the XML file would not be archived and would be
                            reported as an error every single day.

                            It is limited to the first rows on purpose. Auto
                            sizing measures the rendered text of every cell in
                            its range, so running it over the whole worksheet
                            made every run slower than the one before it, while
                            re-measuring rows written days ago that have not
                            changed. The widest values sit in the headers and
                            the first rows anyway, so a sample gives the same
                            widths at a cost that no longer grows.
                        #>
                        try {
                            $lastRow = $sheet.Worksheet.Dimension.End.Row
                            $lastColumn = $sheet.Worksheet.Dimension.End.Column

                            $sampleLastRow = [Math]::Min(
                                $lastRow, 2 + $script:AutoFitSampleRowCount
                            )

                            $sheet.Worksheet.Cells[
                                1, 1, $sampleLastRow, $lastColumn
                            ].AutoFitColumns()
                        }
                        catch {
                            Write-Warning "Failed auto sizing the columns of sheet '$sheetName': $_"
                            if ($Error.Count) { $Error.RemoveAt(0) }
                        }
                    }
                    catch {
                        throw "Failed formatting table '$tableName' in range '$range': $_"
                    }
                }
                #endregion

                #region Freeze top rows
                if ($definition.FreezeRow) {
                    try {
                        $freezeRow = $definition.FreezeRow

                        $sheet.Worksheet.View.FreezePanes(($freezeRow + 1), 1)
                    }
                    catch {
                        throw "Failed to freeze row '$freezeRow': $_"
                    }
                }
                #endregion
            }
            catch {
                $M = "Failed formatting sheet '$sheetName' in file '$($Workbook.Path)'. New rows are not added to sheet. Error: $_"
                if ($Error.Count) { $Error.RemoveAt(0) }
                throw $M
            }
        }

        if ($Workbook.Created) {
            Write-Verbose "Create Excel file '$($Workbook.Path)'"
        }
        else {
            Write-Verbose "Update Excel file '$($Workbook.Path)'"
        }

        $Workbook.Package.Save()
    }
    catch {
        $M = "Failed saving Excel file '$($Workbook.Path)': $_"
        if ($Error.Count) { $Error.RemoveAt(0) }
        throw $M
    }
}