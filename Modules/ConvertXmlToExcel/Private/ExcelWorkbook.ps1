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

                $workbook.Sheet[$definition.Key] = @{
                    Definition = $definition
                    Worksheet  = $worksheet
                    RowNumber  = $worksheet.Dimension.End.Row + 1
                }
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
        }
        #endregion
    }

    $workbook
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

    $lastRow = $worksheet.Dimension.End.Row

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

    foreach ($row in $Rows) {
        $sheet = $Workbook.Sheet[$row.Key]

        $row.Cells.GetEnumerator().ForEach(
            {
                $address = '{0}{1}' -f $_.Key, $sheet.RowNumber

                $sheet.Worksheet.Cells[$address].Value = $_.Value
            }
        )

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

                #region Format dates
                if ($definition.DateColumns) {
                    $lastRow = $sheet.Worksheet.Dimension.End.Row

                    <#
                        Data rows start at row 3. When a worksheet only holds
                        its header rows (for example a batch with no discharging
                        operations) there is nothing to format and the range
                        'H3:H2' would be invalid, so it is skipped.
                    #>
                    foreach ($column in $definition.DateColumns) {
                        if ($lastRow -lt 3) { continue }

                        try {
                            $range = '{0}3:{0}{1}' -f $column, $lastRow
                            $sheet.Worksheet.Cells[$range].Style.NumberFormat.Format = 'dd/mm/yyyy hh:mm:ss'
                        }
                        catch {
                            throw "Failed formatting date column in range '$range': $_"
                        }
                    }
                }
                #endregion

                #region Format tables
                if ($definition.TableName) {
                    try {
                        $tableName = $definition.TableName

                        $range = 'A2:{0}' -f $sheet.Worksheet.Dimension.End.Address

                        if ($table = $sheet.Worksheet.Tables[$tableName]) {
                            $table.TableXml.table.ref = $range
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
                        #>
                        try {
                            $table.WorkSheet.Cells.AutoFitColumns()
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
