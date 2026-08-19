#Requires -Version 7

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
