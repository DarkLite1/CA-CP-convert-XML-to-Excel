#Requires -Version 7

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
