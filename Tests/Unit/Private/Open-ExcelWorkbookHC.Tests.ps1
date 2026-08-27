#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }
#Requires -Modules ImportExcel

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    # ExcelWorkbook builds the header rows, so its helpers are needed too
    . "$moduleRoot\Private\Get-ColumnCellHC.ps1"
    . "$moduleRoot\Private\WorksheetDefinition.ps1"
    . "$moduleRoot\Private\ExcelWorkbook.ps1"
}

Describe 'Open-ExcelWorkbookHC' {
    Context 'the Excel file does not exist yet' {
        BeforeEach {
            $path = Join-Path $TestDrive "new_$(New-Guid).xlsx"
        }

        It 'reports that it created the file' {
            $workbook = Open-ExcelWorkbookHC -Path $path -Type 'Batch'

            try { $workbook.Created | Should-BeTrue }
            finally { Close-ExcelPackage $workbook.Package -NoSave }
        }

        It 'creates a worksheet for every definition of the type' {
            $workbook = Open-ExcelWorkbookHC -Path $path -Type 'Batch'

            try {
                $expected = @(Get-WorksheetDefinitionHC -Type 'Batch')

                $workbook.Sheet.Keys | Should-BeCollection -Count $expected.Count
            }
            finally { Close-ExcelPackage $workbook.Package -NoSave }
        }

        It 'starts writing data on row 3, below the two header rows' {
            $workbook = Open-ExcelWorkbookHC -Path $path -Type 'Batch'

            try { $workbook.Sheet[$workbook.PlantKey].RowNumber | Should-Be 3 }
            finally { Close-ExcelPackage $workbook.Package -NoSave }
        }

        It 'takes the plant worksheet as the first definition' {
            $workbook = Open-ExcelWorkbookHC -Path $path -Type 'Batch'

            try { $workbook.PlantKey | Should-Be 'plant' }
            finally { Close-ExcelPackage $workbook.Package -NoSave }
        }

        It 'works for <Type>' -ForEach @(
            @{ Type = 'Batch' }
            @{ Type = 'Alarm' }
            @{ Type = 'Sequence' }
        ) {
            $typePath = Join-Path $TestDrive "$Type.xlsx"

            $workbook = Open-ExcelWorkbookHC -Path $typePath -Type $Type

            try { $workbook.Created | Should-BeTrue }
            finally { Close-ExcelPackage $workbook.Package -NoSave }
        }
    }

    Context 'the Excel file already exists' {
        It 'opens it instead of creating it and continues below the last row' {
            $path = Join-Path $TestDrive "existing_$(New-Guid).xlsx"

            $first = Open-ExcelWorkbookHC -Path $path -Type 'Batch'
            $first.Sheet['plant'].Worksheet.Cells['A3'].Value = 'File.xml'
            Close-ExcelPackage $first.Package

            $second = Open-ExcelWorkbookHC -Path $path -Type 'Batch'

            try {
                $second.Created | Should-BeFalse
                $second.Sheet['plant'].RowNumber | Should-Be 4
            }
            finally { Close-ExcelPackage $second.Package -NoSave }
        }
    }

    Context 'a worksheet of an existing file is empty' {
        <#
            A workbook that was emptied or edited by hand. Reading the last row
            of a worksheet with no cells gives nothing, and the old fallback of
            'last row + 1' then started writing on row 1, straight over the two
            header rows, while the run reported success.
        #>
        BeforeEach {
            $script:path = Join-Path $TestDrive "emptied_$(New-Guid).xlsx"

            $workbook = Open-ExcelWorkbookHC -Path $path -Type 'Batch'
            Close-ExcelPackage $workbook.Package

            $package = Open-ExcelPackage -Path $path

            try {
                $package.Workbook.Worksheets.Delete('plantBatchComputers')
                $null = $package.Workbook.Worksheets.Add('plantBatchComputers')
            }
            finally { Close-ExcelPackage $package }
        }

        It 'refuses to open it' {
            { Open-ExcelWorkbookHC -Path $path -Type 'Batch' } | Should-Throw
        }

        It 'names the worksheet and the reason' {
            { Open-ExcelWorkbookHC -Path $path -Type 'Batch' } |
            Should-Throw -ExceptionMessage '*plantBatchComputers*header rows are missing*'
        }
    }

    It 'rejects an unknown type' {
        { Open-ExcelWorkbookHC -Path (Join-Path $TestDrive 'x.xlsx') -Type 'Nope' } |
        Should-Throw
    }
}