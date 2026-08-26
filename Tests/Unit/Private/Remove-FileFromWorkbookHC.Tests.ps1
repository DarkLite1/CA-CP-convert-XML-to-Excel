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

    <#
        A real workbook is used, because the function calls DeleteRow on the
        worksheet and reads the dimension back afterwards.
    #>
    function New-FilledWorkbookHC {
        param ([String[]]$FileNames)

        $path = Join-Path $TestDrive "wb_$(New-Guid).xlsx"

        $workbook = Open-ExcelWorkbookHC -Path $path -Type 'Batch'

        $sheet = $workbook.Sheet['plant']

        foreach ($name in $FileNames) {
            $sheet.Worksheet.Cells["A$($sheet.RowNumber)"].Value = $name
            $sheet.RowNumber++
        }

        $workbook
    }
}

Describe 'Remove-FileFromWorkbookHC' {
    It 'removes every row of the given file' {
        $workbook = New-FilledWorkbookHC -FileNames 'A.xml', 'B.xml', 'A.xml'

        try {
            Remove-FileFromWorkbookHC -Workbook $workbook -FileName 'A.xml'

            $names = Get-FileNameInWorkbookHC -Workbook $workbook

            $names.ContainsKey('A.xml') | Should-BeFalse
            $names.ContainsKey('B.xml') | Should-BeTrue
        }
        finally { Close-ExcelPackage $workbook.Package -NoSave }
    }

    It 'leaves the rows of the other files alone' {
        $workbook = New-FilledWorkbookHC -FileNames 'A.xml', 'B.xml', 'C.xml'

        try {
            Remove-FileFromWorkbookHC -Workbook $workbook -FileName 'B.xml'

            $sheet = $workbook.Sheet['plant'].Worksheet

            $sheet.Cells['A3'].Value | Should-Be 'A.xml'
            $sheet.Cells['A4'].Value | Should-Be 'C.xml'
        }
        finally { Close-ExcelPackage $workbook.Package -NoSave }
    }

    It 'corrects the row number so the next write lands below the data' {
        $workbook = New-FilledWorkbookHC -FileNames 'A.xml', 'B.xml', 'A.xml'

        try {
            Remove-FileFromWorkbookHC -Workbook $workbook -FileName 'A.xml'

            $workbook.Sheet['plant'].RowNumber | Should-Be 4
        }
        finally { Close-ExcelPackage $workbook.Package -NoSave }
    }

    It 'does nothing when the file is not in the workbook' {
        $workbook = New-FilledWorkbookHC -FileNames 'A.xml'

        try {
            Remove-FileFromWorkbookHC -Workbook $workbook -FileName 'Other.xml'

            (Get-FileNameInWorkbookHC -Workbook $workbook).ContainsKey('A.xml') |
            Should-BeTrue
        }
        finally { Close-ExcelPackage $workbook.Package -NoSave }
    }

    It 'does nothing on a workbook that only has headers' {
        <#
            There is no 'does not throw' operator: an unhandled exception fails
            the test by itself, so the assertion afterwards is what proves the
            call was harmless.
        #>
        $workbook = New-FilledWorkbookHC -FileNames @()

        try {
            Remove-FileFromWorkbookHC -Workbook $workbook -FileName 'A.xml'

            (Get-FileNameInWorkbookHC -Workbook $workbook).Keys | Should-BeCollection -Count 0
        }
        finally { Close-ExcelPackage $workbook.Package -NoSave }
    }
}
