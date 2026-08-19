#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Get-MonthKeyHC.ps1"
}

Describe 'Get-MonthKeyHC' {
    It 'formats a date as yyyyMM MMMM' {
        Get-MonthKeyHC -Date ([datetime]'2024-08-16') | Should-Be '202408 August'
    }

    It 'gives the same key for two dates in the same month' {
        $a = Get-MonthKeyHC -Date ([datetime]'2024-08-01')
        $b = Get-MonthKeyHC -Date ([datetime]'2024-08-31')
        $a | Should-Be $b
    }

    It 'gives a different key across a month boundary' {
        $aug = Get-MonthKeyHC -Date ([datetime]'2024-08-31')
        $sep = Get-MonthKeyHC -Date ([datetime]'2024-09-01')
        $aug | Should-NotBe $sep
    }

    It 'sorts chronologically as text because the year and month lead' {
        $keys = @(
            Get-MonthKeyHC -Date ([datetime]'2024-09-01')
            Get-MonthKeyHC -Date ([datetime]'2024-08-01')
            Get-MonthKeyHC -Date ([datetime]'2025-01-01')
        ) | Sort-Object

        $keys[0] | Should-Be '202408 August'
        $keys[1] | Should-Be '202409 September'
        $keys[2] | Should-Be '202501 January'
    }
}
