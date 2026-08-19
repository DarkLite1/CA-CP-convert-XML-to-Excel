#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ConvertTo-DateTimeHC.ps1"
}

Describe 'ConvertTo-DateTimeHC' {
    It 'returns null for an empty string' {
        ConvertTo-DateTimeHC -Value '' | Should-BeNull
    }

    It 'returns null for null' {
        ConvertTo-DateTimeHC -Value $null | Should-BeNull
    }

    It 'returns null for text that is not a date' {
        ConvertTo-DateTimeHC -Value 'not a date' | Should-BeNull
    }

    Context 'valid dates' {
        It 'converts an ISO date with time zone to a DateTime' {
            $result = ConvertTo-DateTimeHC -Value '2024-08-16T09:30:25+02:00'
            $result | Should-HaveType ([datetime])
            $result.Year | Should-Be 2024
            $result.Month | Should-Be 8
            $result.Day | Should-Be 16
        }

        It 'converts a plain date string to a DateTime' {
            $result = ConvertTo-DateTimeHC -Value '2024-09-02'
            $result.Month | Should-Be 9
        }
    }
}
