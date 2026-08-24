#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ConvertTo-DateTimeHC.ps1"
    . "$moduleRoot\Private\Get-MonthKeyHC.ps1"
}

Describe 'ConvertTo-DateTimeHC' {
    Context 'no date' {
        It 'returns null for an empty string' {
            ConvertTo-DateTimeHC -Value '' | Should-BeNull
        }

        It 'returns null for null' {
            ConvertTo-DateTimeHC -Value $null | Should-BeNull
        }

        It 'returns null for whitespace only' {
            ConvertTo-DateTimeHC -Value '   ' | Should-BeNull
        }

        It 'returns null for text that is not a date' {
            ConvertTo-DateTimeHC -Value 'not a date' | Should-BeNull
        }
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
            (ConvertTo-DateTimeHC -Value '2024-09-02').Month | Should-Be 9
        }

        It 'keeps the time exactly as written, without shifting it' {
            $result = ConvertTo-DateTimeHC -Value '2024-08-16T09:30:25+02:00'

            $result.Hour | Should-Be 9
            $result.Minute | Should-Be 30
            $result.Second | Should-Be 25
        }
    }

    Context 'the result does not depend on the time zone of this machine' {
        <#
            A record loaded just after midnight on the 1st belongs to the new
            month. A '[DateTime]' cast would convert it to the local time of the
            machine and, on a server running UTC, move it back to the previous
            month, writing the record to the wrong Excel file.
        #>
        It 'files <Value> under <Expected>' -ForEach @(
            @{ Value = '2025-11-01T00:30:00+01:00'; Expected = '202511 November' }
            @{ Value = '2025-10-31T23:45:00+01:00'; Expected = '202510 October' }
            @{ Value = '2026-01-01T00:15:00+01:00'; Expected = '202601 January' }
            @{ Value = '2025-12-31T23:59:59+01:00'; Expected = '202512 December' }
            @{ Value = '2025-06-01T00:30:00+02:00'; Expected = '202506 June' }
            @{ Value = '2025-07-31T22:10:00+02:00'; Expected = '202507 July' }
        ) {
            $result = ConvertTo-DateTimeHC -Value $Value

            Get-MonthKeyHC -Date $result | Should-Be $Expected
        }

        It 'gives the same result for the same value in a different offset notation' {
            <#
                The same moment written with a different offset keeps its own
                wall clock time, because that is the time the plant recorded.
            #>
            $plusOne = ConvertTo-DateTimeHC -Value '2025-11-01T00:30:00+01:00'
            $utc = ConvertTo-DateTimeHC -Value '2025-10-31T23:30:00+00:00'

            Get-MonthKeyHC -Date $plusOne | Should-Be '202511 November'
            Get-MonthKeyHC -Date $utc | Should-Be '202510 October'
        }
    }
}