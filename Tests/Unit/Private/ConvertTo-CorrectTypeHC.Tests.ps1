#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ConvertTo-CorrectTypeHC.ps1"
}

Describe 'ConvertTo-CorrectTypeHC' {
    It 'returns nothing for an empty string' {
        ConvertTo-CorrectTypeHC -Value '' | Should-BeFalsy
    }

    It 'returns nothing for null' {
        ConvertTo-CorrectTypeHC -Value $null | Should-BeFalsy
    }

    Context 'boolean text' {
        It "converts 'true' to a boolean true" {
            $result = ConvertTo-CorrectTypeHC -Value 'true'
            $result | Should-BeTrue
            $result | Should-HaveType ([bool])
        }

        It "converts 'false' to a boolean false" {
            $result = ConvertTo-CorrectTypeHC -Value 'false'
            $result | Should-BeFalse
            $result | Should-HaveType ([bool])
        }
    }

    Context 'numeric text' {
        It 'converts an integer string to a double' {
            $result = ConvertTo-CorrectTypeHC -Value '42'
            $result | Should-Be 42
            $result | Should-HaveType ([double])
        }

        It 'converts a decimal string to a double' {
            ConvertTo-CorrectTypeHC -Value '12.5' | Should-Be 12.5
        }
    }

    Context 'plain text' {
        It 'returns a non-numeric string unchanged' {
            ConvertTo-CorrectTypeHC -Value 'ABC' | Should-Be 'ABC'
        }

        It 'returns a mixed string unchanged' {
            ConvertTo-CorrectTypeHC -Value '12ABC' | Should-Be '12ABC'
        }
    }
}
