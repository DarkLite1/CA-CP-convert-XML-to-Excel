#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ConvertTo-TextHC.ps1"
}

Describe 'ConvertTo-TextHC' {
    It 'returns nothing for an empty string' {
        ConvertTo-TextHC -Value '' | Should-BeFalsy
    }

    It 'returns nothing for null' {
        ConvertTo-TextHC -Value $null | Should-BeFalsy
    }

    Context 'values that only look like a number' {
        It 'keeps the leading zeros of an order number' {
            $result = ConvertTo-TextHC -Value '0000123456'

            $result | Should-Be '0000123456'
            $result | Should-HaveType ([string])
        }

        It 'keeps every digit of an identifier too long for a double' {
            <#
                A double holds about 15 significant digits, so this value comes
                back as 1.23456789012346E+18 the moment it is treated as a
                number. The comparison is on the exact text for that reason.
            #>
            ConvertTo-TextHC -Value '1234567890123456789' |
            Should-Be '1234567890123456789'
        }

        It 'keeps a decimal looking code as written' {
            ConvertTo-TextHC -Value '1.10' | Should-Be '1.10'
        }

        It "keeps the text 'true' as text" {
            $result = ConvertTo-TextHC -Value 'true'

            $result | Should-Be 'true'
            $result | Should-HaveType ([string])
        }
    }

    Context 'plain text' {
        It 'returns a string unchanged' {
            ConvertTo-TextHC -Value 'ABC-123' | Should-Be 'ABC-123'
        }

        It 'keeps surrounding whitespace, because it is part of the value' {
            ConvertTo-TextHC -Value ' A ' | Should-Be ' A '
        }
    }

    Context 'a value that is not a string' {
        It 'returns it as a string' {
            $result = ConvertTo-TextHC -Value 42

            $result | Should-Be '42'
            $result | Should-HaveType ([string])
        }
    }
}