#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Utils.ps1"
}

Describe 'Get-StringValueHC' {
    It 'returns null for null' {
        Get-StringValueHC -Name $null | Should-BeFalsy
    }

    It 'returns null for whitespace only' {
        Get-StringValueHC -Name '   ' | Should-BeFalsy
    }

    It 'returns a literal value unchanged' {
        Get-StringValueHC -Name 'smtp.contoso.com' | Should-Be 'smtp.contoso.com'
    }

    Context "the 'ENV:' prefix" {
        BeforeAll { $env:PESTER_CXTE_VAR = 'resolved-value' }
        AfterAll { Remove-Item 'Env:\PESTER_CXTE_VAR' -ErrorAction Ignore }

        It 'resolves an existing environment variable' {
            Get-StringValueHC -Name 'ENV:PESTER_CXTE_VAR' | Should-Be 'resolved-value'
        }

        It 'matches the prefix case insensitively' {
            Get-StringValueHC -Name 'env:PESTER_CXTE_VAR' | Should-Be 'resolved-value'
        }

        It 'throws when the environment variable does not exist' {
            { Get-StringValueHC -Name 'ENV:CXTE_DOES_NOT_EXIST' } | Should-Throw '*not found*'
        }
    }
}

Describe 'Get-StringOrDefaultHC' {
    It 'returns the default for a blank value' {
        Get-StringOrDefaultHC '' 'None' | Should-Be 'None'
    }

    It 'returns the value when it is not blank' {
        Get-StringOrDefaultHC 'StartTls' 'None' | Should-Be 'StartTls'
    }
}

Describe 'Remove-BlankValueHC' {
    It 'removes a key whose value is an empty string' {
        $result = Remove-BlankValueHC -Hashtable @{ Keep = 'x'; Drop = '' }

        $result.ContainsKey('Drop') | Should-BeFalse
        $result.Keep | Should-Be 'x'
    }

    It 'removes a key whose value is null' {
        (Remove-BlankValueHC -Hashtable @{ Drop = $null }).ContainsKey('Drop') |
        Should-BeFalse
    }

    It 'keeps an empty array so collection parameters are not dropped' {
        (Remove-BlankValueHC -Hashtable @{ Bcc = @() }).ContainsKey('Bcc') |
        Should-BeTrue
    }

    It 'does not modify the original hashtable' {
        $original = @{ Keep = 'x'; Drop = '' }

        $null = Remove-BlankValueHC -Hashtable $original

        $original.ContainsKey('Drop') | Should-BeTrue
    }
}

Describe 'Get-DatedLogFolderPathHC' {
    BeforeAll {
        $startTime = [datetime]'2024-03-07 09:05:08'
    }

    It 'creates the folder and returns its path' {
        $result = Get-DatedLogFolderPathHC -LogFolder 'TestDrive:\Logs' `
            -ScriptStartTime $startTime -JsonFileName 'CP Batch'

        Test-Path -Path $result | Should-BeTrue
    }

    It 'names the folder yyyy_MM_dd_HHmmss (name)' {
        $result = Get-DatedLogFolderPathHC -LogFolder 'TestDrive:\Logs' `
            -ScriptStartTime $startTime -JsonFileName 'CP Batch'

        Split-Path -Path $result -Leaf | Should-Be '2024_03_07_090508 (CP Batch)'
    }
}
