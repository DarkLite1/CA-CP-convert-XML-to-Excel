#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Get-XmlPropertyPathHC.ps1"
}

Describe 'Get-XmlPropertyPathHC' {
    It 'points batch at load_start_date' {
        Get-XmlPropertyPathHC -Type 'Batch' | Should-MatchString 'load_start_date'
    }

    It 'points alarm at the raised property' {
        Get-XmlPropertyPathHC -Type 'Alarm' | Should-MatchString 'alarm\.raised'
    }

    It 'points sequence at file_created_on' {
        Get-XmlPropertyPathHC -Type 'Sequence' | Should-MatchString 'file_created_on'
    }

    It 'rejects an unknown type' {
        { Get-XmlPropertyPathHC -Type 'Nope' } | Should-Throw
    }
}
