#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleManifest = "$root\Modules\ConvertXmlToExcel\ConvertXmlToExcel.psd1"

    Import-Module $moduleManifest -Force
}

AfterAll {
    Remove-Module ConvertXmlToExcel -Force -ErrorAction Ignore
}

Describe 'Invoke-ConvertXmlToExcel' {
    It 'is exported by the module' {
        Get-Command -Module ConvertXmlToExcel -Name Invoke-ConvertXmlToExcel |
        Should-BeTruthy
    }

    It 'is the only public function' {
        (Get-Command -Module ConvertXmlToExcel).Name |
        Should-BeCollection @('Invoke-ConvertXmlToExcel')
    }

    It 'requires only ImportExcel as an external module' {
        $manifest = Import-PowerShellDataFile "$root\Modules\ConvertXmlToExcel\ConvertXmlToExcel.psd1"

        $required = @($manifest.RequiredModules)

        $required | Should-BeCollection -Count 1
        $required[0].ModuleName | Should-Be 'ImportExcel'
    }

    Context 'mandatory parameters' {
        BeforeAll {
            $params = (Get-Command Invoke-ConvertXmlToExcel).Parameters
        }

        It 'requires <Name>' -ForEach @(
            @{ Name = 'ConfigurationJsonFile' }
            @{ Name = 'ScriptPath' }
            @{ Name = 'ModulePath' }
            @{ Name = 'SystemErrors' }
        ) {
            $attribute = $params[$Name].Attributes |
            Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] }

            $attribute.Mandatory | Should-BeTrue
        }
    }
}
