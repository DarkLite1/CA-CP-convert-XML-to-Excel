#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\ErrorHandling.ps1"
}

Describe 'Add-ErrorHC' {
    It 'appends a structured record to the accumulator' {
        $errors = [System.Collections.Generic.List[object]]::new()

        Add-ErrorHC -Type 'FatalError' -Name 'Bad file' -Message 'Missing date.' `
            -Category 'Date' -SystemErrors ([ref]$errors)

        $errors | Should-BeCollection -Count 1
        $errors[0].Type | Should-Be 'FatalError'
        $errors[0].Name | Should-Be 'Bad file'
        $errors[0].Category | Should-Be 'Date'
    }

    It 'accepts Warning as a type' {
        $errors = [System.Collections.Generic.List[object]]::new()

        Add-ErrorHC -Type 'Warning' -Name 'x' -Message 'y' -Category 'Export' `
            -SystemErrors ([ref]$errors)

        $errors[0].Type | Should-Be 'Warning'
    }

    It 'rejects a type outside FatalError and Warning' {
        $errors = [System.Collections.Generic.List[object]]::new()

        { Add-ErrorHC -Type 'Fatal' -Name 'x' -Message 'y' -Category 'z' `
                -SystemErrors ([ref]$errors) } | Should-Throw
    }
}
