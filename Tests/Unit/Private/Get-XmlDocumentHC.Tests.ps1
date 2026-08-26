#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Get-XmlDocumentHC.ps1"
}

Describe 'Get-XmlDocumentHC' {
    It 'reads an XML file into a document' {
        $path = Join-Path $TestDrive 'plant.xml'
        Set-Content -LiteralPath $path -Value '<plant><plantHeader><plant_code>P1</plant_code></plantHeader></plant>'

        $result = Get-XmlDocumentHC -Path $path

        $result | Should-HaveType ([System.Xml.XmlDocument])
        $result.plant.plantHeader.plant_code | Should-Be 'P1'
    }

    It 'throws when the file does not exist' {
        { Get-XmlDocumentHC -Path (Join-Path $TestDrive 'nope.xml') } | Should-Throw
    }

    It 'throws when the file is not valid XML' {
        $path = Join-Path $TestDrive 'broken.xml'
        Set-Content -LiteralPath $path -Value '<plant><unclosed>'

        { Get-XmlDocumentHC -Path $path } | Should-Throw
    }

    It 'does not fetch an external entity' {
        <#
            The resolver is disabled, so a file referring to an external DTD is
            never allowed to make the parser reach out to another system.
        #>
        $path = Join-Path $TestDrive 'external.xml'
        Set-Content -LiteralPath $path -Value @'
<?xml version="1.0"?>
<!DOCTYPE plant SYSTEM "https://example.invalid/plant.dtd">
<plant><plantHeader><plant_code>P1</plant_code></plantHeader></plant>
'@

        $result = Get-XmlDocumentHC -Path $path

        $result.plant.plantHeader.plant_code | Should-Be 'P1'
    }
}
