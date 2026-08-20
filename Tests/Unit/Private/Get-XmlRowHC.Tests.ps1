#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    # Row builders and their helper dependencies
    . "$moduleRoot\Private\ConvertTo-CorrectTypeHC.ps1"
    . "$moduleRoot\Private\ConvertTo-DateTimeHC.ps1"
    . "$moduleRoot\Private\Get-MonthKeyHC.ps1"
    . "$moduleRoot\Private\XmlRow.ps1"

    . "$root\Tests\Unit\Helpers\Fixtures.Xml.ps1"
}

Describe 'Get-XmlRowHC' {
    Context 'a batch file spanning two months' {
        BeforeEach {
            $xml = New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00', '2024-09-02T08:15:00'

            $splat = @{
                Xml      = $xml
                Type     = 'Batch'
                FileName = 'File.xml'
                AddedOn  = [datetime]'2024-10-01'
            }
        }

        It 'returns only the August delivery for the August month key' {
            $rows = @(Get-XmlRowHC @splat -MonthKey '202408 August')

            $deliveryRows = @($rows.Where({ $_.Key -eq 'delivery' }))

            $deliveryRows | Should-BeCollection -Count 1

            $loadStart = $deliveryRows[0].Cells['P']
            (Get-MonthKeyHC -Date $loadStart) | Should-Be '202408 August'
        }

        It 'returns only the September delivery for the September month key' {
            $rows = @(Get-XmlRowHC @splat -MonthKey '202409 September')

            $deliveryRows = @($rows.Where({ $_.Key -eq 'delivery' }))

            $deliveryRows | Should-BeCollection -Count 1

            $loadStart = $deliveryRows[0].Cells['P']
            (Get-MonthKeyHC -Date $loadStart) | Should-Be '202409 September'
        }

        It 'writes exactly one plant row for a month that has records' {
            $rows = @(Get-XmlRowHC @splat -MonthKey '202408 August')

            @($rows.Where({ $_.Key -eq 'plant' })) | Should-BeCollection -Count 1
        }

        It 'writes no rows at all for a month with no records' {
            $rows = @(Get-XmlRowHC @splat -MonthKey '202401 January')

            $rows | Should-BeCollection @()
        }

        It 'does not emit a plant row for a batch computer with no records that month' {
            $rows = @(Get-XmlRowHC @splat -MonthKey '202401 January')

            @($rows.Where({ $_.Key -eq 'plant' })) | Should-BeCollection @()
        }
    }

    Context 'the file name and added-on stamp are carried onto the plant row' {
        It 'puts the file name in column A and the stamp in column B' {
            $xml = New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00'

            $rows = @(Get-XmlRowHC -Xml $xml -Type 'Batch' -MonthKey '202408 August' `
                    -FileName 'MyFile.xml' -AddedOn ([datetime]'2024-10-01'))

            $plant = $rows.Where({ $_.Key -eq 'plant' })[0]

            $plant.Cells['A'] | Should-Be 'MyFile.xml'
            $plant.Cells['B'] | Should-Be ([datetime]'2024-10-01')
        }
    }
}