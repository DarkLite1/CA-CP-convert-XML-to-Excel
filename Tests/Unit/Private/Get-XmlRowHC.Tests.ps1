#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    # Row builders and their helper dependencies
    . "$moduleRoot\Private\ConvertTo-CorrectTypeHC.ps1"
    . "$moduleRoot\Private\ConvertTo-TextHC.ps1"
    . "$moduleRoot\Private\ConvertTo-DateTimeHC.ps1"
    . "$moduleRoot\Private\Get-MonthKeyHC.ps1"
    . "$moduleRoot\Private\XmlRow.ps1"

    . "$root\Tests\Helpers\Fixtures.Xml.ps1"
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

    Context 'the type a cell gets' {
        <#
            An identifier turned into a number is damaged the moment it is
            written and cannot be repaired afterwards, so the delivery row is
            checked value by value: identifiers as text, quantities as numbers.
        #>
        BeforeAll {
            $xml = New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00'

            $rows = @(Get-XmlRowHC -Xml $xml -Type 'Batch' `
                    -MonthKey '202408 August' -FileName 'File.xml' `
                    -AddedOn ([datetime]'2024-10-01'))

            $script:delivery = $rows.Where({ $_.Key -eq 'delivery' })[0]
            $script:item = $rows.Where({ $_.Key -eq 'item' })[0]
        }

        It 'keeps the leading zeros of an order number' {
            $delivery.Cells['J'] | Should-Be '0000123456'
            $delivery.Cells['J'] | Should-HaveType ([string])
        }

        It 'keeps every digit of an identifier too long for a double' {
            <#
                As a double this value comes back as 1.23456789012346E+18, so
                the last digits are gone. Compared as text for that reason.
            #>
            $delivery.Cells['I'] | Should-Be '1234567890123456789'
            $delivery.Cells['I'] | Should-HaveType ([string])
        }

        It 'keeps a quantity as a real number, so Excel can total it' {
            $delivery.Cells['R'] | Should-Be 10
            $delivery.Cells['R'] | Should-HaveType ([double])
        }

        It 'keeps the batch id as text and its counter as a number' {
            $delivery.Cells['AE'] | Should-HaveType ([string])
            $delivery.Cells['AF'] | Should-HaveType ([double])
        }

        It 'writes the dosing times as real dates' {
            <#
                These two were the only timestamps left as text, so they sorted
                alphabetically and a date filter did not see them.
            #>
            $item.Cells['BL'] | Should-HaveType ([datetime])
            $item.Cells['BM'] | Should-HaveType ([datetime])

            $item.Cells['BL'] | Should-Be ([datetime]'2024-08-16T09:31:00')
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