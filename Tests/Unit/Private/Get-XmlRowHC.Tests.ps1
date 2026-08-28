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

    Context 'an element name carrying a prefix' {
        <#
            Get-XmlFileMonthHC decides which month a record belongs to and
            matches on the local name, so the row builder has to read the values
            by local name as well. Keyed on the full name the two disagree on a
            file that carries prefixes: the month is found and not a single
            value of the record is.
        #>
        It 'reads the value by its local name' {
            $xml = [xml]@'
<plant xmlns:cp="urn:contoso:cp">
  <plantHeader>
    <cp:plant_code>P1</cp:plant_code>
    <plant_name>Plant one</plant_name>
  </plantHeader>
  <batchComputers></batchComputers>
</plant>
'@

            $map = Get-ChildValueMapHC -Node $xml.plant.plantHeader

            $map['plant_code'] | Should-Be 'P1'
            $map['plant_name'] | Should-Be 'Plant one'
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
            $script:discharging = $rows.Where({ $_.Key -eq 'discharging' })[0]
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

        It 'reads the discharging operations of a batch' {
            <#
                A container that sits inside the batch header rather than
                beside it, so it is reached with an XPath of its own. Reading
                it wrongly loses every discharging row without failing.
            #>
            $discharging | Should-BeTruthy
            $discharging.Cells['D'] | Should-Be 'B1'
            $discharging.Cells['F'] | Should-Be '0007'
            $discharging.Cells['G'] | Should-Be 'Cement'
            $discharging.Cells['H'] | Should-Be ([datetime]'2024-08-16T09:33:00')
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