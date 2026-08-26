#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    # Function under test plus its direct dependencies
    . "$moduleRoot\Private\ConvertTo-DateTimeHC.ps1"
    . "$moduleRoot\Private\Get-MonthKeyHC.ps1"
    . "$moduleRoot\Private\Get-XmlFileMonthHC.ps1"

    . "$root\Tests\Helpers\Fixtures.Xml.ps1"
}

Describe 'Get-XmlFileMonthHC' {
    Context 'Batch: month taken per delivery from load_start_date' {
        It 'returns one month for a file with a single delivery' {
            $xml = New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00'

            $result = Get-XmlFileMonthHC -Xml $xml -Type 'Batch'

            $result.MonthKeys | Should-BeCollection @('202408 August')
            $result.RecordCount | Should-Be 1
            $result.RecordsWithoutDate | Should-Be 0
        }

        It 'returns two months when two deliveries fall in different months' {
            $xml = New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00', '2024-09-02T08:15:00'

            $result = Get-XmlFileMonthHC -Xml $xml -Type 'Batch'

            $result.MonthKeys | Should-BeCollection @('202408 August', '202409 September')
        }

        It 'collapses two deliveries in the same month to one month key' {
            $xml = New-BatchXmlHC -LoadStartDate '2024-08-01T00:00:00', '2024-08-31T23:00:00'

            $result = Get-XmlFileMonthHC -Xml $xml -Type 'Batch'

            $result.MonthKeys | Should-BeCollection @('202408 August')
            $result.RecordCount | Should-Be 2
        }

        It 'counts a delivery without a date as a record without a date' {
            $xml = New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00', ''

            $result = Get-XmlFileMonthHC -Xml $xml -Type 'Batch'

            $result.RecordCount | Should-Be 2
            $result.RecordsWithoutDate | Should-Be 1
            $result.MonthKeys | Should-BeCollection @('202408 August')
        }
    }

    Context 'Alarm: month taken per alarm from raised' {
        It 'returns two months when alarms are raised in different months' {
            $xml = New-AlarmXmlHC -Raised '2024-08-16T09:30:00', '2024-09-02T08:15:00'

            $result = Get-XmlFileMonthHC -Xml $xml -Type 'Alarm'

            $result.MonthKeys | Should-BeCollection @('202408 August', '202409 September')
        }

        It 'flags an alarm without a raised date' {
            $xml = New-AlarmXmlHC -Raised ''

            $result = Get-XmlFileMonthHC -Xml $xml -Type 'Alarm'

            $result.MonthKeys | Should-BeCollection @()
            $result.RecordsWithoutDate | Should-Be 1
        }
    }
}
