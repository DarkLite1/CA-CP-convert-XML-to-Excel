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

    <#
        The function reads from disk now instead of taking a loaded document,
        so the in memory fixtures are written to a temporary file first. This
        keeps the fixtures as the single place where the XML shape is defined.
    #>
    function Save-TestXmlHC {
        param (
            [Parameter(Mandatory)]
            $Xml
        )

        $path = Join-Path ([System.IO.Path]::GetTempPath()) "$([Guid]::NewGuid()).xml"

        $Xml.Save($path)

        $path
    }
}

Describe 'Get-XmlFileMonthHC' {
    Context 'Batch: month taken per delivery from load_start_date' {
        It 'returns one month for a file with a single delivery' {
            $path = Save-TestXmlHC (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00')

            $result = Get-XmlFileMonthHC -Path $path -Type 'Batch'

            $result.MonthKeys | Should-BeCollection @('202408 August')
            $result.RecordCount | Should-Be 1
            $result.RecordsWithoutDate | Should-Be 0
        }

        It 'returns two months when two deliveries fall in different months' {
            $path = Save-TestXmlHC (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00', '2024-09-02T08:15:00')

            $result = Get-XmlFileMonthHC -Path $path -Type 'Batch'

            $result.MonthKeys | Should-BeCollection @('202408 August', '202409 September')
        }

        It 'collapses two deliveries in the same month to one month key' {
            $path = Save-TestXmlHC (New-BatchXmlHC -LoadStartDate '2024-08-01T00:00:00', '2024-08-31T23:00:00')

            $result = Get-XmlFileMonthHC -Path $path -Type 'Batch'

            $result.MonthKeys | Should-BeCollection @('202408 August')
            $result.RecordCount | Should-Be 2
        }

        It 'counts a delivery without a date as a record without a date' {
            $path = Save-TestXmlHC (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00', '')

            $result = Get-XmlFileMonthHC -Path $path -Type 'Batch'

            $result.RecordCount | Should-Be 2
            $result.RecordsWithoutDate | Should-Be 1
            $result.MonthKeys | Should-BeCollection @('202408 August')
        }
    }

    Context 'Alarm: month taken per alarm from raised' {
        It 'returns two months when alarms are raised in different months' {
            $path = Save-TestXmlHC (New-AlarmXmlHC -Raised '2024-08-16T09:30:00', '2024-09-02T08:15:00')

            $result = Get-XmlFileMonthHC -Path $path -Type 'Alarm'

            $result.MonthKeys | Should-BeCollection @('202408 August', '202409 September')
        }

        It 'flags an alarm without a raised date' {
            $path = Save-TestXmlHC (New-AlarmXmlHC -Raised '')

            $result = Get-XmlFileMonthHC -Path $path -Type 'Alarm'

            $result.MonthKeys | Should-BeCollection @()
            $result.RecordsWithoutDate | Should-Be 1
        }
    }

    Context 'the sample files in TestData' {
        It 'reads the months of every type' -ForEach @(
            @{ FileName = 'Batch.xml'; Type = 'Batch' }
            @{ FileName = 'Alarm.xml'; Type = 'Alarm' }
            @{ FileName = 'Sequence.xml'; Type = 'Sequence' }
        ) {
            $path = Join-Path (Resolve-Path "$PSScriptRoot\..\..\TestData") $FileName

            $result = Get-XmlFileMonthHC -Path $path -Type $Type

            $result.RecordCount | Should-BeGreaterThan 0
            $result.RecordsWithoutDate | Should-Be 0
            $result.MonthKeys.Count | Should-BeGreaterThan 0
        }
    }

    Context 'a date element with the same name elsewhere in the file' {
        It 'only counts the element inside its own parent' {
            <#
                'file_created_on' also sits in the batch computer header of a
                batch file. A batch file must still be read from
                'load_start_date' only.
            #>
            $path = Save-TestXmlHC (New-BatchXmlHC -LoadStartDate '2024-08-16T09:30:00')

            $result = Get-XmlFileMonthHC -Path $path -Type 'Batch'

            $result.RecordCount | Should-Be 1
        }
    }
}