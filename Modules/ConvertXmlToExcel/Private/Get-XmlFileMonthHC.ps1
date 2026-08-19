#Requires -Version 7

function Get-XmlFileMonthHC {
    <#
        .SYNOPSIS
            Return every month present in an XML file.

        .DESCRIPTION
            One XML file can hold records for more than one month. For example
            a batch file with one delivery loaded today and another delivery
            loaded tomorrow, where tomorrow falls in the next month.

            This function returns one entry per month found in the file, so the
            caller can write each part of the file to the matching Excel file.

            The date used depends on the type:

            Batch    : delivery.deliveryHeader.load_start_date
            Alarm    : alarm.raised
            Sequence : batchComputer.batchComputerHeader.file_created_on

            Records without a usable date are counted in 'RecordsWithoutDate'.
            They are never silently written to an arbitrary Excel file.

        .PARAMETER Xml
            The XML document, as returned by Get-XmlDocumentHC.

        .PARAMETER Type
            'Batch', 'Alarm' or 'Sequence'.

        .EXAMPLE
            $xml = Get-XmlDocumentHC -Path 'C:\Data\File.xml'
            Get-XmlFileMonthHC -Xml $xml -Type 'Batch'
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Xml,
        [Parameter(Mandatory)]
        [ValidateSet('Batch', 'Alarm', 'Sequence')]
        [String]$Type
    )

    $dates = @()
    $recordsWithoutDate = 0

    switch ($Type) {
        'Batch' {
            $rawDates = @(
                $Xml.plant.batchComputers.batchComputer.deliveries.delivery.deliveryHeader.load_start_date
            )
            break
        }
        'Alarm' {
            $rawDates = @(
                $Xml.plant.batchComputers.batchComputer.alarms.alarm.raised
            )
            break
        }
        'Sequence' {
            $rawDates = @(
                $Xml.plant.batchComputers.batchComputer.batchComputerHeader.file_created_on
            )
            break
        }
    }

    foreach ($rawDate in $rawDates) {
        $date = ConvertTo-DateTimeHC -Value $rawDate

        if ($date) {
            $dates += $date
        }
        else {
            $recordsWithoutDate++
        }
    }

    [PSCustomObject]@{
        MonthKeys          = @(
            $dates | ForEach-Object { Get-MonthKeyHC -Date $_ } |
            Sort-Object -Unique
        )
        RecordCount        = $rawDates.Count
        RecordsWithoutDate = $recordsWithoutDate
    }
}
