#Requires -Version 7

function Get-XmlPropertyPathHC {
    <#
        .SYNOPSIS
            The XML property that holds the date for a type, used in error
            messages.

        .DESCRIPTION
            Keeps the human readable property path in one place so the date
            reader and its error messages always name the same field.

        .EXAMPLE
            Get-XmlPropertyPathHC -Type 'Alarm'
            # 'plant.batchComputers.batchComputer.alarms.alarm.raised'
    #>
    param (
        [Parameter(Mandatory)]
        [ValidateSet('Batch', 'Alarm', 'Sequence')]
        [String]$Type
    )

    switch ($Type) {
        'Batch' {
            'plant.batchComputers.batchComputer.deliveries.delivery.deliveryHeader.load_start_date'
            break
        }
        'Alarm' {
            'plant.batchComputers.batchComputer.alarms.alarm.raised'
            break
        }
        'Sequence' {
            'plant.batchComputers.batchComputer.batchComputerHeader.file_created_on'
            break
        }
    }
}
