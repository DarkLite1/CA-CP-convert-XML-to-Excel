#Requires -Version 7

<#
    .SYNOPSIS
        The worksheets, headers and formatting used by each type.

    .DESCRIPTION
        The worksheet names and table names are kept exactly as they were in
        the three original scripts. Changing them would make the script unable
        to append to Excel files created by the previous version.
#>

$script:HeaderColor = @{
    plantHeader                  = 'Yellow'
    batchComputerHeader          = 'YellowGreen'
    deliveryHeader               = 'Green'
    batchHeader                  = 'GreenYellow'
    batchHeaderDischargeOptions  = 'GreenYellow'
    batchItem                    = 'Pink'
    alarm                        = 'Pink'
    sequenceParameter            = 'Green'
    subBatch                     = 'GreenYellow'
    parameterAndSubBatchProperty = 'Pink'
}

function Format-WorksheetHeaderHC {
    <#
        .SYNOPSIS
            Apply the background color belonging to a header type.
    #>
    param (
        [Parameter(Mandatory)]
        $Worksheet,
        [Parameter(Mandatory)]
        [String]$Cell,
        [Parameter(Mandatory)]
        [String]$Type
    )

    try {
        $color = $script:HeaderColor[$Type]

        if (-not $color) {
            throw "Type '$Type' not supported"
        }

        $cellStyle = $Worksheet.Cells[$Cell].Style
        $cellStyle.Fill.PatternType = 'MediumGray'
        $cellStyle.Fill.BackgroundColor.SetColor(
            [System.Drawing.Color]::$color
        )
    }
    catch {
        throw "Failed formatting cell '$Cell' of type '$Type': $_"
    }
}

function Get-WorksheetDefinitionHC {
    <#
        .SYNOPSIS
            Return the worksheet layout for a type.

        .DESCRIPTION
            'Key' is used by Get-XmlRowHC to say which worksheet a row belongs
            to. The first worksheet in the list is the 'plant' worksheet and is
            the one used to check whether a file was already added.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('Batch', 'Alarm', 'Sequence')]
        [String]$Type
    )

    $plantWorksheet = @{
        Key            = 'plant'
        Name           = 'plantBatchComputers'
        MergeColumns   = @('A1:B1', 'C1:G1', 'H1:O1')
        FirstHeaderRow = @{
            C1 = 'plantHeader'
            H1 = 'batchComputerHeader'
        }
        TableHeaderRow = @(
            'fileName', 'fileAddedOn', 'countryCode', 'companyCode',
            'companyName', 'plantCode', 'plantName', 'systemType',
            'systemProvider', 'systemId', 'mixerName', 'mixerSize',
            'extractionId', 'fileCreatedOn', 'offset'
        )
        DateColumns    = @('B', 'N')
        FreezeRow      = 2
    }

    switch ($Type) {
        'Batch' {
            $plantWorksheet.TableName = 'plantBatchComputer'

            @(
                $plantWorksheet
                @{
                    Key            = 'delivery'
                    Name           = 'deliveriesBatches'
                    TableName      = 'deliveriesAndBatches'
                    MergeColumns   = @('B1:C1', 'D1:E1', 'F1:AD1', 'AE1:AW1')
                    FirstHeaderRow = @{
                        B1  = 'plantHeader'
                        D1  = 'batchComputerHeader'
                        F1  = 'deliveryHeader'
                        AE1 = 'batchHeader'
                    }
                    TableHeaderRow = @(
                        'fileName', 'plantCode', 'plantName', 'batchComputer',
                        'extractionId', 'loadIdErp', 'referenceDelivery',
                        'originalDelivery', 'loadIdBcc', 'loadOrderNumber',
                        'loadOrderNumberItem', 'loadMixCode',
                        'loadMixCodeVersion', 'loadMixName', 'loadLoadingPoint',
                        'loadStartDate', 'loadEndDate', 'loadQty', 'loadQtyErp',
                        'loadQtyUnit', 'loadQtyProd', 'loadQtyProdUnit',
                        'reuseQty', 'reuseQtyUnit', 'loadTruck', 'licensePlate',
                        'ticketLeadingSystem', 'ticketId', 'ticketTime',
                        'batchCount', 'batchId', 'batchIdNr', 'qty', 'qtyUnit',
                        'startTime', 'endTime', 'manual', 'waterTrim',
                        'waterTrimUnit', 'sequence', 'mixingTime',
                        'mixingTimeUnit', 'mixerPower', 'mixerPowerUnit',
                        'slump', 'slumpUnit', 'mixerDischargeTime',
                        'mixerDischargeTimeUnit', 'aborted'
                    )
                    DateColumns    = @('P', 'Q', 'AC', 'AI', 'AJ')
                    FreezeRow      = 2
                }
                @{
                    Key            = 'discharging'
                    Name           = 'batchesDischarging'
                    TableName      = 'batchesDischarging'
                    MergeColumns   = @('B1:C1', 'D1:I1')
                    FirstHeaderRow = @{
                        B1 = 'plantHeader'
                        D1 = 'batchHeaderDischargeOptions'
                    }
                    TableHeaderRow = @(
                        'fileName', 'plantCode', 'plantName', 'batchId',
                        'batchIdNr', 'scaleId', 'materialType',
                        'dischargeStartTime', 'dischargeEndTime'
                    )
                    DateColumns    = @('H', 'I')
                    FreezeRow      = 2
                }
                @{
                    Key            = 'item'
                    Name           = 'batchItems'
                    TableName      = 'batchItems'
                    MergeColumns   = @(
                        'B1:C1', 'D1:E1', 'F1:AA1', 'AB1:AG1', 'AH1:BM1'
                    )
                    FirstHeaderRow = @{
                        B1  = 'plantHeader'
                        D1  = 'batchComputerHeader'
                        F1  = 'deliveryHeader'
                        AB1 = 'batchHeader'
                        AH1 = 'batchItem'
                    }
                    TableHeaderRow = @(
                        'fileName', 'plantCode', 'plantName', 'batchComputer',
                        'extractionId', 'loadIdErp', 'referenceDelivery',
                        'originalDelivery', 'loadIdBcc', 'loadOrderNumber',
                        'loadOrderNumberItem', 'loadMixCode',
                        'loadMixCodeVersion', 'loadMixName', 'loadLoadingPoint',
                        'loadStartDate', 'loadEndDate', 'loadQty', 'loadQtyErp',
                        'loadQtyUnit', 'loadQtyProd', 'loadQtyProdUnit',
                        'reuseQty', 'reuseQtyUnit', 'ticketLeadingSystem',
                        'ticketId', 'ticketTime', 'batchId', 'batchIdNr', 'qty',
                        'qtyUnit', 'manual', 'aborted', 'pulseCount',
                        'moistureMeasureType', 'materialCode', 'materialName',
                        'vendor', 'vendorSource', 'vendorDelivery',
                        'materialType', 'materialTargetDry',
                        'materialTargetDryUnit', 'materialTarget',
                        'materialTargetUnit', 'materialTargetAdjusted',
                        'materialTargetAdjustedUnit', 'materialAmount',
                        'materialAmountUnit', 'materialMoisture',
                        'materialMoistureUnit', 'materialMoistureAuto',
                        'materialDensity', 'materialDensityUnit',
                        'materialAbsorption', 'materialAbsorptionUnit',
                        'materialSolidContent', 'materialSolidContentUnit',
                        'materialTemperature', 'materialTemperatureUnit',
                        'materialBinNumber', 'materialBinName', 'scaleId',
                        'materialDosingStartTime', 'materialDosingEndTime'
                    )
                    <#
                        'BL' and 'BM' are the dosing start and end times, which
                        are written as real dates like every other timestamp.
                        A column that holds dates and is missing here shows the
                        raw serial number of the date instead of the date.
                    #>
                    DateColumns    = @('P', 'Q', 'AA', 'BL', 'BM')
                    FreezeRow      = 2
                }
            )
            break
        }
        'Alarm' {
            $plantWorksheet.TableName = 'plantBatchComputerTable'

            @(
                $plantWorksheet
                @{
                    Key            = 'alarms'
                    Name           = 'alarms'
                    TableName      = 'alarmsTable'
                    MergeColumns   = @('B1:C1', 'D1:E1', 'F1:AD1')
                    FirstHeaderRow = @{
                        B1 = 'plantHeader'
                        D1 = 'batchComputerHeader'
                        F1 = 'alarm'
                    }
                    TableHeaderRow = @(
                        'fileName', 'plantCode', 'plantName', 'batchComputer',
                        'extractionId', 'alarmId', 'alarmRaised',
                        'alarmDropped', 'alarmHandled', 'alarmClickTimestamp',
                        'alarmText', 'alarmMessage', 'alarmMessage2',
                        'alarmParameter', 'ticketId', 'ticketCode',
                        'ticketReference', 'ticketCustomer', 'ticketProject',
                        'ticketWanted', 'ticketTruck', 'batchId',
                        'batchTimeStart', 'batchTimeReady', 'batchWanted',
                        'batchWeighListCode', 'batchWeighListName',
                        'alarmGroupCode', 'alarmGroupName',
                        'alarmGroupDescription'
                    )
                    DateColumns    = @('G', 'H', 'I', 'J', 'W', 'X')
                    FreezeRow      = 2
                }
            )
            break
        }
        'Sequence' {
            $plantWorksheet.TableName = 'plantBatchComputersTable'

            @(
                $plantWorksheet
                @{
                    Key            = 'sequence'
                    Name           = 'sequences'
                    TableName      = 'sequencesTable'
                    MergeColumns   = @('B1:C1', 'D1:F1', 'G1:M1', 'O1:T1')
                    FirstHeaderRow = @{
                        B1 = 'plantHeader'
                        D1 = 'batchComputerHeader'
                        G1 = 'sequenceParameter'
                        N1 = 'subBatch'
                        O1 = 'parameterAndSubBatchProperty'
                    }
                    TableHeaderRow = @(
                        'fileName', 'plantCode', 'plantName', 'batchComputer',
                        'extractionId', 'fileCreatedOn', 'sequenceId',
                        'sequenceName', 'baseName', 'nrOfSubBatches',
                        'maxBatchSize', 'maxBatchSizeUnit', 'blocked',
                        'subBatchName', 'propertyName', 'prefix', 'itemPath',
                        'upPath', 'value', 'suffix'
                    )
                    DateColumns    = @('F')
                    FreezeRow      = 2
                }
            )
            break
        }
    }
}