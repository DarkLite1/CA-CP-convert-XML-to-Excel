#Requires -Version 7

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
                    DateColumns    = @('P', 'Q', 'AA')
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
