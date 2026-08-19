#Requires -Version 7

function Get-BatchRowHC {
    <#
        .SYNOPSIS
            Rows for the deliveries of one batch computer that were loaded in
            the requested month.

        .DESCRIPTION
            The month is decided per delivery on 'load_start_date'. All batches,
            discharging operations and batch items of a delivery follow their
            delivery, so a delivery is never split over two Excel files.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BatchComputer,
        [Parameter(Mandatory)]
        $BatchComputerHeader,
        [Parameter(Mandatory)]
        $PlantHeader,
        [Parameter(Mandatory)]
        [String]$MonthKey,
        [Parameter(Mandatory)]
        [String]$FileName
    )

    Set-Alias -Name 'Convert' -Value ConvertTo-CorrectTypeHC

    foreach ($delivery in $BatchComputer.deliveries.delivery) {
        #region Only deliveries loaded in the requested month
        $loadStartDate = ConvertTo-DateTimeHC $delivery.deliveryHeader.load_start_date

        if (-not $loadStartDate) { Continue }

        if ((Get-MonthKeyHC -Date $loadStartDate) -ne $MonthKey) { Continue }
        #endregion

        #region Get deliveryHeader
        $deliveryHeader = $delivery.deliveryHeader |
        Select-Object -Property load_mix_code_version, load_mix_name,
        load_loading_point, load_qty_unit, load_qty_prod_unit, reuse_qty_unit,
        load_truck, license_plate, ticket_leading_system,
        @{
            Name       = 'load_id_erp'
            Expression = { Convert $_.load_id_erp }
        },
        @{
            Name       = 'reference_delivery'
            Expression = { Convert $_.reference_delivery }
        },
        @{
            Name       = 'original_delivery'
            Expression = { Convert $_.original_delivery }
        },
        @{
            Name       = 'load_id_bcc'
            Expression = { Convert $_.load_id_bcc }
        },
        @{
            Name       = 'load_order_number'
            Expression = { Convert $_.load_order_number }
        },
        @{
            Name       = 'load_order_number_item'
            Expression = { Convert $_.load_order_number_item }
        },
        @{
            Name       = 'load_mix_code'
            Expression = { Convert $_.load_mix_code }
        },
        @{
            Name       = 'load_qty'
            Expression = { Convert $_.load_qty }
        },
        @{
            Name       = 'load_qty_erp'
            Expression = { Convert $_.load_qty_erp }
        },
        @{
            Name       = 'load_qty_prod'
            Expression = { Convert $_.load_qty_prod }
        },
        @{
            Name       = 'reuse_qty'
            Expression = { Convert $_.reuse_qty }
        },
        @{
            Name       = 'ticket_id'
            Expression = { Convert $_.ticket_id }
        },
        @{
            Name       = 'batch_count'
            Expression = { Convert $_.batch_count }
        },
        @{
            Name       = 'load_end_date'
            Expression = { ConvertTo-DateTimeHC $_.load_end_date }
        },
        @{
            Name       = 'ticket_time'
            Expression = { ConvertTo-DateTimeHC $_.ticket_time }
        }

        $deliveryHeader | Add-Member -NotePropertyName 'load_start_date' -NotePropertyValue $loadStartDate -Force
        #endregion

        foreach ($batch in $delivery.batches.batch) {
            #region Get batchHeader
            $batchHeader = $batch.batchHeader |
            Select-Object -Property qty_unit, water_trim_unit, sequence,
            mixing_time_unit, mixer_power_unit, slump_unit,
            mixer_discharge_time_unit, dischargingOperationTimes,
            @{
                Name       = 'batch_id'
                Expression = { Convert $_.batch_id }
            },
            @{
                Name       = 'batch_id_nr'
                Expression = { Convert $_.batch_id_nr }
            },
            @{
                Name       = 'qty'
                Expression = { Convert $_.qty }
            },
            @{
                Name       = 'water_trim'
                Expression = { Convert $_.water_trim }
            },
            @{
                Name       = 'manual'
                Expression = { Convert $_.manual }
            },
            @{
                Name       = 'mixing_time'
                Expression = { Convert $_.mixing_time }
            },
            @{
                Name       = 'mixer_power'
                Expression = { Convert $_.mixer_power }
            },
            @{
                Name       = 'slump'
                Expression = { Convert $_.slump }
            },
            @{
                Name       = 'mixer_discharge_time'
                Expression = { Convert $_.mixer_discharge_time }
            },
            @{
                Name       = 'start_time'
                Expression = { ConvertTo-DateTimeHC $_.start_time }
            },
            @{
                Name       = 'end_time'
                Expression = { ConvertTo-DateTimeHC $_.end_time }
            },
            @{
                Name       = 'aborted'
                Expression = { Convert $_.batchinformation.aborted }
            }
            #endregion

            #region Add row to worksheet deliveriesBatches
            @{
                Key   = 'delivery'
                Cells = @{
                    A  = $FileName
                    B  = $PlantHeader.plant_code
                    C  = $PlantHeader.plant_name
                    D  = $null
                    E  = $BatchComputerHeader.extraction_id
                    F  = $deliveryHeader.load_id_erp
                    G  = $deliveryHeader.reference_delivery
                    H  = $deliveryHeader.original_delivery
                    I  = $deliveryHeader.load_id_bcc
                    J  = $deliveryHeader.load_order_number
                    K  = $deliveryHeader.load_order_number_item
                    L  = $deliveryHeader.load_mix_code
                    M  = $deliveryHeader.load_mix_code_version
                    N  = $deliveryHeader.load_mix_name
                    O  = $deliveryHeader.load_loading_point
                    P  = $deliveryHeader.load_start_date
                    Q  = $deliveryHeader.load_end_date
                    R  = $deliveryHeader.load_qty
                    S  = $deliveryHeader.load_qty_erp
                    T  = $deliveryHeader.load_qty_unit
                    U  = $deliveryHeader.load_qty_prod
                    V  = $deliveryHeader.load_qty_prod_unit
                    W  = $deliveryHeader.reuse_qty
                    X  = $deliveryHeader.reuse_qty_unit
                    Y  = $deliveryHeader.load_truck
                    Z  = $deliveryHeader.license_plate
                    AA = $deliveryHeader.ticket_leading_system
                    AB = $deliveryHeader.ticket_id
                    AC = $deliveryHeader.ticket_time
                    AD = $deliveryHeader.batch_count
                    AE = $batchHeader.batch_id
                    AF = $batchHeader.batch_id_nr
                    AG = $batchHeader.qty
                    AH = $batchHeader.qty_unit
                    AI = $batchHeader.start_time
                    AJ = $batchHeader.end_time
                    AK = $batchHeader.manual
                    AL = $batchHeader.water_trim
                    AM = $batchHeader.water_trim_unit
                    AN = $batchHeader.sequence
                    AO = $batchHeader.mixing_time
                    AP = $batchHeader.mixing_time_unit
                    AQ = $batchHeader.mixer_power
                    AR = $batchHeader.mixer_power_unit
                    AS = $batchHeader.slump
                    AT = $batchHeader.slump_unit
                    AU = $batchHeader.mixer_discharge_time
                    AV = $batchHeader.mixer_discharge_time_unit
                    AW = $batchHeader.aborted
                }
            }
            #endregion

            #region Add rows to worksheet batchesDischarging
            foreach (
                $dischargingOperation in
                $batchHeader.dischargingOperationTimes.discharging
            ) {
                @{
                    Key   = 'discharging'
                    Cells = @{
                        A = $FileName
                        B = $PlantHeader.plant_code
                        C = $PlantHeader.plant_name
                        D = Convert $dischargingOperation.batch_id
                        E = Convert $dischargingOperation.batch_id_nr
                        F = Convert $dischargingOperation.scale_id
                        G = $dischargingOperation.material_type
                        H = ConvertTo-DateTimeHC $dischargingOperation.discharge_start_time
                        I = ConvertTo-DateTimeHC $dischargingOperation.discharge_end_time
                    }
                }
            }
            #endregion

            #region Add rows to worksheet batchItems
            foreach ($batchItem in $batch.batchItems.batchItem) {
                @{
                    Key   = 'item'
                    Cells = @{
                        A  = $FileName
                        B  = $PlantHeader.plant_code
                        C  = $PlantHeader.plant_name
                        D  = $null
                        E  = $BatchComputerHeader.extraction_id
                        F  = $deliveryHeader.load_id_erp
                        G  = $deliveryHeader.reference_delivery
                        H  = $deliveryHeader.original_delivery
                        I  = $deliveryHeader.load_id_bcc
                        J  = $deliveryHeader.load_order_number
                        K  = $deliveryHeader.load_order_number_item
                        L  = $deliveryHeader.load_mix_code
                        M  = $deliveryHeader.load_mix_code_version
                        N  = $deliveryHeader.load_mix_name
                        O  = $deliveryHeader.load_loading_point
                        P  = $deliveryHeader.load_start_date
                        Q  = $deliveryHeader.load_end_date
                        R  = $deliveryHeader.load_qty
                        S  = $deliveryHeader.load_qty_erp
                        T  = $deliveryHeader.load_qty_unit
                        U  = $deliveryHeader.load_qty_prod
                        V  = $deliveryHeader.load_qty_prod_unit
                        W  = $deliveryHeader.reuse_qty
                        X  = $deliveryHeader.reuse_qty_unit
                        Y  = $deliveryHeader.ticket_leading_system
                        Z  = $deliveryHeader.ticket_id
                        AA = $deliveryHeader.ticket_time
                        AB = $batchHeader.batch_id
                        AC = $batchHeader.batch_id_nr
                        AD = $batchHeader.qty
                        AE = $batchHeader.qty_unit
                        AF = $batchHeader.manual
                        AG = $batchHeader.aborted
                        AH = Convert $batchItem.batchiteminformation.pulsecount
                        AI = $batchItem.batchiteminformation.moisture_measure_type
                        AJ = Convert $batchItem.material_code
                        AK = $batchItem.material_name
                        AL = $batchItem.vendor
                        AM = $batchItem.vendor_source
                        AN = $batchItem.vendor_delivery
                        AO = $batchItem.material_type
                        AP = Convert $batchItem.material_target_dry
                        AQ = $batchItem.material_target_dry_unit
                        AR = Convert $batchItem.material_target
                        AS = $batchItem.material_target_unit
                        AT = Convert $batchItem.material_target_adjusted
                        AU = $batchItem.material_target_adjusted_unit
                        AV = Convert $batchItem.material_amount
                        AW = $batchItem.material_amount_unit
                        AX = Convert $batchItem.material_moisture
                        AY = $batchItem.material_moisture_unit
                        AZ = Convert $batchItem.material_moisture_auto
                        BA = Convert $batchItem.material_density
                        BB = $batchItem.material_density_unit
                        BC = Convert $batchItem.material_absorption
                        BD = $batchItem.material_absorption_unit
                        BE = Convert $batchItem.material_solid_content
                        BF = $batchItem.material_solid_content_unit
                        BG = Convert $batchItem.material_temperature
                        BH = $batchItem.material_temperature_unit
                        BI = Convert $batchItem.material_bin_number
                        BJ = $batchItem.material_bin_name
                        BK = Convert $batchItem.scale_id
                        BL = $batchItem.dosingOperationTimes.dosing.material_dosing_start_time
                        BM = $batchItem.dosingOperationTimes.dosing.material_dosing_end_time
                    }
                }
            }
            #endregion
        }
    }
}
