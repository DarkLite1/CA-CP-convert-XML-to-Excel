#Requires -Version 7

<#
    .SYNOPSIS
        Turn the nodes of an XML file into Excel rows for a single month.

    .DESCRIPTION
        Get-XmlRowHC returns the rows of one XML file that belong to one month.
        Calling it twice with two different months returns two different sets of
        rows, which is what makes it possible for one XML file to end up in more
        than one Excel file.

        Rows are returned in the order they have to be written:

            @{ Key = 'plant'    ; Cells = @{ A = 'file.xml'; B = ... } }
            @{ Key = 'delivery' ; Cells = @{ A = 'file.xml'; B = ... } }

        'Key' matches the worksheet key returned by Get-WorksheetDefinitionHC.

        A plant row is only returned for a batch computer that actually has
        records in the requested month. Without this an Excel file would show a
        plant and batch computer that has no rows anywhere else in the workbook.

        Two aliases decide the type of a cell, and which one a value gets is a
        statement about what that value IS:

        - 'Convert' (ConvertTo-CorrectTypeHC) is for values that are measured,
          counted or answered with yes or no: quantities, times, temperatures,
          percentages, flags. Those become a real number or a real boolean, so
          Excel can sort and total them.

        - 'Text' (ConvertTo-TextHC) is for values that identify or describe
          something: identifiers, codes, order and reference numbers, license
          plates, names and messages. Those stay text even when they consist of
          digits only, because a number is the wrong shape for them: an order
          number '0000123456' loses its leading zeros and an 18 digit
          identifier loses its last digits to the precision of a double. Both
          are silent and neither can be repaired after the Excel file is
          written.

        In doubt the value is an identifier: an identifier stored as text is
        merely unsortable, while an identifier stored as a number is wrong.
#>

function Get-ChildValueMapHC {
    <#
        .SYNOPSIS
            The text of every leaf child of an XML node, by element name.

        .DESCRIPTION
            Reading a value as '$node.material_name' goes through PowerShell's
            XmlNode adapter, which scans the child nodes and builds a member
            table to resolve the name. That happens again for every property
            read, and a single batch item is read about thirty times, so the
            children of one node were walked over and over.

            This walks them once and hands back a hash table. Every read after
            that is a hash lookup.

            A plain PowerShell hash table is used on purpose, because it is
            case insensitive. The XML holds 'itempath' and 'uppath' while the
            code asks for 'itemPath' and 'upPath', which worked because the
            adapter is case insensitive too. A
            'Dictionary[String, String]' would be case sensitive and would
            quietly start returning nothing for those two.

            Only leaf elements are included. A child that holds other elements,
            such as 'dosingOperationTimes' or 'alarmGroups', is skipped: its
            'InnerText' would be the text of the whole subtree glued together,
            which is not what the adapter returns and not what anyone wants.
            Those are still read from the node itself.

            An empty element gives an empty string, which is what the adapter
            returns for '<load_loading_point />' as well.

        .PARAMETER Node
            The XML node whose children are wanted. Missing nodes are allowed
            and give an empty map, so callers do not have to test first.

        .EXAMPLE
            $map = Get-ChildValueMapHC -Node $batchItem
            $map['material_name']
    #>
    param (
        $Node
    )

    $map = @{}

    if (-not $Node) { return $map }

    foreach ($child in $Node.ChildNodes) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }

        $firstChild = $child.FirstChild

        if (-not $firstChild) {
            $map[$child.Name] = ''
            continue
        }

        # Holds other elements, so it is a container and not a value
        if ($firstChild.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            continue
        }

        $map[$child.Name] = $child.InnerText
    }

    $map
}

function Get-XmlRowHC {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Xml,
        [Parameter(Mandatory)]
        [ValidateSet('Batch', 'Alarm', 'Sequence')]
        [String]$Type,
        [Parameter(Mandatory)]
        [String]$MonthKey,
        [Parameter(Mandatory)]
        [String]$FileName,
        [Parameter(Mandatory)]
        [DateTime]$AddedOn
    )

    Set-Alias -Name 'Convert' -Value ConvertTo-CorrectTypeHC
    Set-Alias -Name 'Text' -Value ConvertTo-TextHC

    #region Get plantHeader
    <#
        A PSCustomObject literal instead of 'Select-Object'. Select-Object sets
        up a full pipeline and invokes a script block for every calculated
        property of every record it is handed. On the headers built once per
        file that hardly matters, but the same shape is used per delivery, per
        batch and per sequence parameter further down, so it is done the same
        way everywhere to keep the file readable.
    #>
    $header = Get-ChildValueMapHC -Node $Xml.plant.plantHeader

    $plantHeader = [PSCustomObject]@{
        country_code = $header['country_code']
        company_code = $header['company_code']
        company_name = $header['company_name']
        plant_code   = $header['plant_code']
        plant_name   = $header['plant_name']
    }
    #endregion

    foreach ($batchComputer in $Xml.plant.batchComputers.batchComputer) {
        #region Get batchComputerHeader
        $header = Get-ChildValueMapHC -Node $batchComputer.batchComputerHeader

        $batchComputerHeader = [PSCustomObject]@{
            system_type     = $header['system_type']
            system_provider = $header['system_provider']
            mixer_name      = $header['mixer_name']
            offset          = $header['offset']
            system_id       = Convert $header['system_id']
            mixer_size      = Convert $header['mixer_size']
            extraction_id   = Convert $header['extraction_id']
            file_created_on = ConvertTo-DateTimeHC $header['file_created_on']
        }
        #endregion

        #region Get the child rows for the requested month
        $childRow = switch ($Type) {
            'Batch' {
                $params = @{
                    BatchComputer       = $batchComputer
                    BatchComputerHeader = $batchComputerHeader
                    PlantHeader         = $plantHeader
                    MonthKey            = $MonthKey
                    FileName            = $FileName
                }
                Get-BatchRowHC @params
                break
            }
            'Alarm' {
                $params = @{
                    BatchComputer       = $batchComputer
                    BatchComputerHeader = $batchComputerHeader
                    PlantHeader         = $plantHeader
                    MonthKey            = $MonthKey
                    FileName            = $FileName
                }
                Get-AlarmRowHC @params
                break
            }
            'Sequence' {
                $params = @{
                    BatchComputer       = $batchComputer
                    BatchComputerHeader = $batchComputerHeader
                    PlantHeader         = $plantHeader
                    MonthKey            = $MonthKey
                    FileName            = $FileName
                }
                Get-SequenceRowHC @params
                break
            }
        }

        $childRow = @($childRow)
        #endregion

        #region Nothing in this month, skip the batch computer completely
        if (-not $childRow.Count) {
            Write-Verbose "File '$FileName' month '$MonthKey': no records for batch computer '$($batchComputerHeader.extraction_id)'"
            Continue
        }
        #endregion

        #region Add row to worksheet plantBatchComputers
        @{
            Key   = 'plant'
            Cells = @{
                A = $FileName
                B = $AddedOn
                C = $plantHeader.country_code
                D = $plantHeader.company_code
                E = $plantHeader.company_name
                F = $plantHeader.plant_code
                G = $plantHeader.plant_name
                H = $batchComputerHeader.system_type
                I = $batchComputerHeader.system_provider
                J = $batchComputerHeader.system_id
                K = $batchComputerHeader.mixer_name
                L = $batchComputerHeader.mixer_size
                M = $batchComputerHeader.extraction_id
                N = $batchComputerHeader.file_created_on
                O = $batchComputerHeader.offset
            }
        }
        #endregion

        $childRow
    }
}

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
    Set-Alias -Name 'Text' -Value ConvertTo-TextHC

    <#
        Worked out once instead of formatting the date of every record into a
        month key and comparing two strings. 'End' is the first moment of the
        next month, so the test below is 'at or after Start, before End'.
    #>
    $monthRange = Get-MonthRangeHC -MonthKey $MonthKey

    foreach ($delivery in $BatchComputer.deliveries.delivery) {
        <#
            Built before the month test rather than after it. Walking the
            children once costs about what a single adapter read costs, and
            every delivery that is kept then reads 24 values from it for free.
        #>
        $header = Get-ChildValueMapHC -Node $delivery.deliveryHeader

        #region Only deliveries loaded in the requested month
        $loadStartDate = ConvertTo-DateTimeHC $header['load_start_date']

        if (-not $loadStartDate) { Continue }

        if (
            ($loadStartDate -lt $monthRange.Start) -or
            ($loadStartDate -ge $monthRange.End)
        ) { Continue }
        #endregion

        #region Get deliveryHeader
        <#
            Built once per delivery and read again by every batch and every
            batch item of that delivery, so it stays a cache. Only the way it
            is built changed: 'Select-Object' invoked a script block for each
            of the 16 calculated properties, for every delivery in the file.

            'load_start_date' is set here as a normal property. It was added
            afterwards with 'Add-Member', which rebuilds the member table of
            the object, again once per delivery.
        #>
        $deliveryHeader = [PSCustomObject]@{
            load_mix_code_version  = $header['load_mix_code_version']
            load_mix_name          = $header['load_mix_name']
            load_loading_point     = $header['load_loading_point']
            load_qty_unit          = $header['load_qty_unit']
            load_qty_prod_unit     = $header['load_qty_prod_unit']
            reuse_qty_unit         = $header['reuse_qty_unit']
            load_truck             = $header['load_truck']
            license_plate          = $header['license_plate']
            ticket_leading_system  = $header['ticket_leading_system']
            <#
                Identifiers and codes, kept as text. These come from the source
                system and have to keep matching it, leading zeros and all.
            #>
            load_id_erp            = Text $header['load_id_erp']
            reference_delivery     = Text $header['reference_delivery']
            original_delivery      = Text $header['original_delivery']
            load_id_bcc            = Text $header['load_id_bcc']
            load_order_number      = Text $header['load_order_number']
            load_order_number_item = Text $header['load_order_number_item']
            load_mix_code          = Text $header['load_mix_code']
            ticket_id              = Text $header['ticket_id']
            # Quantities, so real numbers
            load_qty               = Convert $header['load_qty']
            load_qty_erp           = Convert $header['load_qty_erp']
            load_qty_prod          = Convert $header['load_qty_prod']
            reuse_qty              = Convert $header['reuse_qty']
            batch_count            = Convert $header['batch_count']
            load_end_date          = ConvertTo-DateTimeHC $header['load_end_date']
            ticket_time            = ConvertTo-DateTimeHC $header['ticket_time']
            load_start_date        = $loadStartDate
        }
        #endregion

        foreach ($batch in $delivery.batches.batch) {
            #region Get batchHeader
            <#
                Built once per batch and read again by every batch item of that
                batch, so it stays a cache here too.
            #>
            $batchHeaderNode = $batch.batchHeader

            $header = Get-ChildValueMapHC -Node $batchHeaderNode

            $batchHeader = [PSCustomObject]@{
                qty_unit                  = $header['qty_unit']
                water_trim_unit           = $header['water_trim_unit']
                sequence                  = $header['sequence']
                mixing_time_unit          = $header['mixing_time_unit']
                mixer_power_unit          = $header['mixer_power_unit']
                slump_unit                = $header['slump_unit']
                mixer_discharge_time_unit = $header['mixer_discharge_time_unit']
                <#
                    A container of other elements, so it is not in the map and
                    is still read from the node. It is iterated further down.
                #>
                dischargingOperationTimes = $batchHeaderNode.dischargingOperationTimes
                <#
                    'batch_id' identifies the batch, so it is text.
                    'batch_id_nr' is the counter of the batch within its
                    delivery, so it stays a number.
                #>
                batch_id                  = Text $header['batch_id']
                batch_id_nr               = Convert $header['batch_id_nr']
                qty                       = Convert $header['qty']
                water_trim                = Convert $header['water_trim']
                manual                    = Convert $header['manual']
                mixing_time               = Convert $header['mixing_time']
                mixer_power               = Convert $header['mixer_power']
                slump                     = Convert $header['slump']
                mixer_discharge_time      = Convert $header['mixer_discharge_time']
                start_time                = ConvertTo-DateTimeHC $header['start_time']
                end_time                  = ConvertTo-DateTimeHC $header['end_time']
                aborted                   = Convert (
                    Get-ChildValueMapHC -Node $batchHeaderNode.batchinformation
                )['aborted']
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
                $discharging = Get-ChildValueMapHC -Node $dischargingOperation

                @{
                    Key   = 'discharging'
                    Cells = @{
                        A = $FileName
                        B = $PlantHeader.plant_code
                        C = $PlantHeader.plant_name
                        D = Text $discharging['batch_id']
                        E = Convert $discharging['batch_id_nr']
                        F = Text $discharging['scale_id']
                        G = $discharging['material_type']
                        H = ConvertTo-DateTimeHC $discharging['discharge_start_time']
                        I = ConvertTo-DateTimeHC $discharging['discharge_end_time']
                    }
                }
            }
            #endregion

            #region Add rows to worksheet batchItems
            foreach ($batchItem in $batch.batchItems.batchItem) {
                <#
                    The hottest node in the whole script: about thirty values
                    are read from every batch item, and there is one batch item
                    per material per batch.

                    'batchiteminformation' and 'dosingOperationTimes' hold
                    other elements, so they are not in the map. Their nodes are
                    resolved once here instead of once per value read.
                #>
                $item = Get-ChildValueMapHC -Node $batchItem

                $itemInformation = Get-ChildValueMapHC -Node $batchItem.batchiteminformation

                $dosing = Get-ChildValueMapHC -Node $batchItem.dosingOperationTimes.dosing

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
                        AH = Convert $itemInformation['pulsecount']
                        AI = $itemInformation['moisture_measure_type']
                        AJ = Text $item['material_code']
                        AK = $item['material_name']
                        AL = $item['vendor']
                        AM = $item['vendor_source']
                        AN = $item['vendor_delivery']
                        AO = $item['material_type']
                        AP = Convert $item['material_target_dry']
                        AQ = $item['material_target_dry_unit']
                        AR = Convert $item['material_target']
                        AS = $item['material_target_unit']
                        AT = Convert $item['material_target_adjusted']
                        AU = $item['material_target_adjusted_unit']
                        AV = Convert $item['material_amount']
                        AW = $item['material_amount_unit']
                        AX = Convert $item['material_moisture']
                        AY = $item['material_moisture_unit']
                        AZ = Convert $item['material_moisture_auto']
                        BA = Convert $item['material_density']
                        BB = $item['material_density_unit']
                        BC = Convert $item['material_absorption']
                        BD = $item['material_absorption_unit']
                        BE = Convert $item['material_solid_content']
                        BF = $item['material_solid_content_unit']
                        BG = Convert $item['material_temperature']
                        BH = $item['material_temperature_unit']
                        BI = Convert $item['material_bin_number']
                        BJ = $item['material_bin_name']
                        BK = Text $item['scale_id']
                        BL = $dosing['material_dosing_start_time']
                        BM = $dosing['material_dosing_end_time']
                    }
                }
            }
            #endregion
        }
    }
}

function Get-AlarmRowHC {
    <#
        .SYNOPSIS
            Rows for the alarms of one batch computer that were raised in the
            requested month.

        .DESCRIPTION
            The month is decided per alarm on 'raised'. An alarm without a
            raised date cannot be placed in a month and is reported as an error
            by 'Get XML file dates.ps1' instead of being written to an
            arbitrary Excel file.
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
    Set-Alias -Name 'Text' -Value ConvertTo-TextHC

    <#
        Worked out once instead of formatting the date of every record into a
        month key and comparing two strings. 'End' is the first moment of the
        next month, so the test below is 'at or after Start, before End'.
    #>
    $monthRange = Get-MonthRangeHC -MonthKey $MonthKey

    foreach ($alarm in $BatchComputer.alarms.alarm) {
        <#
            'ticket' and 'batch' hold other elements, so they are not in the
            map of the alarm itself and get a map each. 'alarmGroups' can hold
            several 'alarmGroup' elements whose values are joined below, which
            is a collection and not a single value, so it stays on the node.
        #>
        $alarmMap = Get-ChildValueMapHC -Node $alarm

        #region Only alarms raised in the requested month
        $alarmRaised = ConvertTo-DateTimeHC $alarmMap['raised']

        if (-not $alarmRaised) { Continue }

        if (
            ($alarmRaised -lt $monthRange.Start) -or
            ($alarmRaised -ge $monthRange.End)
        ) { Continue }
        #endregion

        $ticket = Get-ChildValueMapHC -Node $alarm.ticket

        $batch = Get-ChildValueMapHC -Node $alarm.batch

        #region Add row to worksheet alarms
        @{
            Key   = 'alarms'
            Cells = @{
                A  = $FileName
                B  = $PlantHeader.plant_code
                C  = $PlantHeader.plant_name
                D  = $null
                E  = $BatchComputerHeader.extraction_id
                F  = Text $alarmMap['Id']
                G  = $alarmRaised
                H  = ConvertTo-DateTimeHC $alarmMap['dropped']
                I  = ConvertTo-DateTimeHC $alarmMap['handled']
                J  = ConvertTo-DateTimeHC $alarmMap['clicktimestamp']
                <#
                    Alarm and ticket text, kept as text. An alarm message that
                    happens to hold nothing but digits is still a message, and
                    'Truck' is a license plate.
                #>
                K  = Text $alarmMap['text']
                L  = Text $alarmMap['alarm_message']
                M  = Text $alarmMap['message']
                N  = Text $alarmMap['parameter']
                O  = Text $ticket['Id']
                P  = Text $ticket['Code']
                Q  = Text $ticket['Reference']
                R  = Text $ticket['Customer']
                S  = Text $ticket['Project']
                T  = Convert $ticket['Wanted']
                U  = Text $ticket['Truck']
                V  = Text $batch['Id']
                W  = ConvertTo-DateTimeHC $batch['TimeStart']
                X  = ConvertTo-DateTimeHC $batch['TimeReady']
                Y  = Convert $batch['Wanted']
                Z  = Text $batch['WeighListCode']
                AA = Text $batch['WeighListName']
                AB = if ($alarm.alarmGroups.alarmGroup.Code) {
                    $alarm.alarmGroups.alarmGroup.Code -join ','
                }
                AC = if ($alarm.alarmGroups.alarmGroup.Name) {
                    $alarm.alarmGroups.alarmGroup.Name -join ','
                }
                AD = if ($alarm.alarmGroups.alarmGroup.Description) {
                    $alarm.alarmGroups.alarmGroup.Description -join ','
                }
            }
        }
        #endregion
    }
}

function Get-SequenceRowHC {
    <#
        .SYNOPSIS
            Rows for the sequence parameters of one batch computer, when that
            batch computer was created in the requested month.

        .DESCRIPTION
            The month is decided per batch computer on 'file_created_on'.
            A file holding two batch computers created in two different months
            is written to two Excel files.
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
    Set-Alias -Name 'Text' -Value ConvertTo-TextHC

    #region Only batch computers created in the requested month
    $fileCreatedOn = $BatchComputerHeader.file_created_on

    if (-not $fileCreatedOn) { return }

    <#
        Unlike the batch and alarm builders this runs once per batch computer,
        not once per record, so there is nothing to hoist. It uses the same
        range test to keep all three builders deciding the month the same way.
    #>
    $monthRange = Get-MonthRangeHC -MonthKey $MonthKey

    if (
        ($fileCreatedOn -lt $monthRange.Start) -or
        ($fileCreatedOn -ge $monthRange.End)
    ) { return }
    #endregion

    foreach (
        $sequenceParameter in
        $BatchComputer.sequenceParameters.sequenceParameter
    ) {
        #region Get sequenceParameterHeader
        <#
            Read again by every property of every sub batch and every parameter
            below, so it stays a cache.
        #>
        $sequenceParameterMap = Get-ChildValueMapHC -Node $sequenceParameter

        $sequenceHeader = [PSCustomObject]@{
            name             = $sequenceParameterMap['name']
            baseName         = $sequenceParameterMap['baseName']
            maxBatchSizeUnit = $sequenceParameterMap['maxBatchSizeUnit']
            ID               = Text $sequenceParameterMap['ID']
            NrofSubBatches   = Convert $sequenceParameterMap['NrofSubBatches']
            maxbatchsize     = Convert $sequenceParameterMap['maxbatchsize']
            blocked          = Convert $sequenceParameterMap['blocked']
        }
        #endregion

        foreach ($subBatch in $sequenceParameter.subBatches.subBatch) {
            $subBatchName = Text $subBatch.name

            foreach ($property in $subBatch.properties.property) {
                <#
                    The map is case insensitive, which matters here: the XML
                    holds 'itempath' and 'uppath' while the cells below ask for
                    'itemPath' and 'upPath'.
                #>
                $propertyMap = Get-ChildValueMapHC -Node $property

                #region Add row to worksheet sequences
                @{
                    Key   = 'sequence'
                    Cells = @{
                        A = $FileName
                        B = $PlantHeader.plant_code
                        C = $PlantHeader.plant_name
                        D = $null
                        E = $BatchComputerHeader.extraction_id
                        F = $fileCreatedOn
                        G = $sequenceHeader.ID
                        H = $sequenceHeader.name
                        I = $sequenceHeader.baseName
                        J = $sequenceHeader.NrofSubBatches
                        K = $sequenceHeader.maxbatchsize
                        L = $sequenceHeader.maxBatchSizeUnit
                        M = $sequenceHeader.blocked
                        N = $subBatchName
                        O = $propertyMap['name']
                        P = $propertyMap['prefix']
                        Q = $propertyMap['itemPath']
                        R = $propertyMap['upPath']
                        S = Convert $propertyMap['value']
                        T = $propertyMap['suffix']
                    }
                }
                #endregion
            }
        }

        foreach ($parameter in $sequenceParameter.parameters.parameter) {
            foreach ($property in $parameter.properties.property) {
                $propertyMap = Get-ChildValueMapHC -Node $property

                #region Add row to worksheet sequences
                @{
                    Key   = 'sequence'
                    Cells = @{
                        A = $FileName
                        B = $PlantHeader.plant_code
                        C = $PlantHeader.plant_name
                        D = $null
                        E = $BatchComputerHeader.extraction_id
                        F = $fileCreatedOn
                        G = $sequenceHeader.ID
                        H = $sequenceHeader.name
                        I = $sequenceHeader.baseName
                        J = $sequenceHeader.NrofSubBatches
                        K = $sequenceHeader.maxbatchsize
                        L = $sequenceHeader.maxBatchSizeUnit
                        M = $sequenceHeader.blocked
                        N = $null
                        O = $propertyMap['name']
                        P = $propertyMap['prefix']
                        Q = $propertyMap['itemPath']
                        R = $propertyMap['upPath']
                        S = Convert $propertyMap['value']
                        T = $propertyMap['suffix']
                    }
                }
                #endregion
            }
        }
    }
}