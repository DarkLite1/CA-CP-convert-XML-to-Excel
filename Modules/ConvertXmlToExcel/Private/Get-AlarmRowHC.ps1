#Requires -Version 7

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

    foreach ($alarm in $BatchComputer.alarms.alarm) {
        #region Only alarms raised in the requested month
        $alarmRaised = ConvertTo-DateTimeHC $alarm.raised

        if (-not $alarmRaised) { Continue }

        if ((Get-MonthKeyHC -Date $alarmRaised) -ne $MonthKey) { Continue }
        #endregion

        #region Add row to worksheet alarms
        @{
            Key   = 'alarms'
            Cells = @{
                A  = $FileName
                B  = $PlantHeader.plant_code
                C  = $PlantHeader.plant_name
                D  = $null
                E  = $BatchComputerHeader.extraction_id
                F  = Convert $alarm.Id
                G  = $alarmRaised
                H  = ConvertTo-DateTimeHC $alarm.dropped
                I  = ConvertTo-DateTimeHC $alarm.handled
                J  = ConvertTo-DateTimeHC $alarm.clicktimestamp
                K  = Convert $alarm.text
                L  = Convert $alarm.alarm_message
                M  = Convert $alarm.message
                N  = Convert $alarm.parameter
                O  = Convert $alarm.ticket.Id
                P  = Convert $alarm.ticket.Code
                Q  = Convert $alarm.ticket.Reference
                R  = Convert $alarm.ticket.Customer
                S  = Convert $alarm.ticket.Project
                T  = Convert $alarm.ticket.Wanted
                U  = Convert $alarm.ticket.Truck
                V  = Convert $alarm.batch.Id
                W  = ConvertTo-DateTimeHC $alarm.batch.TimeStart
                X  = ConvertTo-DateTimeHC $alarm.batch.TimeReady
                Y  = Convert $alarm.batch.Wanted
                Z  = Convert $alarm.batch.WeighListCode
                AA = Convert $alarm.batch.WeighListName
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
