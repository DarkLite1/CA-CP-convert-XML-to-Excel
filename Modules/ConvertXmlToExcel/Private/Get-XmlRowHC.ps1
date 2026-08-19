#Requires -Version 7

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

    #region Get plantHeader
    $plantHeader = $Xml.plant.plantHeader |
    Select-Object -Property country_code, company_code, company_name,
    plant_code, plant_name
    #endregion

    foreach ($batchComputer in $Xml.plant.batchComputers.batchComputer) {
        #region Get batchComputerHeader
        $batchComputerHeader = $batchComputer.batchComputerHeader |
        Select-Object -Property system_type, system_provider, mixer_name,
        offset,
        @{
            Name       = 'system_id'
            Expression = { Convert $_.system_id }
        },
        @{
            Name       = 'mixer_size'
            Expression = { Convert $_.mixer_size }
        },
        @{
            Name       = 'extraction_id'
            Expression = { Convert $_.extraction_id }
        },
        @{
            Name       = 'file_created_on'
            Expression = { ConvertTo-DateTimeHC $_.file_created_on }
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
