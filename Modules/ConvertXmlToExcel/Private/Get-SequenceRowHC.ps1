#Requires -Version 7

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

    #region Only batch computers created in the requested month
    $fileCreatedOn = $BatchComputerHeader.file_created_on

    if (-not $fileCreatedOn) { return }

    if ((Get-MonthKeyHC -Date $fileCreatedOn) -ne $MonthKey) { return }
    #endregion

    foreach (
        $sequenceParameter in
        $BatchComputer.sequenceParameters.sequenceParameter
    ) {
        #region Get sequenceParameterHeader
        $sequenceHeader = $sequenceParameter |
        Select-Object -Property name, baseName, maxBatchSizeUnit,
        @{
            Name       = 'ID'
            Expression = { Convert $_.ID }
        },
        @{
            Name       = 'NrofSubBatches'
            Expression = { Convert $_.NrofSubBatches }
        },
        @{
            Name       = 'maxbatchsize'
            Expression = { Convert $_.maxbatchsize }
        },
        @{
            Name       = 'blocked'
            Expression = { Convert $_.blocked }
        }
        #endregion

        foreach ($subBatch in $sequenceParameter.subBatches.subBatch) {
            $subBatchName = Convert $subBatch.name

            foreach ($property in $subBatch.properties.property) {
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
                        O = $property.name
                        P = $property.prefix
                        Q = $property.itemPath
                        R = $property.upPath
                        S = Convert $property.value
                        T = $property.suffix
                    }
                }
                #endregion
            }
        }

        foreach ($parameter in $sequenceParameter.parameters.parameter) {
            foreach ($property in $parameter.properties.property) {
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
                        O = $property.name
                        P = $property.prefix
                        Q = $property.itemPath
                        R = $property.upPath
                        S = Convert $property.value
                        T = $property.suffix
                    }
                }
                #endregion
            }
        }
    }
}
