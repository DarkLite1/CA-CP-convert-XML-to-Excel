#Requires -Version 7

function Format-WorksheetHeaderHC {
    <#
        .SYNOPSIS
            Apply the background color belonging to a header type.

        .DESCRIPTION
            The first header row of each worksheet groups the columns by where
            they come from (plant, batch computer, delivery, ...). Each group
            has its own color so the sheet is easy to read.

        .EXAMPLE
            Format-WorksheetHeaderHC -Worksheet $ws -Cell 'C1' -Type 'plantHeader'
    #>
    param (
        [Parameter(Mandatory)]
        $Worksheet,
        [Parameter(Mandatory)]
        [String]$Cell,
        [Parameter(Mandatory)]
        [ValidateSet(
            'plantHeader', 'batchComputerHeader', 'deliveryHeader',
            'batchHeader', 'batchHeaderDischargeOptions', 'batchItem',
            'alarm', 'sequenceParameter', 'subBatch',
            'parameterAndSubBatchProperty'
        )]
        [String]$Type
    )

    $colorMap = @{
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

    try {
        $color = $colorMap[$Type]

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
