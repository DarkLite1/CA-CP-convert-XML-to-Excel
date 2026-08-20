#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\WorksheetDefinition.ps1"

    # A tiny fake worksheet whose cell exposes the Style shape the function sets
    function New-FakeWorksheet {
        [PSCustomObject]@{
            Cells = @{
                'C1' = [PSCustomObject]@{
                    Style = [PSCustomObject]@{
                        Fill = [PSCustomObject]@{
                            PatternType     = $null
                            BackgroundColor = [PSCustomObject]@{
                                Color = $null
                            }
                        }
                    }
                }
            }
        } | Add-Member -PassThru -Force -MemberType ScriptMethod -Name Dummy -Value {}
    }
}

Describe 'Format-WorksheetHeaderHC' {
    BeforeAll {
        # SetColor is added per test cell so we can capture what was requested
        function Add-SetColor {
            param($Worksheet)
            $bg = $Worksheet.Cells['C1'].Style.Fill.BackgroundColor
            $bg | Add-Member -Force -MemberType ScriptMethod -Name SetColor -Value {
                param($Color) $this.Color = $Color
            }
        }
    }

    It 'sets a pattern and a known color for a valid header type' {
        $ws = New-FakeWorksheet
        Add-SetColor -Worksheet $ws

        Format-WorksheetHeaderHC -Worksheet $ws -Cell 'C1' -Type 'plantHeader'

        $ws.Cells['C1'].Style.Fill.PatternType | Should-Be 'MediumGray'
        $ws.Cells['C1'].Style.Fill.BackgroundColor.Color | Should-BeTruthy
    }

    It 'rejects an unsupported header type via parameter validation' {
        $ws = New-FakeWorksheet
        { Format-WorksheetHeaderHC -Worksheet $ws -Cell 'C1' -Type 'nope' } |
        Should-Throw
    }
}
