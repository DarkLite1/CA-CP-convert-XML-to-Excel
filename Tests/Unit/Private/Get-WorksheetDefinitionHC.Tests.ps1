#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Get-WorksheetDefinitionHC.ps1"
}

Describe 'Get-WorksheetDefinitionHC' {
    Context 'the plant worksheet is always first' {
        It 'names the first worksheet plantBatchComputers for every type' {
            foreach ($type in 'Batch', 'Alarm', 'Sequence') {
                $definitions = @(Get-WorksheetDefinitionHC -Type $type)
                $definitions[0].Key | Should-Be 'plant'
                $definitions[0].Name | Should-Be 'plantBatchComputers'
            }
        }
    }

    Context 'worksheet names match the files made by the previous version' {
        It 'keeps the batch worksheet names' {
            $names = (Get-WorksheetDefinitionHC -Type 'Batch').Name
            $names | Should-ContainCollection 'deliveriesBatches'
            $names | Should-ContainCollection 'batchesDischarging'
            $names | Should-ContainCollection 'batchItems'
        }

        It 'keeps the alarm worksheet name' {
            (Get-WorksheetDefinitionHC -Type 'Alarm').Name | Should-ContainCollection 'alarms'
        }

        It 'keeps the sequence worksheet name' {
            (Get-WorksheetDefinitionHC -Type 'Sequence').Name | Should-ContainCollection 'sequences'
        }
    }

    Context 'each worksheet carries a table name so it can be formatted' {
        It 'gives every worksheet of every type a TableName' {
            foreach ($type in 'Batch', 'Alarm', 'Sequence') {
                foreach ($definition in Get-WorksheetDefinitionHC -Type $type) {
                    $definition.TableName | Should-BeTruthy
                }
            }
        }
    }

    It 'rejects an unknown type' {
        { Get-WorksheetDefinitionHC -Type 'Nope' } | Should-Throw
    }
}
