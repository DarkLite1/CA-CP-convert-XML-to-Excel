@{
    RootModule           = 'ConvertXmlToExcel.psm1'
    ModuleVersion        = '1.0.0'
    GUID                 = '73ec6e3e-b195-4266-ab7b-5258c2bf46a4'
    CompatiblePSEditions = @('Core')

    Author               = 'DarkLite1'
    CompanyName          = ''
    Copyright            = '(c) No warranty. Provided as-is.'

    Description          = @'
Convert TMS CP batch, alarm and sequence XML files to Excel files, one Excel
file per month. A single XML file that holds records of more than one month is
written to more than one Excel file, each getting only the records of its own
month.
'@

    PowerShellVersion    = '7.0'

    RequiredModules      = @(
        @{
            ModuleName    = 'ImportExcel'
            ModuleVersion = '7.8.5'
        }
    )

    FunctionsToExport    = @(
        'Invoke-ConvertXmlToExcel'
    )

    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()

    PrivateData          = @{
        PSData = @{
            Tags         = @('TMS', 'CP', 'XML', 'Excel', 'Batch', 'Alarm', 'Sequence')
            ProjectUri   = 'https://github.com/DarkLite1/CA-CP-convert-XML-file-to-Excel'
            ReleaseNotes = 'Merge of the batch, alarm and sequence scripts. One XML file can now be written to more than one monthly Excel file.'
        }
    }
}
