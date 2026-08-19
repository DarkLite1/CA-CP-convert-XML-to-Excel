#Requires -Version 7

<#
    .SYNOPSIS
        Convert XML files to Excel files.

    .DESCRIPTION
        Thin entrypoint. Loads the ConvertXmlToExcel module and invokes the
        public orchestrator 'Invoke-ConvertXmlToExcel'. All the work lives in
        the module; this script only resolves paths and starts the run so it can
        be called from a scheduled task.

    .PARAMETER ScriptName
        Name used for the event log source and the log file names.

    .PARAMETER ImportFile
        The .json configuration file, based on one of the files in 'Examples'.

    .EXAMPLE
        & '.\Scripts\Entrypoints\ConvertXmlToExcel.ps1' `
            -ScriptName 'CP Batch' -ImportFile 'C:\Batch.json'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile
)

begin {
    try {
        $projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

        $modulePath = Join-Path $projectRoot 'Modules\ConvertXmlToExcel\ConvertXmlToExcel.psm1'
        $opsRoot = Join-Path $projectRoot 'Scripts\Operations'

        $ScriptPath = @{
            GetXmlFileDates = Join-Path $opsRoot 'GetXmlFileDates.ps1'
            ExportToExcel   = Join-Path $opsRoot 'ExportXmlFileToExcel.ps1'
        }

        #region Import module
        if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
            Write-Warning "ConvertXmlToExcel module not found at '$modulePath'."
            exit 1
        }

        if (Get-Module -Name 'ConvertXmlToExcel') {
            Write-Verbose 'ConvertXmlToExcel module already loaded'
        }
        else {
            Write-Verbose "Importing ConvertXmlToExcel module: $modulePath"
            Import-Module $modulePath -Force -ErrorAction Stop
        }
        #endregion
    }
    catch {
        Write-Warning "BEGIN stage crashed before orchestrator could run: $_"
        exit 1
    }
}

process { }

end {
    try {
        Invoke-ConvertXmlToExcel `
            -ScriptName $ScriptName `
            -ImportFile $ImportFile `
            -ScriptPath $ScriptPath `
            -ModulePath $modulePath
    }
    catch {
        Write-Warning "Unhandled fatal error: $_"
        exit 1
    }
}
