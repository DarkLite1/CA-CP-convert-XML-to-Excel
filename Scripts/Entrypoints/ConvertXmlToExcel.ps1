#Requires -Version 7

<#
    .SYNOPSIS
        Convert XML files to Excel files.

    .DESCRIPTION
        Thin entrypoint. Loads the ConvertXmlToExcel module and invokes the
        public orchestrator 'Invoke-ConvertXmlToExcel'. All the work lives in
        the module; this script only resolves paths, imports the module and
        starts the run so it can be called from a scheduled task.

        ImportExcel is the only external module dependency. Event logging, log
        files and mail are handled by the module's own helper functions.

    .PARAMETER ConfigurationJsonFile
        The .json configuration file, based on a file in 'Examples'.

    .EXAMPLE
        & '.\Scripts\Entrypoints\ConvertXmlToExcel.ps1' -ConfigurationJsonFile 'C:\Batch.json'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [String]$ConfigurationJsonFile
)

begin {
    try {
        # Failures before the orchestrator can run, for example a module that
        # cannot be loaded. There is no mail or event log yet at that point.
        $bootstrapErrors = [System.Collections.Generic.List[object]]::new()

        # Failures raised by the orchestrator itself. Passed by reference so
        # every stage can add to the same list.
        $runtimeErrors = [System.Collections.Generic.List[object]]::new()

        $projectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

        $modulePath = Join-Path $projectRoot 'Modules\ConvertXmlToExcel\ConvertXmlToExcel.psm1'
        $opsRoot = Join-Path $projectRoot 'Scripts\Operations'

        $ScriptPath = @{
            GetXmlFileDates = Join-Path $opsRoot 'GetXmlFileDates.ps1'
            ExportToExcel   = Join-Path $opsRoot 'ExportXmlFileToExcel.ps1'
        }

        #region Import the module
        if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
            $bootstrapErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Type     = 'FatalError'
                    Name     = 'Module not found'
                    Message  = "ConvertXmlToExcel module not found at '$modulePath'."
                    Category = 'Bootstrap'
                }
            )
        }
        else {
            try {
                if (Get-Module -Name 'ConvertXmlToExcel') {
                    Write-Verbose 'ConvertXmlToExcel module already loaded'
                }
                else {
                    Write-Verbose "Importing ConvertXmlToExcel module: $modulePath"
                    Import-Module $modulePath -Force -ErrorAction Stop
                }
            }
            catch {
                $bootstrapErrors.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Type     = 'FatalError'
                        Name     = 'Module import failed'
                        Message  = "Failed to import the ConvertXmlToExcel module: $_"
                        Category = 'Bootstrap'
                    }
                )
            }
        }
        #endregion
    }
    catch {
        Write-Warning "BEGIN stage crashed before the orchestrator could run: $_"
        exit 1
    }
}

process { }

end {
    #region Bootstrap failed, the orchestrator is not available
    if ($bootstrapErrors.Count -gt 0) {
        foreach ($err in $bootstrapErrors) {
            Write-Warning "[$($err.Category)] $($err.Message)"
        }
        exit 1
    }
    #endregion

    try {
        Invoke-ConvertXmlToExcel `
            -ConfigurationJsonFile $ConfigurationJsonFile `
            -ScriptPath $ScriptPath `
            -ModulePath $modulePath `
            -SystemErrors ([ref]$runtimeErrors)
    }
    catch {
        Write-Warning "Unhandled fatal error: $_"
        exit 1
    }

    #region Report the collected errors and set the exit code
    if ($runtimeErrors.Count -gt 0) {
        foreach ($err in $runtimeErrors) {
            Write-Warning "[$($err.Type)] $($err.Name): $($err.Message)"
        }

        if ($runtimeErrors.where({ $_.Type -eq 'FatalError' })) {
            exit 1
        }
    }
    #endregion
}