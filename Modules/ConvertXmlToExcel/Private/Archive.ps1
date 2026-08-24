#Requires -Version 7

<#
    .SYNOPSIS
        Move a converted XML file to the archive folder.

    .DESCRIPTION
        An XML file keeps its name when it is archived, so a file that is
        delivered a second time under the same name collides with the copy that
        is already there. What to do then depends on the content, not on the
        name.
#>

function Move-ToArchiveFolderHC {
    <#
        .SYNOPSIS
            Move one XML file to the archive folder.

        .DESCRIPTION
            When no file of that name is in the archive yet the file is simply
            moved.

            When a file of that name is already there the content decides:

            - Same content: the file was delivered twice and nothing was lost,
              because the data of the first copy is already in the Excel file.
              It is archived under a name carrying the 'Duplicate_<timestamp>_'
              prefix, so the archive keeps every delivery without ever
              overwriting an older one.

            - Different content: this is not a duplicate but a different file
              that happens to share a name. Its records were NOT added to the
              Excel file, because an XML file is recognized by its name and that
              name was already present. Overwriting the archived copy would
              destroy the only proof of what changed, so the file is left where
              it is for a manual action.

        .PARAMETER File
            The XML file to archive.

        .PARAMETER ArchiveFolder
            The folder to move the file to. Created when missing.

        .PARAMETER ScriptStartTime
            Start time of the run, used in the prefix of a duplicate, so every
            duplicate of the same run carries the same stamp.

        .OUTPUTS
            A hashtable with:
            - Archived   : whether the file was moved
            - ArchivedAs : the name it received, when it was moved
            - IsDuplicate: whether it was archived as a duplicate
            - Reason     : why it was not moved, when it was not

        .EXAMPLE
            Move-ToArchiveFolderHC -File $file -ArchiveFolder 'C:\Xml\Archive' `
                -ScriptStartTime (Get-Date)
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$File,
        [Parameter(Mandatory)]
        [String]$ArchiveFolder,
        [Parameter(Mandatory)]
        [DateTime]$ScriptStartTime
    )

    $result = @{
        Archived    = $false
        ArchivedAs  = $null
        IsDuplicate = $false
        Reason      = $null
    }

    try {
        #region The file can be gone when a previous run moved it
        if (-not (Test-Path -LiteralPath $File.FullName -PathType Leaf)) {
            $result.Reason = 'The XML file is no longer in the folder'
            return $result
        }
        #endregion

        $null = New-Item -Path $ArchiveFolder -ItemType 'Directory' -Force -EA Stop

        $destination = Join-Path $ArchiveFolder $File.Name

        #region No collision, simply move the file
        if (-not (Test-Path -LiteralPath $destination -PathType Leaf)) {
            Move-Item -LiteralPath $File.FullName -Destination $destination -EA Stop

            $result.Archived = $true
            $result.ArchivedAs = $File.Name

            return $result
        }
        #endregion

        #region Compare the content with the file already in the archive
        <#
            The length is checked first because it is free and rules out most
            differences without reading both files.
        #>
        $archivedFile = Get-Item -LiteralPath $destination -EA Stop

        $isSameContent = $false

        if ($archivedFile.Length -eq $File.Length) {
            $isSameContent = (Get-FileHash -LiteralPath $File.FullName -Algorithm 'SHA256' -EA Stop).Hash -eq
            (Get-FileHash -LiteralPath $destination -Algorithm 'SHA256' -EA Stop).Hash
        }
        #endregion

        #region Different content, leave the file for a manual action
        if (-not $isSameContent) {
            $result.Reason = (
                "An XML file named '{0}' is already in the archive folder but its content is different. " +
                'The records of this file were NOT added to the Excel file, because an XML file is ' +
                'recognized by its name and that name was already present. Compare both files and ' +
                'rename or remove this one manually.'
            ) -f $File.Name

            return $result
        }
        #endregion

        #region Same content, archive it as a duplicate
        $duplicateName = 'Duplicate_{0:yyyyMMdd_HHmmss}_{1}' -f $ScriptStartTime, $File.Name

        $duplicateDestination = Join-Path $ArchiveFolder $duplicateName

        <#
            Two duplicates of the same file in one run would collide again, so a
            counter is added until the name is free.
        #>
        $counter = 1

        while (Test-Path -LiteralPath $duplicateDestination -PathType Leaf) {
            $counter++

            $duplicateName = 'Duplicate_{0:yyyyMMdd_HHmmss}_{1}_{2}' -f
            $ScriptStartTime, $counter, $File.Name

            $duplicateDestination = Join-Path $ArchiveFolder $duplicateName
        }

        Move-Item -LiteralPath $File.FullName -Destination $duplicateDestination -EA Stop

        $result.Archived = $true
        $result.ArchivedAs = $duplicateName
        $result.IsDuplicate = $true
        #endregion
    }
    catch {
        $result.Archived = $false
        $result.Reason = "Failed moving the file to the archive folder '$ArchiveFolder': $_"

        if ($Error.Count) { $Error.RemoveAt(0) }
    }

    $result
}