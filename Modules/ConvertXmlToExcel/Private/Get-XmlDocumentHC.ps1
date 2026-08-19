#Requires -Version 7

function Get-XmlDocumentHC {
    <#
        .SYNOPSIS
            Read an XML file with the external entity resolver disabled.

        .DESCRIPTION
            Disabling the resolver stops the parser from reaching out to
            external systems for DTDs or entities referenced in the file.

        .EXAMPLE
            Get-XmlDocumentHC -Path 'C:\Data\File.xml'
    #>
    param (
        [Parameter(Mandatory)]
        [String]$Path
    )

    $xmlDocument = [xml]''
    $xmlDocument.XmlResolver = $null
    $xmlDocument.Load($Path)

    $xmlDocument
}
