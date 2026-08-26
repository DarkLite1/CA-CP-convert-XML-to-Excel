#Requires -Version 7

function Get-XmlFileMonthHC {
    <#
        .SYNOPSIS
            Return every month present in an XML file.

        .DESCRIPTION
            One XML file can hold records for more than one month. For example
            a batch file with one delivery loaded today and another delivery
            loaded tomorrow, where tomorrow falls in the next month.

            This function returns one entry per month found in the file, so the
            caller can write each part of the file to the matching Excel file.

            The date used depends on the type:

            Batch    : delivery.deliveryHeader.load_start_date
            Alarm    : alarm.raised
            Sequence : batchComputer.batchComputerHeader.file_created_on

            Records without a usable date are counted in 'RecordsWithoutDate'.
            They are never silently written to an arbitrary Excel file.

            The file is read with an XmlReader instead of being loaded into an
            XmlDocument. Only one element per record is needed here, so there
            is no reason to build a full document tree in memory first. A
            reader walks the file forward once, holds a single node at a time
            and never allocates the tree, which matters for the large files:
            the old version loaded the whole document, and the export step then
            loaded the very same file again for every month it found.

            Reading through PowerShell's dotted property access, as in
            '$Xml.plant.batchComputers.batchComputer.deliveries.delivery', was
            the other cost. Each step goes through the XmlNode adapter, which
            builds a member table per node, and the chain materialized every
            delivery in the file just to reach one child element.

        .PARAMETER Path
            Full path to the XML file.

        .PARAMETER Type
            'Batch', 'Alarm' or 'Sequence'.

        .EXAMPLE
            Get-XmlFileMonthHC -Path 'C:\Data\File.xml' -Type 'Batch'
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$Path,
        [Parameter(Mandatory)]
        [ValidateSet('Batch', 'Alarm', 'Sequence')]
        [String]$Type
    )

    #region The element holding the date, and the element it sits in
    <#
        The parent is checked as well so that an element that happens to carry
        the same name somewhere else in the file cannot be mistaken for a
        record date. The dotted property access this replaces was anchored to a
        fixed path and had that guarantee for free.
    #>
    switch ($Type) {
        'Batch' {
            $elementName = 'load_start_date'
            $parentName = 'deliveryHeader'
            break
        }
        'Alarm' {
            $elementName = 'raised'
            $parentName = 'alarm'
            break
        }
        'Sequence' {
            $elementName = 'file_created_on'
            $parentName = 'batchComputerHeader'
            break
        }
    }
    #endregion

    #region Reader settings
    <#
        The resolver is disabled and DTDs are refused for the same reason as in
        Get-XmlDocumentHC: the parser must not reach out to external systems
        for anything referenced inside the file.
    #>
    $readerSettings = [System.Xml.XmlReaderSettings]::new()
    $readerSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $readerSettings.XmlResolver = $null
    $readerSettings.IgnoreWhitespace = $true
    $readerSettings.IgnoreComments = $true
    $readerSettings.IgnoreProcessingInstructions = $true
    #endregion

    $monthKeys = [System.Collections.Generic.HashSet[String]]::new()
    $recordCount = 0
    $recordsWithoutDate = 0

    <#
        The name of the element open at each depth. An element start is always
        seen before its children, so when the date element is reached the entry
        one level up holds the name of the element it sits in.
    #>
    $elementNameByDepth = [String[]]::new(256)

    $reader = [System.Xml.XmlReader]::Create($Path, $readerSettings)

    try {
        while (-not $reader.EOF) {
            if ($reader.NodeType -ne [System.Xml.XmlNodeType]::Element) {
                $null = $reader.Read()
                continue
            }

            $depth = $reader.Depth

            if ($depth -lt $elementNameByDepth.Length) {
                $elementNameByDepth[$depth] = $reader.LocalName
            }

            if ($reader.LocalName -ne $elementName) {
                $null = $reader.Read()
                continue
            }

            if (
                ($depth -eq 0) -or
                ($elementNameByDepth[$depth - 1] -cne $parentName)
            ) {
                $null = $reader.Read()
                continue
            }

            $recordCount++

            <#
                Reads the text and moves the reader past the closing tag on its
                own, so the loop must not call Read() after it. Doing both
                would step over the node that follows.
            #>
            $rawDate = $reader.ReadElementContentAsString()

            $date = ConvertTo-DateTimeHC -Value $rawDate

            if ($date) {
                $null = $monthKeys.Add((Get-MonthKeyHC -Date $date))
            }
            else {
                $recordsWithoutDate++
            }
        }
    }
    catch {
        throw "Failed reading XML file '$Path': $_"
    }
    finally {
        $reader.Dispose()
    }

    [PSCustomObject]@{
        MonthKeys          = @($monthKeys | Sort-Object)
        RecordCount        = $recordCount
        RecordsWithoutDate = $recordsWithoutDate
    }
}