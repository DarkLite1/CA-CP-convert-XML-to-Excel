#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Archive.ps1"

    $script:startTime = [datetime]'2026-08-24 09:57:48'
    $script:stamp = '20260824_095748'

    function New-XmlFileHC {
        param ([String]$Folder, [String]$Name = 'Plant.xml', [String]$Content = '<plant/>')

        $null = New-Item -Path $Folder -ItemType 'Directory' -Force
        $path = Join-Path $Folder $Name
        Set-Content -LiteralPath $path -Value $Content -NoNewline

        Get-Item -LiteralPath $path
    }
}

Describe 'Move-ToArchiveFolderHC' {
    BeforeEach {
        $testRoot = Join-Path $TestDrive "case_$(New-Guid)"
        $inputFolder = Join-Path $testRoot 'in'
        $archiveFolder = Join-Path $testRoot 'Archive'
    }

    Context 'no file of that name in the archive' {
        It 'moves the file and keeps its name' {
            $file = New-XmlFileHC -Folder $inputFolder

            $result = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            $result.Archived | Should-BeTrue
            $result.IsDuplicate | Should-BeFalse
            $result.ArchivedAs | Should-Be 'Plant.xml'

            Test-Path (Join-Path $archiveFolder 'Plant.xml') | Should-BeTrue
            Test-Path $file.FullName | Should-BeFalse
        }

        It 'creates the archive folder when it does not exist yet' {
            $file = New-XmlFileHC -Folder $inputFolder

            $null = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            Test-Path $archiveFolder | Should-BeTrue
        }
    }

    Context 'a file of that name is in the archive with the same content' {
        BeforeEach {
            $null = New-XmlFileHC -Folder $archiveFolder -Content '<plant>same</plant>'
            $script:file = New-XmlFileHC -Folder $inputFolder -Content '<plant>same</plant>'
        }

        It 'archives it as a duplicate' {
            $result = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            $result.Archived | Should-BeTrue
            $result.IsDuplicate | Should-BeTrue
        }

        It 'gives the duplicate a name with the Duplicate prefix and the run time' {
            $result = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            $result.ArchivedAs | Should-Be "Duplicate_${stamp}_Plant.xml"
            Test-Path (Join-Path $archiveFolder "Duplicate_${stamp}_Plant.xml") | Should-BeTrue
        }

        It 'never overwrites the copy that was already archived' {
            $null = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            Get-Content (Join-Path $archiveFolder 'Plant.xml') -Raw |
            Should-Be '<plant>same</plant>'
        }

        It 'removes the file from the input folder' {
            $null = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            Test-Path $file.FullName | Should-BeFalse
        }

        It 'adds a counter when a duplicate of the same run is already there' {
            $null = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            $second = New-XmlFileHC -Folder $inputFolder -Content '<plant>same</plant>'

            $result = Move-ToArchiveFolderHC -File $second `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            $result.Archived | Should-BeTrue
            $result.ArchivedAs | Should-Be "Duplicate_${stamp}_2_Plant.xml"
        }
    }

    Context 'a file of that name is in the archive with different content' {
        BeforeEach {
            $null = New-XmlFileHC -Folder $archiveFolder -Content '<plant>old</plant>'
            $script:file = New-XmlFileHC -Folder $inputFolder -Content '<plant>new and longer</plant>'
        }

        It 'does not archive the file' {
            $result = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            $result.Archived | Should-BeFalse
            $result.IsDuplicate | Should-BeFalse
        }

        It 'leaves the file in place so a manual action can be taken' {
            $null = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            Test-Path $file.FullName | Should-BeTrue
        }

        It 'leaves the archived file untouched' {
            $null = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            Get-Content (Join-Path $archiveFolder 'Plant.xml') -Raw |
            Should-Be '<plant>old</plant>'
        }

        It 'explains that the records were not added to the Excel file' {
            $result = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            $result.Reason | Should-MatchString 'content is different'
            $result.Reason | Should-MatchString 'NOT added to the Excel file'
        }

        It 'notices a difference even when both files are the same size' {
            <#
                The length is only a shortcut, the hash has the last word.
            #>
            $null = New-XmlFileHC -Folder $archiveFolder -Content '<plant>aaa</plant>'
            $sameSize = New-XmlFileHC -Folder $inputFolder -Content '<plant>bbb</plant>'

            $result = Move-ToArchiveFolderHC -File $sameSize `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            $result.Archived | Should-BeFalse
        }
    }

    Context 'the file is gone' {
        It 'reports it instead of throwing' {
            $file = New-XmlFileHC -Folder $inputFolder
            Remove-Item -LiteralPath $file.FullName -Force

            $result = Move-ToArchiveFolderHC -File $file `
                -ArchiveFolder $archiveFolder -ScriptStartTime $startTime

            $result.Archived | Should-BeFalse
            $result.Reason | Should-MatchString 'no longer in the folder'
        }
    }
}