#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\SummaryMail.ps1"

    function New-CountHC {
        param (
            $Total = 0, $Added = 0, $Already = 0, $FanOut = 0,
            $Archived = 0, $NotArchived = 0, $Errors = 0
        )
        @{
            TotalXmlFiles = $Total; Added = $Added; AlreadyInSheet = $Already
            FanOut = $FanOut; Archived = $Archived
            NotArchived = $NotArchived; Errors = $Errors
        }
    }

    $script:path = @{ XmlFiles = 'C:\Xml'; ExcelFiles = 'C:\Xlsx' }
    $script:emptyCollection = @{ Added = @(); NotArchived = @() }
}

Describe 'New-SummaryMailHC' {
    Context 'the subject' {
        It 'uses the plural for more than one file' {
            $mail = New-SummaryMailHC -Count (New-CountHC -Added 2) `
                -Collection $emptyCollection -Type 'Batch' -Path $path

            $mail.Subject | Should-MatchString '2 files added'
        }

        It 'uses the singular for one file' {
            $mail = New-SummaryMailHC -Count (New-CountHC -Added 1) `
                -Collection $emptyCollection -Type 'Batch' -Path $path

            $mail.Subject | Should-MatchString '1 file added'
        }

        It 'reports the files that were not archived and the errors' {
            $mail = New-SummaryMailHC -Count (New-CountHC -Added 1 -NotArchived 1 -Errors 2) `
                -Collection $emptyCollection -Type 'Batch' -Path $path

            $mail.Subject | Should-MatchString 'not archived'
            $mail.Subject | Should-MatchString '2 errors'
        }
    }

    Context 'the body' {
        It 'names the type' {
            $mail = New-SummaryMailHC -Count (New-CountHC) `
                -Collection $emptyCollection -Type 'Alarm' -Path $path

            $mail.Body | Should-MatchString "type 'Alarm'"
        }

        It 'highlights the files that need a manual action' {
            $collection = @{
                Added       = @()
                NotArchived = @(
                    [PSCustomObject]@{
                        File   = [PSCustomObject]@{ Name = 'Bad.xml' }
                        Reason = 'No date'
                    }
                )
            }

            $mail = New-SummaryMailHC -Count (New-CountHC -NotArchived 1) `
                -Collection $collection -Type 'Batch' -Path $path

            $mail.Body | Should-MatchString 'manual action required'
            $mail.Body | Should-MatchString 'Bad.xml'
        }

        It 'links both folders' {
            $mail = New-SummaryMailHC -Count (New-CountHC) `
                -Collection $emptyCollection -Type 'Batch' -Path $path

            $mail.Body | Should-MatchString 'C:\\Xml'
            $mail.Body | Should-MatchString 'C:\\Xlsx'
        }
    }
}
