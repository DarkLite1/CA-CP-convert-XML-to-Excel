#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\SummaryMail.ps1"

    function New-CountHC {
        param (
            $Total = 0, $Added = 0, $Already = 0, $FanOut = 0,
            $Archived = 0, $Duplicates = 0, $NotArchived = 0, $Errors = 0
        )
        @{
            TotalXmlFiles = $Total; Added = $Added; AlreadyInSheet = $Already
            FanOut = $FanOut; Archived = $Archived; Duplicates = $Duplicates
            NotArchived = $NotArchived; Errors = $Errors
        }
    }

    function New-AddedHC {
        param ([String]$MonthKey, [int]$Files = 1, [int]$Rows = 10)

        1..$Files | ForEach-Object {
            [PSCustomObject]@{
                MonthKey      = $MonthKey
                ExcelFilePath = "C:\Excel\Alarms - $MonthKey.xlsx"
                RowsAdded     = $Rows
                AddedToSheet  = $true
            }
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

    Context 'the Excel file table' {
        BeforeAll {
            $collection = @{
                NotArchived = @()
                Added       = @(
                    New-AddedHC -MonthKey '202510 October' -Files 9
                    New-AddedHC -MonthKey '202601 January' -Files 14
                    New-AddedHC -MonthKey '202511 November' -Files 21
                )
            }

            $script:body = (New-SummaryMailHC -Count (New-CountHC -Total 44 -Added 44) `
                    -Collection $collection -Type 'Alarm' -Path $path).Body
        }

        It 'lists the most recent month first' {
            $months = [regex]::Matches($body, '(\d{6} \w+)') |
            ForEach-Object { $_.Groups[1].Value } |
            Select-Object -Unique

            $months[0] | Should-Be '202601 January'
            $months[1] | Should-Be '202511 November'
            $months[2] | Should-Be '202510 October'
        }

        It 'writes one row per Excel file, not one per XML file' {
            <#
                44 XML files went into 3 workbooks, so the table holds 3 data
                rows plus the header row.
            #>
            $rows = [regex]::Matches($body, '<tr>').Count

            $rows | Should-BeGreaterThan 3
            $body | Should-MatchString 'Alarms - 202601 January\.xlsx'
        }

        It 'shows the quantity of XML files per Excel file' {
            $body | Should-MatchString '>21<'
            $body | Should-MatchString '>14<'
        }

        It 'says so when no file was added' {
            $mail = New-SummaryMailHC -Count (New-CountHC -Total 3) `
                -Collection $emptyCollection -Type 'Batch' -Path $path

            $mail.Body | Should-MatchString 'No XML file was added'
        }
    }

    Context 'the summary comes after the detail' {
        It 'places the totals below the Excel file table' {
            $collection = @{
                NotArchived = @()
                Added       = @(New-AddedHC -MonthKey '202601 January' -Files 2)
            }

            $body = (New-SummaryMailHC -Count (New-CountHC -Total 2 -Added 2) `
                    -Collection $collection -Type 'Batch' -Path $path).Body

            $summaryAt = $body.IndexOf('Summary</b>')
            $tableAt = $body.IndexOf('Excel file</th>')

            $tableAt | Should-BeLessThan $summaryAt
        }

        It 'always shows the total and the added quantity' {
            $body = (New-SummaryMailHC -Count (New-CountHC -Total 7 -Added 5) `
                    -Collection $emptyCollection -Type 'Batch' -Path $path).Body

            $body | Should-MatchString 'Total XML files'
            $body | Should-MatchString 'Added to Excel'
        }

        It 'hides an optional total that is zero' {
            $body = (New-SummaryMailHC -Count (New-CountHC -Total 2 -Added 2) `
                    -Collection $emptyCollection -Type 'Batch' -Path $path).Body

            $body.Contains('Already in Excel') | Should-BeFalse
            $body.Contains('Archived as duplicate') | Should-BeFalse
            $body.Contains('NOT archived') | Should-BeFalse
        }

        It 'always shows the error count, also when it is zero' {
            <#
                A missing line reads as 'not reported' rather than as 'none'.
            #>
            $body = (New-SummaryMailHC -Count (New-CountHC -Total 2 -Added 2) `
                    -Collection $emptyCollection -Type 'Batch' -Path $path).Body

            $body | Should-MatchString 'Errors'
        }

        It 'reports the files archived as a duplicate' {
            $body = (New-SummaryMailHC -Count (New-CountHC -Total 2 -Already 2 -Duplicates 2) `
                    -Collection $emptyCollection -Type 'Batch' -Path $path).Body

            $body | Should-MatchString 'Archived as duplicate'
        }

        It 'shows the errors and the not archived quantity when there are any' {
            $body = (New-SummaryMailHC -Count (New-CountHC -Total 2 -NotArchived 1 -Errors 3) `
                    -Collection $emptyCollection -Type 'Batch' -Path $path).Body

            $body | Should-MatchString 'NOT archived'
            $body | Should-MatchString 'Errors'
        }
    }

    Context 'files that need a manual action' {
        It 'lists the file name and the reason' {
            $collection = @{
                Added       = @()
                NotArchived = @(
                    [PSCustomObject]@{
                        File   = [PSCustomObject]@{ Name = 'Bad.xml' }
                        Reason = 'No date'
                    }
                )
            }

            $body = (New-SummaryMailHC -Count (New-CountHC -NotArchived 1) `
                    -Collection $collection -Type 'Batch' -Path $path).Body

            $body | Should-MatchString 'manual action required'
            $body | Should-MatchString 'Bad.xml'
            $body | Should-MatchString 'No date'
        }
    }

    Context 'the HTML works in the classic Outlook client' {
        BeforeAll {
            $collection = @{
                NotArchived = @()
                Added       = @(New-AddedHC -MonthKey '202601 January' -Files 2)
            }

            $script:outlookBody = (New-SummaryMailHC -Count (New-CountHC -Total 2 -Added 2) `
                    -Collection $collection -Type 'Batch' -Path $path).Body
        }

        It 'uses table attributes the Word engine understands' {
            $outlookBody | Should-MatchString "border='1'"
            $outlookBody | Should-MatchString "cellpadding='5'"
            $outlookBody | Should-MatchString "cellspacing='0'"
        }

        It 'styles every cell inline instead of using a stylesheet' {
            $outlookBody.Contains('<style') | Should-BeFalse
            $outlookBody.Contains('class=') | Should-BeFalse
        }

        It 'avoids layout the Word engine cannot render' {
            ($outlookBody -match 'display:\s*flex') | Should-BeFalse
            ($outlookBody -match 'display:\s*grid') | Should-BeFalse
            $outlookBody.Contains('float:') | Should-BeFalse
        }

        It 'links both folders' {
            $outlookBody | Should-MatchString 'C:\\Xml'
            $outlookBody | Should-MatchString 'C:\\Xlsx'
        }
    }
}