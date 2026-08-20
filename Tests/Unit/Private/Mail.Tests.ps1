#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Mail.ps1"
}

Describe 'Get-MailRecipientListHC' {
    It 'trims the addresses' {
        $settings = [PSCustomObject]@{ To = @('  bob@contoso.com  ') }

        Get-MailRecipientListHC -SendMailSettings $settings | Should-Be 'bob@contoso.com'
    }

    It 'de-duplicates case insensitively' {
        $settings = [PSCustomObject]@{ To = @('bob@contoso.com', 'BOB@contoso.com') }

        @(Get-MailRecipientListHC -SendMailSettings $settings) | Should-BeCollection -Count 1
    }

    It 'drops blank entries' {
        $settings = [PSCustomObject]@{ To = @('a@b.com', '', $null) }

        @(Get-MailRecipientListHC -SendMailSettings $settings) |
        Should-BeCollection @('a@b.com')
    }
}

Describe 'Get-MailBodyLogPathHC' {
    It 'replaces characters that are invalid in a file name' {
        $result = Get-MailBodyLogPathHC -MailParams @{ Subject = 'Results Q1/Q2' } `
            -LogFolder 'TestDrive:\'

        Split-Path $result -Leaf | Should-Be 'Mail - Results Q1 Q2.html'
    }

    It 'returns nothing when the log folder does not exist' {
        Get-MailBodyLogPathHC -MailParams @{ Subject = 'x' } `
            -LogFolder 'TestDrive:\DoesNotExist' | Should-BeFalsy
    }
}

Describe 'Save-MailBodyToLogHC' {
    It 'writes the body to an html file named after the subject' {
        $path = Save-MailBodyToLogHC `
            -MailParams @{ Subject = 'Nightly run'; Body = '<p>done</p>' } `
            -LogFolder 'TestDrive:\'

        Test-Path $path | Should-BeTrue
        Get-Content $path -Raw | Should-MatchString '<p>done</p>'
    }
}
