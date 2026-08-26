#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\ConvertXmlToExcel"

    . "$moduleRoot\Private\Get-MonthKeyHC.ps1"
}

Describe 'Get-MonthKeyHC' {
    It 'formats a date as yyyyMM MMMM' {
        Get-MonthKeyHC -Date ([datetime]'2024-08-16') | Should-Be '202408 August'
    }

    It 'gives the same key for two dates in the same month' {
        $a = Get-MonthKeyHC -Date ([datetime]'2024-08-01')
        $b = Get-MonthKeyHC -Date ([datetime]'2024-08-31')
        $a | Should-Be $b
    }

    It 'gives a different key across a month boundary' {
        $aug = Get-MonthKeyHC -Date ([datetime]'2024-08-31')
        $sep = Get-MonthKeyHC -Date ([datetime]'2024-09-01')
        $aug | Should-NotBe $sep
    }

    It 'gives a different key for the same month in another year' {
        $a = Get-MonthKeyHC -Date ([datetime]'2024-08-16')
        $b = Get-MonthKeyHC -Date ([datetime]'2025-08-16')
        $a | Should-NotBe $b
    }

    It 'sorts chronologically as text because the year and month lead' {
        $keys = @(
            Get-MonthKeyHC -Date ([datetime]'2024-09-01')
            Get-MonthKeyHC -Date ([datetime]'2024-08-01')
            Get-MonthKeyHC -Date ([datetime]'2025-01-01')
        ) | Sort-Object

        $keys[0] | Should-Be '202408 August'
        $keys[1] | Should-Be '202409 September'
        $keys[2] | Should-Be '202501 January'
    }
}

Describe 'Get-MonthKeyHC is not affected by the culture of the machine' {
    BeforeEach {
        $originalCulture = [Threading.Thread]::CurrentThread.CurrentCulture
    }

    AfterEach {
        [Threading.Thread]::CurrentThread.CurrentCulture = $originalCulture
    }

    It 'writes the month name in English on a <Name> machine' -ForEach @(
        @{ Name = 'Dutch'; Culture = 'nl-BE' }
        @{ Name = 'French'; Culture = 'fr-FR' }
        @{ Name = 'German'; Culture = 'de-DE' }
    ) {
        <#
            Without InvariantCulture this returns '202408 augustus' on the
            Dutch culture, which becomes a different Excel file name and
            silently starts an empty workbook next to the real one.
        #>
        [Threading.Thread]::CurrentThread.CurrentCulture = [CultureInfo]::new($Culture)

        Get-MonthKeyHC -Date ([datetime]'2024-08-16') | Should-Be '202408 August'
    }

    It 'gives the same key on two different cultures' {
        [Threading.Thread]::CurrentThread.CurrentCulture = [CultureInfo]::new('nl-BE')
        $dutch = Get-MonthKeyHC -Date ([datetime]'2025-03-16')

        [Threading.Thread]::CurrentThread.CurrentCulture = [CultureInfo]::new('en-US')
        $english = Get-MonthKeyHC -Date ([datetime]'2025-03-16')

        $dutch | Should-Be $english
    }
}

Describe 'Get-MonthRangeHC' {
    It 'returns the first moment of the month and of the next month' {
        $result = Get-MonthRangeHC -MonthKey '202408 August'

        $result.Start | Should-Be ([datetime]'2024-08-01T00:00:00')
        $result.End | Should-Be ([datetime]'2024-09-01T00:00:00')
    }

    It 'rolls over to the next year in December' {
        $result = Get-MonthRangeHC -MonthKey '202412 December'

        $result.Start | Should-Be ([datetime]'2024-12-01T00:00:00')
        $result.End | Should-Be ([datetime]'2025-01-01T00:00:00')
    }

    It 'agrees with Get-MonthKeyHC on every month of a year' -ForEach (1..12) {
        $date = [datetime]::new(2024, $_, 15)

        $range = Get-MonthRangeHC -MonthKey (Get-MonthKeyHC -Date $date)

        $date -ge $range.Start | Should-BeTrue
        $date -lt $range.End | Should-BeTrue
    }

    It 'places the last moment of a month inside that month' {
        $range = Get-MonthRangeHC -MonthKey '202408 August'

        $lastMoment = [datetime]'2024-08-31T23:59:59.999'

        $lastMoment -ge $range.Start | Should-BeTrue
        $lastMoment -lt $range.End | Should-BeTrue
    }

    It 'places the first moment of the next month outside the month' {
        $range = Get-MonthRangeHC -MonthKey '202408 August'

        [datetime]'2024-09-01T00:00:00' -lt $range.End | Should-BeFalse
    }

    It 'ignores the month name, only yyyyMM is read' {
        $withName = Get-MonthRangeHC -MonthKey '202408 August'
        $withoutName = Get-MonthRangeHC -MonthKey '202408'

        $withName.Start | Should-Be $withoutName.Start
        $withName.End | Should-Be $withoutName.End
    }

    It 'reads a key on a machine with another regional setting' {
        $originalCulture = [Threading.Thread]::CurrentThread.CurrentCulture

        try {
            [Threading.Thread]::CurrentThread.CurrentCulture = [CultureInfo]::new('nl-BE')

            $result = Get-MonthRangeHC -MonthKey '202408 August'

            $result.Start | Should-Be ([datetime]'2024-08-01T00:00:00')
            $result.End | Should-Be ([datetime]'2024-09-01T00:00:00')
        }
        finally {
            [Threading.Thread]::CurrentThread.CurrentCulture = $originalCulture
        }
    }

    It 'throws on a key that does not start with yyyyMM' -ForEach @(
        'August', '2024', '2024AA August', '202499 Nonsense'
    ) {
        { Get-MonthRangeHC -MonthKey $_ } | Should-Throw
    }
}