# Pester 5 tests for the Display Resolution section of AssetSelfReport.ps1.
# Only the pure functions are loaded (parsed out of the script), so nothing in the script runs.
# Run: docker run --rm -v "$PWD:/src" -w /src mcr.microsoft.com/powershell:latest pwsh -NoProfile -Command "Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser; Invoke-Pester ./tests -Output Detailed"

BeforeAll {
    $ScriptPath = Join-Path $PSScriptRoot '..' 'AssetSelfReport.ps1'
    $Ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$null, [ref]$null)
    $Wanted = 'Get-MonitorName', 'Get-PnpCodeFromInstanceName', 'Get-NativeDisplayMode', 'Get-PnpCodesFromSetId',
              'ConvertTo-ScalePercent', 'Format-DisplayLine', 'Format-DisplayReport', 'Select-CurrentDisplayLayout',
              'Get-PrimaryScaleForLayout'
    $Functions = $Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $Wanted -contains $n.Name }, $true)
    if ($Functions.Count -ne $Wanted.Count) { throw "Expected $($Wanted.Count) functions in AssetSelfReport.ps1, found $($Functions.Count)" }
    foreach ($f in $Functions) { . ([scriptblock]::Create($f.Extent.Text)) }

    function Bytes([string] $s) { [int[]][char[]]$s + 0, 0 }
    function Monitor($Code, $Name, $W, $H) {
        [PSCustomObject]@{ PnpCode = $Code; Name = $Name; Native = if ($W) { [PSCustomObject]@{ Width = $W; Height = $H } } else { $null } }
    }
    function Display($W, $H, $X = 0, $Y = 0) { [PSCustomObject]@{ Width = $W; Height = $H; X = $X; Y = $Y } }
}

Describe 'Get-MonitorName' {
    It 'keeps the field 93 format and drops NUL padding' {
        $m = [PSCustomObject]@{ ManufacturerName = Bytes 'DEL'; UserFriendlyName = Bytes 'DELL U2723QE'; SerialNumberID = Bytes 'ABC123' }
        Get-MonitorName -Monitor $m | Should -Be 'DEL DELL U2723QE [ABC123]'
    }
}

Describe 'Get-PnpCodeFromInstanceName' {
    It 'returns the manufacturer and product code' {
        Get-PnpCodeFromInstanceName -InstanceName 'DISPLAY\del40f5\5&2b1b1c0b&0&UID4357_0' | Should -Be 'DEL40F5'
    }
    It 'returns empty for an unexpected value' {
        Get-PnpCodeFromInstanceName -InstanceName 'garbage' | Should -Be ''
    }
}

Describe 'Get-NativeDisplayMode' {
    It 'picks the preferred mode' {
        $modes = [PSCustomObject]@{
            PreferredMonitorSourceModeIndex = 1
            MonitorSourceModes = @(
                [PSCustomObject]@{ HorizontalActivePixels = 1920; VerticalActivePixels = 1080 },
                [PSCustomObject]@{ HorizontalActivePixels = 3840; VerticalActivePixels = 2160 })
        }
        $n = Get-NativeDisplayMode -SourceModes $modes
        $n.Width | Should -Be 3840
        $n.Height | Should -Be 2160
    }
    It 'returns $null when the index is out of range' {
        $modes = [PSCustomObject]@{ PreferredMonitorSourceModeIndex = 5; MonitorSourceModes = @([PSCustomObject]@{ HorizontalActivePixels = 1920; VerticalActivePixels = 1080 }) }
        Get-NativeDisplayMode -SourceModes $modes | Should -BeNullOrEmpty
    }
    It 'returns $null when there is no EDID data' {
        Get-NativeDisplayMode -SourceModes $null | Should -BeNullOrEmpty
    }
}

Describe 'Get-PnpCodesFromSetId' {
    It 'finds each display code in order' {
        Get-PnpCodesFromSetId -SetId 'DELA0F2HXNPC12345_1F_07E3_8A*SAM0D79H4ZM_0F_07E4_2C' | Should -Be @('DELA0F2', 'SAM0D79')
    }
    It 'finds nothing in a NOEDID set' {
        @(Get-PnpCodesFromSetId -SetId 'NOEDID_8086_3E98_00000000_00020000_31063').Count | Should -Be 0
    }
}

Describe 'ConvertTo-ScalePercent' {
    It 'maps <Dpi> DPI to <Percent>%' -ForEach @(
        @{ Dpi = 96; Percent = 100 }, @{ Dpi = 120; Percent = 125 }, @{ Dpi = 144; Percent = 150 }, @{ Dpi = 192; Percent = 200 }) {
        ConvertTo-ScalePercent -Dpi $Dpi | Should -Be $Percent
    }
    It 'returns $null for a missing value' {
        ConvertTo-ScalePercent -Dpi $null | Should -BeNullOrEmpty
    }
}

Describe 'Format-DisplayLine' {
    It 'flags a 4K panel run at 1080p' {
        Format-DisplayLine -Name 'Dell' -Width 1920 -Height 1080 -ScalePercent 100 -Native ([PSCustomObject]@{ Width = 3840; Height = 2160 }) -IsPrimary $true |
            Should -Be 'Dell: 1920x1080 @100% = 1920x1080 usable (native 3840x2160) primary LOW-RES'
    }
    It 'flags a native 1080p laptop at 125% as under 1080p usable' {
        Format-DisplayLine -Name 'Panel' -Width 1920 -Height 1080 -ScalePercent 125 -Native ([PSCustomObject]@{ Width = 1920; Height = 1080 }) -IsPrimary $true |
            Should -Be 'Panel: 1920x1080 @125% = 1536x864 usable (native) primary UNDER-1080P'
    }
    It 'does not flag a 4K panel at 150% (2560x1440 usable)' {
        Format-DisplayLine -Name 'LG' -Width 3840 -Height 2160 -ScalePercent 150 -Native ([PSCustomObject]@{ Width = 3840; Height = 2160 }) -IsPrimary $true |
            Should -Be 'LG: 3840x2160 @150% = 2560x1440 usable (native) primary'
    }
    It 'with unknown scaling, still flags a resolution already under 1080p' {
        Format-DisplayLine -Name 'Old' -Width 1366 -Height 768 -ScalePercent $null -Native ([PSCustomObject]@{ Width = 1366; Height = 768 }) -IsPrimary $false |
            Should -Be 'Old: 1366x768 @? (native) UNDER-1080P'
    }
    It 'with unknown scaling and native mode, does not flag a 1080p display' {
        Format-DisplayLine -Name 'Display 2' -Width 1920 -Height 1080 -ScalePercent $null -Native $null -IsPrimary $false |
            Should -Be 'Display 2: 1920x1080 @? (native ?)'
    }
    It 'does not flag usable space of exactly 1920x1080 (4K at 200%)' {
        Format-DisplayLine -Name 'K' -Width 3840 -Height 2160 -ScalePercent 200 -Native ([PSCustomObject]@{ Width = 3840; Height = 2160 }) -IsPrimary $true |
            Should -Be 'K: 3840x2160 @200% = 1920x1080 usable (native) primary'
    }
    It 'does not call a portrait-rotated native panel LOW-RES' {
        Format-DisplayLine -Name 'Rotated' -Width 1080 -Height 1920 -ScalePercent $null -Native ([PSCustomObject]@{ Width = 1920; Height = 1080 }) -IsPrimary $false |
            Should -Be 'Rotated: 1080x1920 @? (native) UNDER-1080P'
    }
}

Describe 'Select-CurrentDisplayLayout' {
    BeforeAll {
        function Layout($SetId, $Stamp, $Count = 1) { [PSCustomObject]@{ SetId = $SetId; Timestamp = [uint64]$Stamp; Displays = @(if ($Count -gt 0) { 1..$Count | ForEach-Object { Display 1920 1080 } }) } }
    }
    It 'picks the newest layout' {
        (Select-CurrentDisplayLayout -Layouts @((Layout 'DELA0F2A' 100), (Layout 'SAM0D79B' 300), (Layout 'BOE095FC' 200))).SetId | Should -Be 'SAM0D79B'
    }
    It 'ignores layouts with no timestamp or no displays' {
        (Select-CurrentDisplayLayout -Layouts @((Layout 'DELA0F2A' 100), (Layout 'SAM0D79B' 0), (Layout 'BOE095FC' 900 0))).SetId | Should -Be 'DELA0F2A'
    }
    It 'prefers a real monitor layout over a newer NOEDID or SIMULATED one' {
        (Select-CurrentDisplayLayout -Layouts @((Layout 'DELA0F2A' 100), (Layout 'NOEDID_8086_3E98' 500), (Layout 'SIMULATED_8086_9A49' 600))).SetId | Should -Be 'DELA0F2A'
    }
    It 'treats a layout with any NOEDID display as not real' {
        (Select-CurrentDisplayLayout -Layouts @((Layout 'DELA0F2A' 100), (Layout 'DELA0F5B_FF+NOEDID_8086' 500 2))).SetId | Should -Be 'DELA0F2A'
    }
    It 'ignores a corrupt timestamp beyond the FILETIME range' {
        (Select-CurrentDisplayLayout -Layouts @((Layout 'DELA0F2A' 100), (Layout 'SAM0D79B' ([uint64]::MaxValue)))).SetId | Should -Be 'DELA0F2A'
    }
    It 'uses a NOEDID layout when it is all there is' {
        (Select-CurrentDisplayLayout -Layouts @((Layout 'NOEDID_1234_5678' 100))).SetId | Should -Be 'NOEDID_1234_5678'
    }
    It 'returns nothing when no layout is usable' {
        Select-CurrentDisplayLayout -Layouts @((Layout 'DELA0F2A' 0)) | Should -BeNullOrEmpty
    }
}

Describe 'Get-PrimaryScaleForLayout' {
    BeforeAll {
        $SignIn = [DateTime]::new(2026, 11, 6, 9, 0, 0, [DateTimeKind]::Utc)
    }
    It 'keeps the sign-in scaling when the layout predates sign-in' {
        Get-PrimaryScaleForLayout -ScalePercent 150 -SignedInUtc $SignIn -LayoutTimestamp ([uint64]$SignIn.AddMinutes(-5).ToFileTimeUtc()) | Should -Be 150
    }
    It 'drops it when the layout changed after sign-in (e.g. docked)' {
        Get-PrimaryScaleForLayout -ScalePercent 150 -SignedInUtc $SignIn -LayoutTimestamp ([uint64]$SignIn.AddHours(1).ToFileTimeUtc()) | Should -BeNullOrEmpty
    }
    It 'returns $null without a sign-in time' {
        Get-PrimaryScaleForLayout -ScalePercent 150 -SignedInUtc $null -LayoutTimestamp 1 | Should -BeNullOrEmpty
    }
}

Describe 'Format-DisplayReport' {
    It 'matches displays to monitors via the SetId and applies scaling to the primary only' {
        $monitors = @((Monitor 'SAM0D79' 'SAM S24 [1]' 1920 1080), (Monitor 'DELA0F2' 'DEL U2723QE [2]' 3840 2160))
        $displays = @((Display 1920 1080 0 0), (Display 1920 1080 1920 0))
        $report = Format-DisplayReport -Displays $displays -Monitors $monitors -SetId 'DELA0F2HXN_1F*SAM0D79H4Z_0F' -PrimaryScalePercent 100
        $report | Should -Be ("DEL U2723QE [2]: 1920x1080 @100% = 1920x1080 usable (native 3840x2160) primary LOW-RES`n" +
                              "SAM S24 [1]: 1920x1080 @? (native)")
    }
    It 'gives two identical monitors one line each' {
        $monitors = @((Monitor 'DELA0F2' 'DEL A [1]' 2560 1440), (Monitor 'DELA0F2' 'DEL A [2]' 2560 1440))
        $displays = @((Display 2560 1440 0 0), (Display 2560 1440 2560 0))
        $lines = (Format-DisplayReport -Displays $displays -Monitors $monitors -SetId 'DELA0F2X*DELA0F2Y' -PrimaryScalePercent 100) -split "`n"
        $lines[0] | Should -BeLike 'DEL A `[1`]:*'
        $lines[1] | Should -BeLike 'DEL A `[2`]:*'
    }
    It 'matches a single display to its monitor' {
        $report = Format-DisplayReport -Displays @((Display 1920 1080 0 0)) -Monitors @((Monitor 'DELA0F2' 'DEL U2723QE [2]' 3840 2160)) -SetId 'DELA0F2HXNPC12345_1F_07E3_8A' -PrimaryScalePercent 100
        $report | Should -Be 'DEL U2723QE [2]: 1920x1080 @100% = 1920x1080 usable (native 3840x2160) primary LOW-RES'
    }
    It 'names only the unmatched display "Display N" when one code has no connected monitor' {
        $monitors = @((Monitor 'SAM0D79' 'SAM S24 [1]' 1920 1080))
        $displays = @((Display 1920 1080 0 0), (Display 1920 1080 1920 0))
        $lines = (Format-DisplayReport -Displays $displays -Monitors $monitors -SetId 'DELA0F2X*SAM0D79Y' -PrimaryScalePercent $null) -split "`n"
        $lines[0] | Should -Be 'Display 1: 1920x1080 @? (native ?) primary'
        $lines[1] | Should -Be 'SAM S24 [1]: 1920x1080 @? (native)'
    }
    It 'gives both mirrored displays the shared desktop scaling' {
        $monitors = @((Monitor 'BOE095F' 'BOE [1]' 1920 1080), (Monitor 'EPS0001' 'EPS Projector [2]' 1920 1080))
        $displays = @((Display 1920 1080 0 0), (Display 1920 1080 0 0))
        $lines = (Format-DisplayReport -Displays $displays -Monitors $monitors -SetId 'BOE095F_0*EPS0001_0' -PrimaryScalePercent 125) -split "`n"
        $lines[0] | Should -Be 'BOE [1]: 1920x1080 @125% = 1536x864 usable (native) primary UNDER-1080P'
        $lines[1] | Should -Be 'EPS Projector [2]: 1920x1080 @125% = 1536x864 usable (native) primary UNDER-1080P'
    }
    It 'falls back to numbered displays when the SetId cannot be matched' {
        $report = Format-DisplayReport -Displays @((Display 1024 768 0 0)) -Monitors @() -SetId 'NOEDID_8086_3E98' -PrimaryScalePercent $null
        $report | Should -Be 'Display 1: 1024x768 @? (native ?) primary UNDER-1080P'
    }
}
