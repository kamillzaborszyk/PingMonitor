#Requires -Version 5.1
<#
.SYNOPSIS
    Ping Monitor Widget — resizebar, transparent, Always-on-Top
.DESCRIPTION
    Verwendet System.Net.NetworkInformation.Ping.SendAsync() statt Threads.
    Kein Runspace-Konflikt, kein UI-Freeze.
.NOTES
    Starten mit: powershell -ExecutionPolicy Bypass -File PingMonitor.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ─── Theme definitions ────────────────────────────────────────────────────────
$script:IsDark = $true

$THEMES = @{
    Dark = @{
        BG       = [System.Drawing.Color]::FromArgb(13,  13,  15)
        SURFACE  = [System.Drawing.Color]::FromArgb(23,  23,  28)
        SURFACE2 = [System.Drawing.Color]::FromArgb(17,  17,  22)
        BORDER   = [System.Drawing.Color]::FromArgb(34,  34,  48)
        TEXT     = [System.Drawing.Color]::FromArgb(220, 220, 228)
        MUTED    = [System.Drawing.Color]::FromArgb(80,  82,  108)
        NETPANEL = [System.Drawing.Color]::FromArgb(15,  15,  20)
        LOG      = [System.Drawing.Color]::FromArgb(10,  10,  14)
        CHART    = [System.Drawing.Color]::FromArgb(10,  10,  14)
        TOPBAR   = [System.Drawing.Color]::FromArgb(17,  17,  22)
    }
    Light = @{
        BG       = [System.Drawing.Color]::FromArgb(242, 242, 246)
        SURFACE  = [System.Drawing.Color]::FromArgb(255, 255, 255)
        SURFACE2 = [System.Drawing.Color]::FromArgb(232, 232, 238)
        BORDER   = [System.Drawing.Color]::FromArgb(190, 190, 205)
        TEXT     = [System.Drawing.Color]::FromArgb(20,  20,  30)
        MUTED    = [System.Drawing.Color]::FromArgb(100, 100, 130)
        NETPANEL = [System.Drawing.Color]::FromArgb(225, 225, 235)
        LOG      = [System.Drawing.Color]::FromArgb(252, 252, 255)
        CHART    = [System.Drawing.Color]::FromArgb(230, 230, 240)
        TOPBAR   = [System.Drawing.Color]::FromArgb(210, 210, 220)
    }
}

# Active color variables (set by Apply-Theme)
$C_BG        = $THEMES.Dark.BG
$C_SURFACE   = $THEMES.Dark.SURFACE
$C_SURFACE2  = $THEMES.Dark.SURFACE2
$C_BORDER    = $THEMES.Dark.BORDER
$C_TEXT      = $THEMES.Dark.TEXT
$C_MUTED     = $THEMES.Dark.MUTED

# Signal strength colors — same in both themes
$C_GOOD      = [System.Drawing.Color]::FromArgb(74,  222, 128)
$C_GOOD_DIM  = [System.Drawing.Color]::FromArgb(15,  60,  30)
$C_WARN      = [System.Drawing.Color]::FromArgb(251, 146, 60)
$C_WARN_DIM  = [System.Drawing.Color]::FromArgb(80,  42,  10)
$C_BAD       = [System.Drawing.Color]::FromArgb(248, 80,  80)
$C_BAD_DIM   = [System.Drawing.Color]::FromArgb(80,  15,  15)
$C_TIMEOUT   = [System.Drawing.Color]::FromArgb(50,  8,   8)
$C_TO_FG     = [System.Drawing.Color]::FromArgb(160, 50,  50)
$C_SCORE_A   = [System.Drawing.Color]::FromArgb(74,  222, 128)
$C_SCORE_B   = [System.Drawing.Color]::FromArgb(160, 230, 100)
$C_SCORE_C   = [System.Drawing.Color]::FromArgb(251, 146, 60)
$C_SCORE_D   = [System.Drawing.Color]::FromArgb(248, 80,  80)

$FONT_MONO   = New-Object System.Drawing.Font("Consolas", 9,  [System.Drawing.FontStyle]::Regular)
$FONT_MONO_B = New-Object System.Drawing.Font("Consolas", 13, [System.Drawing.FontStyle]::Bold)
$FONT_SCORE  = New-Object System.Drawing.Font("Consolas", 22, [System.Drawing.FontStyle]::Bold)
$FONT_SMALL  = New-Object System.Drawing.Font("Segoe UI", 7,  [System.Drawing.FontStyle]::Regular)
$FONT_UI     = New-Object System.Drawing.Font("Segoe UI", 9,  [System.Drawing.FontStyle]::Regular)

$DASH        = [string][char]0x2014

# Anchor style shortcuts
$AL   = [System.Windows.Forms.AnchorStyles]
$TL   = $AL::Top    -bor $AL::Left
$TR   = $AL::Top    -bor $AL::Right
$TLR  = $AL::Top    -bor $AL::Left -bor $AL::Right
$TLRB = $AL::Top    -bor $AL::Left -bor $AL::Right -bor $AL::Bottom
$BL   = $AL::Bottom -bor $AL::Left
$BR   = $AL::Bottom -bor $AL::Right
$BLR  = $AL::Bottom -bor $AL::Left -bor $AL::Right

# ─── State ─────────────────────────────────────────────────────────────────────
$script:Running          = $false
$script:History          = [System.Collections.Generic.List[object]]::new()
$script:Sent             = 0
$script:Lost             = 0
$script:MaxHist          = 40
$script:MaxLog           = 150
$script:PingPending      = $false
$script:LastGrade        = ""
$script:ConsecTimeouts   = 0    # Counts consecutive timeouts
$script:PingSentAt       = [datetime]::MinValue  # Watchdog: when last ping was sent
$script:DynIntEnabled    = $false    # Dynamic interval mode
$script:DynIntCurrent    = 2         # Current effective interval (seconds)
$script:DynStableStreak  = 0         # Consecutive ticks at same tier (for hysteresis)

# ─── P/Invoke: DestroyIcon + Drag/Resize (single Add-Type block) ───────────────
if (-not ([System.Management.Automation.PSTypeName]'WinApi').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class WinApi {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
    [DllImport("user32.dll")]
    public static extern bool ReleaseCapture();
    [DllImport("user32.dll")]
    public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
    public const int WM_NCLBUTTONDOWN = 0xA1;
    public const int HTCAPTION        = 2;
}
'@
}

# Tray-Icon: Score-Buchstabe + Länderflagge als Bitmap — GDI-Handle wird sauber freigegeben
$script:GeoCountryCode = ""

function Get-FlagStripes($countryCode) {
    if (-not $countryCode) { return @() }
    switch ($countryCode.ToUpper()) {
        'DE' { return @([System.Drawing.Color]::FromArgb(0,0,0),       [System.Drawing.Color]::FromArgb(221,0,0),     [System.Drawing.Color]::FromArgb(255,206,0)) }
        'PL' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(220,20,60)) }
        'US' { return @([System.Drawing.Color]::FromArgb(60,59,110),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(178,34,52)) }
        'GB' { return @([System.Drawing.Color]::FromArgb(0,36,125),    [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(207,20,43)) }
        'FR' { return @([System.Drawing.Color]::FromArgb(0,35,149),    [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(237,41,57)) }
        'NL' { return @([System.Drawing.Color]::FromArgb(174,28,40),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(33,70,139)) }
        'IT' { return @([System.Drawing.Color]::FromArgb(0,140,69),    [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(205,33,42)) }
        'ES' { return @([System.Drawing.Color]::FromArgb(198,11,30),   [System.Drawing.Color]::FromArgb(255,196,0),   [System.Drawing.Color]::FromArgb(198,11,30)) }
        'AT' { return @([System.Drawing.Color]::FromArgb(237,41,57),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(237,41,57)) }
        'CH' { return @([System.Drawing.Color]::FromArgb(218,41,28),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(218,41,28)) }
        'SE' { return @([System.Drawing.Color]::FromArgb(0,106,167),   [System.Drawing.Color]::FromArgb(254,204,0),   [System.Drawing.Color]::FromArgb(0,106,167)) }
        'NO' { return @([System.Drawing.Color]::FromArgb(186,12,47),   [System.Drawing.Color]::FromArgb(0,32,91),     [System.Drawing.Color]::FromArgb(186,12,47)) }
        'DK' { return @([System.Drawing.Color]::FromArgb(198,12,48),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(198,12,48)) }
        'FI' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(0,47,108),    [System.Drawing.Color]::FromArgb(255,255,255)) }
        'BE' { return @([System.Drawing.Color]::FromArgb(0,0,0),       [System.Drawing.Color]::FromArgb(255,205,0),   [System.Drawing.Color]::FromArgb(237,41,57)) }
        'CZ' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(17,69,126),   [System.Drawing.Color]::FromArgb(215,20,26)) }
        'HU' { return @([System.Drawing.Color]::FromArgb(206,17,38),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(67,111,77)) }
        'RO' { return @([System.Drawing.Color]::FromArgb(0,43,127),    [System.Drawing.Color]::FromArgb(252,209,22),  [System.Drawing.Color]::FromArgb(206,17,38)) }
        'BG' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(0,150,110),   [System.Drawing.Color]::FromArgb(214,38,18)) }
        'PT' { return @([System.Drawing.Color]::FromArgb(0,102,0),     [System.Drawing.Color]::FromArgb(255,0,0),     [System.Drawing.Color]::FromArgb(255,0,0)) }
        'IE' { return @([System.Drawing.Color]::FromArgb(0,155,72),    [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(255,121,0)) }
        'RU' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(0,57,166),    [System.Drawing.Color]::FromArgb(213,43,30)) }
        'UA' { return @([System.Drawing.Color]::FromArgb(0,87,183),    [System.Drawing.Color]::FromArgb(255,215,0)) }
        'JP' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(188,0,45),    [System.Drawing.Color]::FromArgb(255,255,255)) }
        'KR' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(205,46,58),   [System.Drawing.Color]::FromArgb(0,71,160)) }
        'CN' { return @([System.Drawing.Color]::FromArgb(222,41,16),   [System.Drawing.Color]::FromArgb(255,222,0),   [System.Drawing.Color]::FromArgb(222,41,16)) }
        'IN' { return @([System.Drawing.Color]::FromArgb(255,153,51),  [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(19,136,8)) }
        'AU' { return @([System.Drawing.Color]::FromArgb(0,0,139),     [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(255,0,0)) }
        'BR' { return @([System.Drawing.Color]::FromArgb(0,156,59),    [System.Drawing.Color]::FromArgb(255,223,0),   [System.Drawing.Color]::FromArgb(0,39,118)) }
        'CA' { return @([System.Drawing.Color]::FromArgb(255,0,0),     [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(255,0,0)) }
        'MX' { return @([System.Drawing.Color]::FromArgb(0,104,71),    [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(206,17,38)) }
        'TR' { return @([System.Drawing.Color]::FromArgb(227,10,23),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(227,10,23)) }
        'GR' { return @([System.Drawing.Color]::FromArgb(13,94,175),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(13,94,175)) }
        'HR' { return @([System.Drawing.Color]::FromArgb(255,0,0),     [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(0,55,145)) }
        'SK' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(11,78,162),   [System.Drawing.Color]::FromArgb(238,28,37)) }
        'SI' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(0,51,160),    [System.Drawing.Color]::FromArgb(237,28,36)) }
        'LT' { return @([System.Drawing.Color]::FromArgb(253,185,19),  [System.Drawing.Color]::FromArgb(0,106,68),    [System.Drawing.Color]::FromArgb(193,39,45)) }
        'LV' { return @([System.Drawing.Color]::FromArgb(158,48,57),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(158,48,57)) }
        'EE' { return @([System.Drawing.Color]::FromArgb(0,114,206),   [System.Drawing.Color]::FromArgb(0,0,0),       [System.Drawing.Color]::FromArgb(255,255,255)) }
        'LU' { return @([System.Drawing.Color]::FromArgb(237,41,57),   [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(0,161,222)) }
        'IL' { return @([System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(0,56,184),    [System.Drawing.Color]::FromArgb(255,255,255)) }
        'ZA' { return @([System.Drawing.Color]::FromArgb(0,119,73),    [System.Drawing.Color]::FromArgb(255,184,28),  [System.Drawing.Color]::FromArgb(0,20,137)) }
        'AR' { return @([System.Drawing.Color]::FromArgb(116,172,223), [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(116,172,223)) }
        'TH' { return @([System.Drawing.Color]::FromArgb(237,28,36),   [System.Drawing.Color]::FromArgb(45,45,116),   [System.Drawing.Color]::FromArgb(237,28,36)) }
        'SG' { return @([System.Drawing.Color]::FromArgb(237,28,36),   [System.Drawing.Color]::FromArgb(255,255,255)) }
        'NZ' { return @([System.Drawing.Color]::FromArgb(0,0,107),     [System.Drawing.Color]::FromArgb(255,255,255), [System.Drawing.Color]::FromArgb(204,0,0)) }
        default { return @() }
    }
}

function New-ScoreIcon($grade, $color) {
    # DPI-aware: create bitmap at the system's actual icon size for crisp rendering
    # SmallIconSize adapts to DPI: 16@100%, 20@125%, 24@150%, 32@200%, 40@250%
    # We use 2x the system size (minimum 32) for high-quality downscaling
    $sysSize = [System.Windows.Forms.SystemInformation]::SmallIconSize.Width
    $size    = [Math]::Max($sysSize * 2, 32)

    $bmp = New-Object System.Drawing.Bitmap($size, $size)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

    $bgBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30, 30, 36))
    $g.FillRectangle($bgBrush, 0, 0, $size, $size)
    $bgBrush.Dispose()

    # Scale factor relative to the baseline 32px design
    $scale = $size / 32.0

    # Score letter — always full size, centered
    $fontSize = [Math]::Max(10, [Math]::Round(20 * $scale))
    $font  = New-Object System.Drawing.Font("Consolas", $fontSize, [System.Drawing.FontStyle]::Bold)
    $brush = New-Object System.Drawing.SolidBrush($color)
    $sf    = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString($grade, $font, $brush, ([System.Drawing.RectangleF]::new(0, 0, $size, $size)), $sf)
    $font.Dispose(); $brush.Dispose(); $sf.Dispose()

    # Flag badge — small rectangle in the bottom-right corner (scaled)
    $stripes = Get-FlagStripes $script:GeoCountryCode
    if ($stripes.Count -gt 0) {
        $fw = [Math]::Round(14 * $scale)
        $fh = [Math]::Round(10 * $scale)
        $margin = [Math]::Max(1, [Math]::Round(1 * $scale))
        $fx = $size - $fw - $margin
        $fy = $size - $fh - $margin

        # Dark border around flag for contrast
        $bw = [Math]::Max(1, [Math]::Round($scale))
        $borderBr = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(10, 10, 14))
        $g.FillRectangle($borderBr, $fx - $bw, $fy - $bw, $fw + 2*$bw, $fh + 2*$bw)
        $borderBr.Dispose()

        # Draw stripes inside the flag area
        $stripeH = [Math]::Floor($fh / $stripes.Count)
        for ($i = 0; $i -lt $stripes.Count; $i++) {
            $sy = $fy + ($i * $stripeH)
            $sh = if ($i -eq $stripes.Count - 1) { ($fy + $fh) - $sy } else { $stripeH }
            $sb = New-Object System.Drawing.SolidBrush($stripes[$i])
            $g.FillRectangle($sb, $fx, $sy, $fw, $sh)
            $sb.Dispose()
        }
    }

    $g.Dispose()

    $hIcon     = $bmp.GetHicon()
    $iconCopy  = [System.Drawing.Icon]::FromHandle($hIcon)
    $safeIcon  = $iconCopy.Clone()
    $iconCopy.Dispose()
    [WinApi]::DestroyIcon($hIcon) | Out-Null
    $bmp.Dispose()
    return $safeIcon
}

# ─── Create .NET ping objects once and reuse ────────────────────────────────────
$script:Pinger     = New-Object System.Net.NetworkInformation.Ping
$script:PingOpts   = New-Object System.Net.NetworkInformation.PingOptions
$script:PingOpts.Ttl = 128
$script:PingBuffer = [byte[]]::new(32)

# ─── Helper functions ──────────────────────────────────────────────────────────
function Get-PingColor($ms) {
    if ($null -eq $ms) { return $C_BAD }
    if ($ms -lt 80)    { return $C_GOOD }
    if ($ms -lt 200)   { return $C_WARN }
    return $C_BAD
}
function Get-PingColorDim($ms) {
    if ($null -eq $ms) { return $C_TIMEOUT }
    if ($ms -lt 80)    { return $C_GOOD_DIM }
    if ($ms -lt 200)   { return $C_WARN_DIM }
    return $C_BAD_DIM
}
function Get-Jitter {
    $valid = @($script:History | Where-Object { -not $_.Timeout } | ForEach-Object { $_.Ms })
    if ($valid.Count -lt 2) { return $null }
    $sum = 0
    for ($i = 1; $i -lt $valid.Count; $i++) { $sum += [Math]::Abs($valid[$i] - $valid[$i-1]) }
    return [Math]::Round($sum / ($valid.Count - 1))
}
function Get-QualityScore {
    # Connection lost: 5+ consecutive timeouts → "X"
    if ($script:ConsecTimeouts -ge 5) {
        return @{ Grade="X"; Color=$C_BAD; Desc="No connection" }
    }
    $total = $script:History.Count
    if ($total -lt 3) { return @{ Grade="?"; Color=$C_MUTED; Desc="Not enough data" } }
    $valid = @($script:History | Where-Object { -not $_.Timeout } | ForEach-Object { $_.Ms })
    $timeouts = $total - $valid.Count
    $avg   = if ($valid.Count -gt 0) { ($valid | Measure-Object -Average).Average } else { 999 }
    $jit   = Get-Jitter
    $lossP = ($timeouts / $total) * 100
    $pts   = 0
    if     ($avg -lt 50)  { $pts += 0 } elseif ($avg -lt 100) { $pts += 1 } elseif ($avg -lt 200) { $pts += 3 } else { $pts += 5 }
    if ($null -ne $jit) {
        if ($jit -lt 15) { $pts += 0 } elseif ($jit -lt 40) { $pts += 1 } elseif ($jit -lt 80) { $pts += 2 } else { $pts += 3 }
    }
    if ($lossP -eq 0) { $pts += 0 } elseif ($lossP -lt 2) { $pts += 1 } elseif ($lossP -lt 5) { $pts += 2 } else { $pts += 3 }
    if     ($pts -le 1) { return @{ Grade="A"; Color=$C_SCORE_A; Desc="Excellent" } }
    elseif ($pts -le 3) { return @{ Grade="B"; Color=$C_SCORE_B; Desc="Good" } }
    elseif ($pts -le 6) { return @{ Grade="C"; Color=$C_SCORE_C; Desc="Limited" } }
    else                { return @{ Grade="D"; Color=$C_SCORE_D; Desc="Poor" } }
}

# ─── Dynamic interval adjustment ──────────────────────────────────────────────
# Adjusts ping frequency based on network quality:
#   Stable (A/B for 5+ ticks) → slow down to 2× base (save resources)
#   Degraded (C/D)            → use 1× base interval
#   Down (X)                  → speed up to 0.5× base (detect recovery fast)
# Hysteresis: needs 5 consecutive same-tier readings before changing.

function Update-DynamicInterval($grade) {
    if (-not $script:DynIntEnabled -or -not $script:Running) { return }

    $base = [double]$numInt.Value   # user-set base interval in seconds

    # Determine target tier: 0=fast, 1=normal, 2=slow
    $targetTier = switch ($grade) {
        'X' { 0 }
        'D' { 1 }
        'C' { 1 }
        'B' { 2 }
        'A' { 2 }
        default { 1 }
    }

    # Track streak for hysteresis
    if ($targetTier -ne $script:DynLastTier) {
        $script:DynStableStreak = 1
        $script:DynLastTier = $targetTier
    } else {
        $script:DynStableStreak++
    }

    # Only change after 5 consecutive same-tier readings (hysteresis)
    if ($script:DynStableStreak -lt 5) { return }

    # Compute new interval
    $newInt = switch ($targetTier) {
        0 { [Math]::Max(1.0, $base * 0.5) }   # Down: fast polling
        1 { $base }                             # Degraded: normal
        2 { [Math]::Min(60.0, $base * 2.0) }   # Stable: relax
    }

    # Round to 1 decimal
    $newInt = [Math]::Round($newInt, 1)

    # Only apply if actually different (avoid timer churn)
    if ($newInt -ne $script:DynIntCurrent) {
        $script:DynIntCurrent = $newInt
        $pingTimer.Interval = [int]($newInt * 1000)

        # Update status bar
        $arrow = if ($newInt -lt $base) { "▲" } elseif ($newInt -gt $base) { "▼" } else { "=" }
        $lblStatus.Text = "Monitoring: $($txtHost.Text)  ·  ${newInt}s $arrow"
    }
}

$script:DynLastTier = 1   # Start at normal tier

# ─── FIX: Find best PHYSICAL adapter (skip VPN/virtual adapters) ───────────────
# CACHED: The lookup is expensive (multiple Get-NetAdapter calls per route).
# Cache result for 10 seconds; invalidate on network change via Reset-PhysicalAdapterCache.

$script:CachedPhysRoute      = $null
$script:CachedPhysRouteTime  = [datetime]::MinValue
$script:CachedPhysRouteTTL   = 10   # seconds

function Reset-PhysicalAdapterCache {
    $script:CachedPhysRoute     = $null
    $script:CachedPhysRouteTime = [datetime]::MinValue
}

function Get-PhysicalAdapterRoute {
    # Return cached result if still valid
    $now = [datetime]::UtcNow
    if ($null -ne $script:CachedPhysRoute -and
        ($now - $script:CachedPhysRouteTime).TotalSeconds -lt $script:CachedPhysRouteTTL) {
        return $script:CachedPhysRoute
    }

    $result = $null
    $allRoutes = @(Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue |
                   Sort-Object { $_.InterfaceMetric + $_.RouteMetric })

    if ($allRoutes.Count -eq 0) {
        $script:CachedPhysRoute = $null
        $script:CachedPhysRouteTime = $now
        return $null
    }

    # Pass 1: Find a route on a physical, "Up" adapter
    foreach ($r in $allRoutes) {
        $a = Get-NetAdapter -InterfaceIndex $r.InterfaceIndex -ErrorAction SilentlyContinue
        if ($null -eq $a)          { continue }
        if ($a.Status -ne 'Up')    { continue }

        $isVirtual = $false

        try {
            if ($a.Virtual -eq $true)          { $isVirtual = $true }
            if ($a.HardwareInterface -eq $false) { $isVirtual = $true }
        } catch {}

        if (-not $isVirtual) {
            $desc = "$($a.InterfaceDescription) $($a.Name)"
            if ($desc -match 'Virtual|Hyper-V|vEthernet|VPN|TAP-Windows|Fortinet|Cisco AnyConnect|Palo Alto|GlobalProtect|WireGuard|OpenVPN|Wintun|Docker|WSL|Loopback|Teredo|ISATAP|6to4') {
                $isVirtual = $true
            }
        }

        try {
            if ($a.ConnectorPresent -eq $false) { $isVirtual = $true }
        } catch {}

        if ($isVirtual) { continue }

        $result = $r; break
    }

    # Pass 2: Fallback — no physical adapter found, return first Up adapter route
    if ($null -eq $result) {
        foreach ($r in $allRoutes) {
            $a = Get-NetAdapter -InterfaceIndex $r.InterfaceIndex -ErrorAction SilentlyContinue
            if ($null -ne $a -and $a.Status -eq 'Up') { $result = $r; break }
        }
    }

    # Pass 3: Last resort
    if ($null -eq $result) { $result = $allRoutes[0] }

    $script:CachedPhysRoute     = $result
    $script:CachedPhysRouteTime = $now
    return $result
}

# ─── Read network metadata ─────────────────────────────────────────────────────
function Get-NetworkInfo {
    $info = @{
        SSID          = $DASH
        Adapter       = $DASH
        Speed         = $DASH
        LocalIP       = $DASH
        Gateway       = $DASH
        DNS           = $DASH
        SignalQuality = $null
        Error         = ""
    }
    try {
        # FIX: Use physical adapter route instead of raw lowest-metric route
        $route = Get-PhysicalAdapterRoute
        if ($null -eq $route) { $info.Error = "Keine Route"; return $info }
        $ifIdx = $route.InterfaceIndex

        # Adapter object
        $adapter = Get-NetAdapter -InterfaceIndex $ifIdx -ErrorAction SilentlyContinue
        if ($null -eq $adapter) { $info.Error = "Kein Adapter idx=$ifIdx"; return $info }

        $info.Adapter = $adapter.Name

        # LinkSpeed: returned as UInt64 (bps) OR as String ("54 Mbps") depending on Windows
        $rawSpeed = $adapter.LinkSpeed
        if ($rawSpeed -is [string]) {
            if ($rawSpeed -match "([\d,.]+)\s*(G|M|K)?bps?") {
                $num  = [double]($Matches[1] -replace ',','.')
                $unit = $Matches[2]
                $bps  = switch ($unit) {
                    'G' { $num * 1e9 }
                    'M' { $num * 1e6 }
                    'K' { $num * 1e3 }
                    default { $num }
                }
                if     ($bps -ge 1e9) { $info.Speed = "$([Math]::Round($bps/1e9,1)) Gbit/s" }
                elseif ($bps -ge 1e6) { $info.Speed = "$([Math]::Round($bps/1e6)) Mbit/s" }
                else                  { $info.Speed = "$([Math]::Round($bps/1e3)) Kbit/s" }
            } else {
                $info.Speed = $rawSpeed
            }
        } elseif ($rawSpeed -gt 0) {
            if     ($rawSpeed -ge 1e9) { $info.Speed = "$([Math]::Round($rawSpeed/1e9,1)) Gbit/s" }
            elseif ($rawSpeed -ge 1e6) { $info.Speed = "$([Math]::Round($rawSpeed/1e6)) Mbit/s" }
            else                       { $info.Speed = "$([Math]::Round($rawSpeed/1e3)) Kbit/s" }
        }

        # IP address
        $ipCfg = Get-NetIPAddress -InterfaceIndex $ifIdx -AddressFamily IPv4 `
                     -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($ipCfg) { $info.LocalIP = "$($ipCfg.IPAddress)/$($ipCfg.PrefixLength)" }

        # Gateway
        if ($route.NextHop -and $route.NextHop -ne "0.0.0.0") {
            $info.Gateway = $route.NextHop
        }

        # DNS — 3 methods, first one that returns something wins
        $dnsAddrs = @()

        # Method 1: Get-NetIPConfiguration
        try {
            $ipCfgFull = Get-NetIPConfiguration -InterfaceIndex $ifIdx -ErrorAction Stop
            $dnsAddrs  = @($ipCfgFull.DNSServer | Where-Object { $_.AddressFamily -eq 2 } |
                           ForEach-Object { $_.ServerAddresses } |
                           Where-Object { $_ -notmatch '^127\.' -and $_ -match '^\d+\.\d+\.\d+\.\d+$' })
        } catch {}

        # Method 2: Get-DnsClientServerAddress
        if ($dnsAddrs.Count -eq 0) {
            try {
                $dnsObj   = Get-DnsClientServerAddress -InterfaceIndex $ifIdx -AddressFamily IPv4 -ErrorAction Stop
                $dnsAddrs = @($dnsObj.ServerAddresses | Where-Object { $_ -notmatch '^127\.' })
            } catch {}
        }

        # Method 3: Registry
        if ($dnsAddrs.Count -eq 0) {
            try {
                $ifGuid = (Get-NetAdapter -InterfaceIndex $ifIdx -ErrorAction SilentlyContinue).InterfaceGuid
                if ($ifGuid) {
                    foreach ($regBase in @(
                        "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\$ifGuid",
                        "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\$($ifGuid.ToLower())"
                    )) {
                        $reg = Get-ItemProperty -Path $regBase -ErrorAction SilentlyContinue
                        if ($reg) {
                            foreach ($key in @('DhcpNameServer','NameServer')) {
                                $val = $reg.$key
                                if ($val) {
                                    $candidates = @($val -split '[,\s]+' |
                                        Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' -and $_ -notmatch '^127\.' })
                                    if ($candidates.Count -gt 0) { $dnsAddrs = $candidates; break }
                                }
                            }
                            if ($dnsAddrs.Count -gt 0) { break }
                        }
                    }
                }
            } catch {}
        }

        # Method 4: Show gateway as likely DNS server
        if ($dnsAddrs.Count -eq 0 -and $info.Gateway -and $info.Gateway -ne $DASH) {
            $info.DNS = "$($info.Gateway) (via gateway)"
        } elseif ($dnsAddrs.Count -gt 0) {
            $info.DNS = ($dnsAddrs | Select-Object -First 2) -join ", "
        }

        # Detect WLAN
        $isWlan = ($adapter.PhysicalMediaType -like "*802.11*") -or
                  ($adapter.Name -like "*Wi-Fi*") -or
                  ($adapter.Name -like "*WLAN*") -or
                  ($adapter.Name -like "*Wireless*") -or
                  ($adapter.InterfaceDescription -like "*Wireless*") -or
                  ($adapter.InterfaceDescription -like "*Wi-Fi*") -or
                  ($adapter.NdisPhysicalMedium -eq 9)

        if ($isWlan) {
            $wlan = netsh wlan show interfaces 2>$null
            if ($wlan) {
                $ssidLine = $wlan | Where-Object { $_ -match "^\s+SSID\s*:" -and $_ -notmatch "BSSID" } |
                            Select-Object -First 1
                if ($ssidLine -match ":\s+(.+)$") { $info.SSID = $Matches[1].Trim() }

                $sigLine = $wlan | Where-Object { $_ -match "Signal|Empfangs" } | Select-Object -First 1
                if ($sigLine -match ":\s+(\d+)%") { $info.SignalQuality = [int]$Matches[1] }
            }
            if ($info.SSID -eq $DASH) { $info.SSID = "WLAN (SSID n/v)" }
        } else {
            $info.SSID = "LAN / Ethernet"
        }
    } catch {
        $info.Error = $_.Exception.Message
    }
    return $info
}

# ─── Public IP + Geolocation (ip-api.com, free, no key required) ───────────────
function Get-PublicIPInfo {
    try {
        $resp = Invoke-RestMethod -Uri "http://ip-api.com/json/?fields=status,message,country,regionName,city,isp,query" `
                                  -TimeoutSec 6 -ErrorAction Stop
        if ($resp.status -eq "success") {
            return @{
                IP      = $resp.query
                City    = $resp.city
                Region  = $resp.regionName
                Country = $resp.country
                ISP     = $resp.isp
                OK      = $true
            }
        }
    } catch {}
    return @{ IP = "n/v"; City = ""; Region = ""; Country = ""; ISP = ""; OK = $false }
}

# Load async so app start is not blocked (PowerShell job)
$script:GeoJob        = $null
$script:GeoChecked    = $false
$script:GeoFailed     = $false
$script:GeoRetryAt    = [datetime]::MinValue

function Start-GeoLookup {
    $script:GeoChecked = $false
    $script:GeoFailed  = $false
    $script:GeoJob = Start-Job -ScriptBlock {
        try {
            $r = Invoke-RestMethod -Uri "http://ip-api.com/json/?fields=status,message,country,countryCode,regionName,city,isp,query" -TimeoutSec 8
            if ($r.status -eq "success") {
                return "$($r.query)|$($r.city)|$($r.regionName)|$($r.country)|$($r.isp)|$($r.countryCode)"
            }
        } catch {}
        return "ERR"
    }
}

function Check-GeoJob {
    if ($script:GeoFailed -and $script:GeoChecked) {
        if ([datetime]::UtcNow -ge $script:GeoRetryAt) {
            $lblNetPubIP.Text = "Retrying..."
            $lblNetOrt.Text   = "Trying to connect..."
            Start-GeoLookup
        }
        return
    }

    if ($script:GeoChecked -or $null -eq $script:GeoJob) { return }
    $state = $script:GeoJob.State
    if ($state -eq "Completed" -or $state -eq "Failed") {
        $script:GeoChecked = $true
        try {
            $result = Receive-Job -Job $script:GeoJob -ErrorAction SilentlyContinue
            Remove-Job -Job $script:GeoJob -Force -ErrorAction SilentlyContinue
            $script:GeoJob = $null
            if ($result -and $result -ne "ERR") {
                $parts = $result -split '\|'
                $lblNetPubIP.Text  = $parts[0]
                $city   = $parts[1]; $region = $parts[2]; $country = $parts[3]; $isp = $parts[4]
                $cc     = if ($parts.Count -ge 6) { $parts[5] } else { "" }
                $ort    = @($city, $region, $country) | Where-Object { $_ -ne "" } | Select-Object -First 3
                $lblNetOrt.Text = ($ort -join ", ") + $(if ($isp) { "  ($isp)" } else { "" })
                if ($lblNetOrt.Text.Length -gt 55) { $lblNetOrt.Text = $lblNetOrt.Text.Substring(0, 52) + "..." }
                $script:GeoFailed = $false

                # Store country code and force tray icon refresh to show flag
                if ($cc -ne $script:GeoCountryCode) {
                    $script:GeoCountryCode = $cc
                    $script:LastGrade = ""   # Force tray icon rebuild on next Update-Stats
                    # Refresh immediately if we already have a score
                    if ($script:History.Count -ge 3) {
                        $q = Get-QualityScore
                        $newIcon  = New-ScoreIcon $q.Grade $q.Color
                        $oldIcon  = $trayIcon.Icon
                        $trayIcon.Icon = $newIcon
                        $script:LastGrade = $q.Grade
                        if ($null -ne $oldIcon) { try { $oldIcon.Dispose() } catch {} }
                    } elseif ($trayIcon) {
                        # Even in idle state, refresh to show flag under "?"
                        $newIcon  = New-ScoreIcon "?" $C_MUTED
                        $oldIcon  = $trayIcon.Icon
                        $trayIcon.Icon = $newIcon
                        if ($null -ne $oldIcon) { try { $oldIcon.Dispose() } catch {} }
                    }
                }
            } else {
                $lblNetPubIP.Text    = "Error  (Retry in 10s)"
                $lblNetOrt.Text      = "GeoIP unavailable"
                $script:GeoFailed    = $true
                $script:GeoRetryAt   = [datetime]::UtcNow.AddSeconds(10)
            }
        } catch {
            $lblNetPubIP.Text    = "Fehler  (Retry in 10s)"
            $lblNetOrt.Text      = $_.Exception.Message
            $script:GeoFailed    = $true
            $script:GeoRetryAt   = [datetime]::UtcNow.AddSeconds(10)
        }
    }
}

# ─── Bandwidth state ───────────────────────────────────────────────────────────
$script:BwIfIdx      = -1
$script:BwLastRx     = 0
$script:BwLastTx     = 0
$script:BwLastTime   = [datetime]::MinValue
$script:BwSessionRx  = 0
$script:BwSessionTx  = 0

function Format-Bytes($bytes) {
    if ($bytes -ge 1048576) { return "$([Math]::Round($bytes/1048576, 1)) MB/s" }
    elseif ($bytes -ge 1024) { return "$([Math]::Round($bytes/1024, 0)) KB/s" }
    else { return "$bytes B/s" }
}
function Format-BytesTotal($bytes) {
    if ($bytes -ge 1073741824) { return "$([Math]::Round($bytes/1073741824, 2)) GB" }
    elseif ($bytes -ge 1048576) { return "$([Math]::Round($bytes/1048576, 1)) MB" }
    elseif ($bytes -ge 1024) { return "$([Math]::Round($bytes/1024, 0)) KB" }
    else { return "$bytes B" }
}

$script:BwAdapterName = ""   # Cached adapter name for statistics lookup

function Update-Bandwidth {
    try {
        $route = Get-PhysicalAdapterRoute
        if ($null -eq $route) { $lblNetDL.Text = $DASH; $lblNetUL.Text = $DASH; return }

        # If the interface changed, resolve adapter name (otherwise reuse cached)
        if ($script:BwIfIdx -ne $route.InterfaceIndex) {
            $adapter = Get-NetAdapter -InterfaceIndex $route.InterfaceIndex -ErrorAction SilentlyContinue
            if ($null -eq $adapter) { $lblNetDL.Text = $DASH; $lblNetUL.Text = $DASH; return }
            $script:BwAdapterName = $adapter.Name
            $script:BwIfIdx    = $route.InterfaceIndex
            $script:BwLastTime = [datetime]::MinValue
            $script:BwLastRx   = 0
            $script:BwLastTx   = 0
        }

        $stats = Get-NetAdapterStatistics -Name $script:BwAdapterName -ErrorAction SilentlyContinue
        if ($null -eq $stats) { $lblNetDL.Text = $DASH; $lblNetUL.Text = $DASH; return }

        $now = [datetime]::UtcNow
        $rx  = [long]$stats.ReceivedBytes
        $tx  = [long]$stats.SentBytes

        if ($script:BwLastTime -ne [datetime]::MinValue) {
            $secs = ($now - $script:BwLastTime).TotalSeconds
            if ($secs -gt 0) {
                $dlRate = [Math]::Max(0, ($rx - $script:BwLastRx) / $secs)
                $ulRate = [Math]::Max(0, ($tx - $script:BwLastTx) / $secs)
                $script:BwSessionRx += [Math]::Max(0, $rx - $script:BwLastRx)
                $script:BwSessionTx += [Math]::Max(0, $tx - $script:BwLastTx)
                $lblNetDL.Text      = Format-Bytes $dlRate
                $lblNetUL.Text      = Format-Bytes $ulRate
                $lblNetDLtotal.Text = Format-BytesTotal $script:BwSessionRx
                $lblNetULtotal.Text = Format-BytesTotal $script:BwSessionTx
            }
        }
        $script:BwLastRx   = $rx
        $script:BwLastTx   = $tx
        $script:BwLastTime = $now
    } catch {
        $lblNetDL.Text = "Err"
        $lblNetUL.Text = $_.Exception.Message.Substring(0, [Math]::Min(20, $_.Exception.Message.Length))
    }
}

# ─── Form ──────────────────────────────────────────────────────────────────────
$form                  = New-Object System.Windows.Forms.Form
$form.Text             = "Ping Monitor"
$form.ClientSize       = New-Object System.Drawing.Size(350, 700)
$form.MinimumSize      = New-Object System.Drawing.Size(300, 610)
$form.FormBorderStyle  = [System.Windows.Forms.FormBorderStyle]::None
$form.BackColor        = $C_BG
$form.ForeColor        = $C_TEXT
$form.TopMost          = $true
$form.StartPosition    = [System.Windows.Forms.FormStartPosition]::Manual
$form.Location         = New-Object System.Drawing.Point(20, 80)
$form.Font             = $FONT_UI
$form.ShowInTaskbar    = $true
$form.Opacity          = 0.92

# ─── Custom titlebar ───────────────────────────────────────────────────────────
$pnlTop               = New-Object System.Windows.Forms.Panel
$pnlTop.BackColor     = [System.Drawing.Color]::FromArgb(17, 17, 22)
$pnlTop.Size          = New-Object System.Drawing.Size(350, 32)
$pnlTop.Location      = New-Object System.Drawing.Point(0, 0)
$pnlTop.Anchor        = $TLR

$lblTitle             = New-Object System.Windows.Forms.Label
$lblTitle.Text        = "◉  PING MONITOR"
$lblTitle.Font        = New-Object System.Drawing.Font("Segoe UI", 7.5, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor   = $C_MUTED
$lblTitle.Size        = New-Object System.Drawing.Size(86, 32)
$lblTitle.Location    = New-Object System.Drawing.Point(8, 0)
$lblTitle.TextAlign   = [System.Drawing.ContentAlignment]::MiddleLeft
$lblTitle.Cursor      = [System.Windows.Forms.Cursors]::SizeAll
$pnlTop.Controls.Add($lblTitle)

$btnTheme             = New-Object System.Windows.Forms.Button
$btnTheme.Text        = [string][char]0x263C
$btnTheme.Font        = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$btnTheme.ForeColor   = $C_MUTED
$btnTheme.BackColor   = [System.Drawing.Color]::FromArgb(17, 17, 22)
$btnTheme.FlatStyle   = [System.Windows.Forms.FlatStyle]::Flat
$btnTheme.FlatAppearance.BorderSize = 0
$btnTheme.Size        = New-Object System.Drawing.Size(26, 26)
$btnTheme.Location    = New-Object System.Drawing.Point(96, 3)
$btnTheme.Cursor      = [System.Windows.Forms.Cursors]::Hand
$btnTheme.TabStop     = $false
$pnlTop.Controls.Add($btnTheme)

$lblTransIcon         = New-Object System.Windows.Forms.Label
$lblTransIcon.Text    = "opacity"
$lblTransIcon.Font    = $FONT_SMALL
$lblTransIcon.ForeColor = $C_MUTED
$lblTransIcon.Size    = New-Object System.Drawing.Size(40, 32)
$lblTransIcon.Location= New-Object System.Drawing.Point(126, 0)
$lblTransIcon.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$pnlTop.Controls.Add($lblTransIcon)

$trkOpacity           = New-Object System.Windows.Forms.TrackBar
$trkOpacity.Minimum   = 20; $trkOpacity.Maximum = 100; $trkOpacity.Value = 92
$trkOpacity.SmallChange = 5; $trkOpacity.LargeChange = 10
$trkOpacity.TickStyle = [System.Windows.Forms.TickStyle]::None
$trkOpacity.BackColor = [System.Drawing.Color]::FromArgb(17, 17, 22)
$trkOpacity.Size      = New-Object System.Drawing.Size(100, 32)
$trkOpacity.Location  = New-Object System.Drawing.Point(168, 0)
$trkOpacity.Anchor    = $TR
$pnlTop.Controls.Add($trkOpacity)

$lblOpacityVal        = New-Object System.Windows.Forms.Label
$lblOpacityVal.Text   = "92%"
$lblOpacityVal.Font   = $FONT_SMALL; $lblOpacityVal.ForeColor = $C_MUTED
$lblOpacityVal.Size   = New-Object System.Drawing.Size(30, 32)
$lblOpacityVal.Location = New-Object System.Drawing.Point(270, 0)
$lblOpacityVal.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblOpacityVal.Anchor = $TR
$pnlTop.Controls.Add($lblOpacityVal)

# ─── Custom window buttons ─────────────────────────────────────────────────────
function New-WinBtn($symbol, $x, $hoverColor, $clickAction) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text      = $symbol
    $b.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $b.ForeColor = $C_MUTED
    $b.BackColor = [System.Drawing.Color]::FromArgb(17, 17, 22)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderSize  = 0
    $b.FlatAppearance.MouseOverBackColor = $hoverColor
    $b.Size      = New-Object System.Drawing.Size(32, 32)
    $b.Location  = New-Object System.Drawing.Point($x, 0)
    $b.Anchor    = $TR
    $b.TabStop   = $false
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $b.Add_Click($clickAction)
    $pnlTop.Controls.Add($b)
    return $b
}

$btnWinMin   = New-WinBtn "_"  282 ([System.Drawing.Color]::FromArgb(50,50,60))  { $form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized }
$btnWinClose = New-WinBtn "✕" 314 ([System.Drawing.Color]::FromArgb(180,30,30)) { $form.Close() }

$btnWinClose.Location = New-Object System.Drawing.Point(314, 0)
$btnWinMin.Location   = New-Object System.Drawing.Point(282, 0)
$btnWinClose.Anchor   = $TR; $btnWinMin.Anchor = $TR

$trkOpacity.Add_ValueChanged({
    $form.Opacity = $trkOpacity.Value / 100.0
    $lblOpacityVal.Text = "$($trkOpacity.Value)%"
})

$script:Dragging = $false; $script:DragStart = [System.Drawing.Point]::Empty
$dragHandler = {
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        [WinApi]::ReleaseCapture()
        [WinApi]::SendMessage($form.Handle, [WinApi]::WM_NCLBUTTONDOWN, [WinApi]::HTCAPTION, 0) | Out-Null
    }
}
$pnlTop.Add_MouseDown($dragHandler)
$lblTitle.Add_MouseDown($dragHandler)
$lblTransIcon.Add_MouseDown($dragHandler)

# ─── Resize-Grip Panel ────────────────────────────────────────────────────────
$script:ResizeStart  = [System.Drawing.Point]::Empty
$script:ResizeFormSz = [System.Drawing.Size]::Empty
$script:ResizeFormPt = [System.Drawing.Point]::Empty
$script:ResizeDir    = ""

$pnlGrip             = New-Object System.Windows.Forms.Panel
$pnlGrip.BackColor   = [System.Drawing.Color]::Transparent
$pnlGrip.Size        = New-Object System.Drawing.Size(350, 5)
$pnlGrip.Location    = New-Object System.Drawing.Point(0, 695)
$pnlGrip.Anchor      = $BLR
$pnlGrip.Cursor      = [System.Windows.Forms.Cursors]::SizeNS
$form.Controls.Add($pnlGrip)

$pnlGripCorner        = New-Object System.Windows.Forms.Panel
$pnlGripCorner.BackColor = [System.Drawing.Color]::Transparent
$pnlGripCorner.Size   = New-Object System.Drawing.Size(12, 12)
$pnlGripCorner.Location = New-Object System.Drawing.Point(338, 688)
$pnlGripCorner.Anchor = $BR
$pnlGripCorner.Cursor = [System.Windows.Forms.Cursors]::SizeNWSE
$form.Controls.Add($pnlGripCorner)

$resizeMouseDown = {
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        $script:ResizeStart  = [System.Windows.Forms.Control]::MousePosition
        $script:ResizeFormSz = $form.Size
        $script:ResizeFormPt = $form.Location
        $script:ResizeDir    = if ($s -eq $pnlGripCorner) { "SE" } else { "S" }
    }
}
$resizeMouseMove = {
    param($s, $e)
    if ($script:ResizeDir -ne "" -and [System.Windows.Forms.Control]::MouseButtons -eq [System.Windows.Forms.MouseButtons]::Left) {
        $cur   = [System.Windows.Forms.Control]::MousePosition
        $dx    = $cur.X - $script:ResizeStart.X
        $dy    = $cur.Y - $script:ResizeStart.Y
        $newW  = [Math]::Max($form.MinimumSize.Width,  $script:ResizeFormSz.Width  + $(if ($script:ResizeDir -eq "SE") { $dx } else { 0 }))
        $newH  = [Math]::Max($form.MinimumSize.Height, $script:ResizeFormSz.Height + $dy)
        $form.Size = New-Object System.Drawing.Size($newW, $newH)
    }
}
$resizeMouseUp = { $script:ResizeDir = "" }

foreach ($grip in @($pnlGrip, $pnlGripCorner)) {
    $grip.Add_MouseDown($resizeMouseDown)
    $grip.Add_MouseMove($resizeMouseMove)
    $grip.Add_MouseUp($resizeMouseUp)
}

$form.Controls.Add($pnlTop)

# ─── Apply theme ───────────────────────────────────────────────────────────────
function Apply-Theme {
    $t = if ($script:IsDark) { $THEMES.Dark } else { $THEMES.Light }

    $form.BackColor      = $t.BG
    $pnlTop.BackColor    = $t.TOPBAR
    $netPanel.BackColor  = $t.NETPANEL
    $chartPanel.BackColor= $t.CHART
    $chartCanvas.BackColor= $t.CHART
    $legendPanel.BackColor= $t.BG
    $logBox.BackColor    = $t.LOG
    $sep1.BackColor      = $t.BORDER
    $sep1b.BackColor     = $t.BORDER
    $sep2.BackColor      = $t.BORDER
    $statsTable.BackColor= $t.BG

    $txtHost.BackColor   = $t.SURFACE; $txtHost.ForeColor  = $t.TEXT
    $numInt.BackColor    = $t.SURFACE; $numInt.ForeColor   = $t.TEXT
    $chkDynInt.ForeColor = $t.MUTED;  $chkDynInt.BackColor = $t.BG

    foreach ($c in @($cardCur, $cardJitter, $cardLoss, $cardScore)) {
        $c.Panel.BackColor = $t.SURFACE2
    }

    foreach ($ctrl in $netPanel.Controls) {
        if ($ctrl -is [System.Windows.Forms.Label]) {
            $ctrl.BackColor = [System.Drawing.Color]::Transparent
            if ($ctrl.Font.Bold) { $ctrl.ForeColor = $t.TEXT }
            else                 { $ctrl.ForeColor = $t.MUTED }
        }
    }

    $muted  = @($lblHostCap, $lblIntCap, $lblLog, $lblStatus, $lblMinMax,
                $lblChart, $lblTitle, $lblTransIcon, $lblOpacityVal,
                $lblScoreDesc)
    foreach ($l in $muted) { try { $l.ForeColor = $t.MUTED } catch {} }

    foreach ($wb in @($btnWinMin, $btnWinClose)) {
        $wb.BackColor = $t.TOPBAR
        $wb.ForeColor = $t.MUTED
        $wb.FlatAppearance.MouseOverBackColor = if ($wb -eq $btnWinClose) {
            [System.Drawing.Color]::FromArgb(180, 30, 30)
        } else { [System.Drawing.Color]::FromArgb(60, 60, 80) }
    }

    $logBox.ForeColor    = $t.TEXT
    $btnTheme.BackColor  = $t.TOPBAR
    $btnTheme.ForeColor  = $t.MUTED
    $trkOpacity.BackColor = $t.TOPBAR

    $chartCanvas.Invalidate()
    $form.Refresh()
}

$btnTheme.Add_Click({
    $script:IsDark = -not $script:IsDark
    $btnTheme.Text = if ($script:IsDark) { [string][char]0x263C } else { [string][char]0x263D }
    Apply-Theme
})

$sep1           = New-Object System.Windows.Forms.Panel
$sep1.BackColor = $C_BORDER
$sep1.Size      = New-Object System.Drawing.Size(350, 1)
$sep1.Location  = New-Object System.Drawing.Point(0, 32)
$sep1.Anchor    = $TLR
$form.Controls.Add($sep1)

# ─── Network info panel ────────────────────────────────────────────────────────
$netPanel             = New-Object System.Windows.Forms.Panel
$netPanel.BackColor   = [System.Drawing.Color]::FromArgb(15, 15, 20)
$netPanel.Size        = New-Object System.Drawing.Size(350, 136)
$netPanel.Location    = New-Object System.Drawing.Point(0, 33)
$netPanel.Anchor      = $TLR

function New-NetLabel($text, $x, $y, $w, $bold) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $text
    $l.Font      = if ($bold) { New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold) } else { $FONT_SMALL }
    $l.ForeColor = if ($bold) { $C_TEXT } else { $C_MUTED }
    $l.Size      = New-Object System.Drawing.Size($w, 14)
    $l.Location  = New-Object System.Drawing.Point($x, $y)
    $l.AutoEllipsis = $true
    $netPanel.Controls.Add($l)
    return $l
}

New-NetLabel "NET"    12  6  40 $false | Out-Null
New-NetLabel "ADAPT." 12 24  40 $false | Out-Null
New-NetLabel "SPEED"   12 42  40 $false | Out-Null
New-NetLabel "IP"     188  6  24 $false | Out-Null
New-NetLabel "GW"      188 24  22 $false | Out-Null
New-NetLabel "DNS"    188 42  24 $false | Out-Null
New-NetLabel "DL"      12 62  22 $false | Out-Null
New-NetLabel "UL"      12 80  22 $false | Out-Null
New-NetLabel "DL tot." 188 62  44 $false | Out-Null
New-NetLabel "UL tot." 188 80  44 $false | Out-Null
New-NetLabel "PUB IP"  12  98  44 $false | Out-Null
New-NetLabel "LOC"     12 116  28 $false | Out-Null

$lblNetSSID    = New-NetLabel $DASH  55  6 125 $true
$lblNetAdapter = New-NetLabel $DASH  55 24 125 $true
$lblNetSpeed   = New-NetLabel $DASH  55 42 125 $true
$lblNetIP      = New-NetLabel $DASH 216  6 120 $true
$lblNetGateway = New-NetLabel $DASH 213 24 123 $true
$lblNetDNS     = New-NetLabel $DASH 216 42 120 $true
$lblNetDL      = New-NetLabel $DASH  38 62 140 $true
$lblNetUL      = New-NetLabel $DASH  38 80 140 $true
$lblNetDLtotal = New-NetLabel $DASH 236 62 100 $true
$lblNetULtotal = New-NetLabel $DASH 236 80 100 $true
$lblNetPubIP   = New-NetLabel "..."  60  98 270 $true
$lblNetOrt     = New-NetLabel "..."  44 116 292 $true

$form.Controls.Add($netPanel)

$sep1b            = New-Object System.Windows.Forms.Panel
$sep1b.BackColor  = $C_BORDER
$sep1b.Size       = New-Object System.Drawing.Size(350, 1)
$sep1b.Location   = New-Object System.Drawing.Point(0, 169)
$sep1b.Anchor     = $TLR
$form.Controls.Add($sep1b)

# ─── Load network info initially + refresh timer ───────────────────────────────
function Update-NetInfo {
    $ni = Get-NetworkInfo
    $lblNetSSID.Text    = $ni.SSID
    $lblNetAdapter.Text = $ni.Adapter

    if ($ni.Error -ne "") {
        $lblNetSpeed.Text   = "Err: $($ni.Error.Substring(0, [Math]::Min(30, $ni.Error.Length)))"
        $lblNetIP.Text      = $DASH
        $lblNetGateway.Text = $DASH
        $lblNetDNS.Text     = $DASH
    } else {
        if ($null -ne $ni.SignalQuality) {
            $bar = switch ($true) {
                ($ni.SignalQuality -ge 80) { "||||" }
                ($ni.SignalQuality -ge 60) { "|||." }
                ($ni.SignalQuality -ge 40) { "||.." }
                ($ni.SignalQuality -ge 20) { "|..." }
                default                   { "...." }
            }
            $lblNetSpeed.Text = "$bar $($ni.SignalQuality)%  $($ni.Speed)"
        } else {
            $lblNetSpeed.Text = $ni.Speed
        }
        $lblNetIP.Text      = $ni.LocalIP
        $lblNetGateway.Text = $ni.Gateway
        $lblNetDNS.Text     = $ni.DNS
    }
}
# Defer heavy network lookups to after the window is visible
# This lets the form appear instantly with placeholder values ("—")
# Update-NetInfo and initial route detection run ~300ms after shown

$netRefreshTimer          = New-Object System.Windows.Forms.Timer
$netRefreshTimer.Interval = 15000
$netRefreshTimer.Add_Tick({ Update-NetInfo })
# Don't start yet — started by deferred init below

$lblNetPubIP.Text = "..."
$lblNetOrt.Text   = "..."

$script:geoCheckTimer          = New-Object System.Windows.Forms.Timer
$script:geoCheckTimer.Interval = 500
$script:geoCheckTimer.Add_Tick({
    Check-GeoJob
    if ($script:GeoChecked -and -not $script:GeoFailed) {
        $script:geoCheckTimer.Stop()
        $script:geoCheckTimer.Dispose()
        $script:geoCheckTimer = $null
    }
})
# Don't start yet — started by deferred init below

$bwTimer          = New-Object System.Windows.Forms.Timer
$bwTimer.Interval = 2000
$bwTimer.Add_Tick({ Update-Bandwidth })
# Don't start yet — started by deferred init below

# ─── Detect network changes automatically ──────────────────────────────────────
$script:LastNetChangeAt = [datetime]::MinValue
$script:LastKnownIP     = ""
$script:LastKnownGW     = ""

# Event registration is cheap — keep synchronous
$null = Register-ObjectEvent `
    -InputObject ([System.Net.NetworkInformation.NetworkChange]) `
    -EventName   "NetworkAddressChanged" `
    -SourceIdentifier "PingMon.NetworkChanged"

# ─── Deferred startup: run heavy init after form is visible ────────────────────
$script:startupTimer          = New-Object System.Windows.Forms.Timer
$script:startupTimer.Interval = 300
$script:startupTimer.Add_Tick({
    $script:startupTimer.Stop()
    $script:startupTimer.Dispose()
    $script:startupTimer = $null

    # Now do the heavy lookups
    Update-NetInfo

    # Initial route for change detection
    $_initRoute = Get-PhysicalAdapterRoute
    if ($_initRoute) {
        $_initIP = Get-NetIPAddress -InterfaceIndex $_initRoute.InterfaceIndex -AddressFamily IPv4 `
                       -ErrorAction SilentlyContinue | Select-Object -First 1
        $script:LastKnownIP = if ($_initIP) { $_initIP.IPAddress } else { "" }
        $script:LastKnownGW = if ($_initRoute.NextHop) { $_initRoute.NextHop } else { "" }
    }

    # Start GeoIP lookup
    Start-GeoLookup
    $script:geoCheckTimer.Start()

    # Start periodic timers
    $netRefreshTimer.Start()
    $bwTimer.Start()
})
$script:startupTimer.Start()

$script:LastNetRefreshAt = [datetime]::MinValue

function Trigger-NetworkRefresh {
    # Debounce: don't trigger more than once every 10 seconds
    $now = [datetime]::UtcNow
    if (($now - $script:LastNetRefreshAt).TotalSeconds -lt 10) { return }
    $script:LastNetRefreshAt = $now

    Reset-PhysicalAdapterCache
    $script:BwIfIdx    = -1
    $script:BwLastTime = [datetime]::MinValue
    $script:BwLastRx   = 0
    $script:BwLastTx   = 0

    Update-NetInfo

    # Clean up previous delay timer if it exists
    if ($null -ne $script:netRefreshDelay) {
        try { $script:netRefreshDelay.Stop(); $script:netRefreshDelay.Dispose() } catch {}
        $script:netRefreshDelay = $null
    }

    $script:netRefreshDelay          = New-Object System.Windows.Forms.Timer
    $script:netRefreshDelay.Interval = 3000
    $script:netRefreshDelay.Add_Tick({
        if ($null -ne $script:netRefreshDelay) {
            try { $script:netRefreshDelay.Stop(); $script:netRefreshDelay.Dispose() } catch {}
            $script:netRefreshDelay = $null
        }
        Update-NetInfo
        $script:BwIfIdx    = -1
        $script:BwLastTime = [datetime]::MinValue
        $script:BwLastRx   = 0
        $script:BwLastTx   = 0
    })
    $script:netRefreshDelay.Start()

    $script:GeoChecked = $false
    if ($null -ne $script:GeoJob) {
        try { Remove-Job -Job $script:GeoJob -Force -ErrorAction SilentlyContinue } catch {}
        $script:GeoJob = $null
    }
    $lblNetPubIP.Text = "re-fetching..."
    $lblNetOrt.Text   = "re-fetching..."
    Start-GeoLookup

    if ($null -ne $script:geoRecheck) {
        try { $script:geoRecheck.Stop(); $script:geoRecheck.Dispose() } catch {}
        $script:geoRecheck = $null
    }
    $script:geoRecheck          = New-Object System.Windows.Forms.Timer
    $script:geoRecheck.Interval = 1000
    $script:geoRecheck.Add_Tick({
        Check-GeoJob
        if ($script:GeoChecked -and -not $script:GeoFailed) {
            $script:geoRecheck.Stop()
            $script:geoRecheck.Dispose()
            $script:geoRecheck = $null
        }
    })
    $script:geoRecheck.Start()
}

$netChangeTimer          = New-Object System.Windows.Forms.Timer
$netChangeTimer.Interval = 5000
$netChangeTimer.Add_Tick({
    $pending = @(Get-Event -SourceIdentifier "PingMon.NetworkChanged" -ErrorAction SilentlyContinue)
    $hadEvents = $pending.Count -gt 0
    foreach ($ev in $pending) {
        Remove-Event -EventIdentifier $ev.EventIdentifier -ErrorAction SilentlyContinue
    }
    if (-not $hadEvents) { return }

    try {
        # FIX: Use Get-PhysicalAdapterRoute for change detection too
        $route = Get-PhysicalAdapterRoute

        if ($null -eq $route) {
            if ($script:LastKnownIP -ne "" -or $script:LastKnownGW -ne "") {
                $script:LastKnownIP = ""
                $script:LastKnownGW = ""
            }
            return
        }

        $ipCfg     = Get-NetIPAddress -InterfaceIndex $route.InterfaceIndex -AddressFamily IPv4 `
                         -ErrorAction SilentlyContinue | Select-Object -First 1
        $currentIP = if ($ipCfg) { $ipCfg.IPAddress } else { "" }
        $currentGW = if ($route.NextHop -and $route.NextHop -ne "0.0.0.0") { $route.NextHop } else { "" }

        if ($currentIP -eq "") { return }

        if ($currentIP -eq $script:LastKnownIP -and $currentGW -eq $script:LastKnownGW) { return }

        $script:LastKnownIP = $currentIP
        $script:LastKnownGW = $currentGW
    } catch { return }

    Trigger-NetworkRefresh
})
$netChangeTimer.Start()

$lblHostCap       = New-Object System.Windows.Forms.Label
$lblHostCap.Text  = "TARGET"; $lblHostCap.Font = $FONT_SMALL; $lblHostCap.ForeColor = $C_MUTED
$lblHostCap.Size  = New-Object System.Drawing.Size(120, 14); $lblHostCap.Location = New-Object System.Drawing.Point(12, 177)
$lblHostCap.Anchor= $TL
$form.Controls.Add($lblHostCap)

$lblIntCap        = New-Object System.Windows.Forms.Label
$lblIntCap.Text   = "INTERVAL (s)"; $lblIntCap.Font = $FONT_SMALL; $lblIntCap.ForeColor = $C_MUTED
$lblIntCap.Size   = New-Object System.Drawing.Size(80, 14); $lblIntCap.Location = New-Object System.Drawing.Point(258, 177)
$lblIntCap.Anchor = $TR
$form.Controls.Add($lblIntCap)

$txtHost          = New-Object System.Windows.Forms.TextBox
$txtHost.Text     = "8.8.8.8"; $txtHost.Font = $FONT_MONO
$txtHost.BackColor= $C_SURFACE; $txtHost.ForeColor = $C_TEXT
$txtHost.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtHost.Size     = New-Object System.Drawing.Size(210, 24); $txtHost.Location = New-Object System.Drawing.Point(12, 192)
$txtHost.Anchor   = $TLR
$form.Controls.Add($txtHost)

$numInt           = New-Object System.Windows.Forms.NumericUpDown
$numInt.Minimum   = 1; $numInt.Maximum = 60; $numInt.Value = 2
$numInt.Font      = $FONT_MONO; $numInt.BackColor = $C_SURFACE; $numInt.ForeColor = $C_TEXT
$numInt.Size      = New-Object System.Drawing.Size(46, 24); $numInt.Location = New-Object System.Drawing.Point(258, 192)
$numInt.Anchor    = $TR
$form.Controls.Add($numInt)

# Dynamic interval checkbox — right of interval input
$chkDynInt          = New-Object System.Windows.Forms.CheckBox
$chkDynInt.Text     = "↕"
$chkDynInt.Font     = New-Object System.Drawing.Font("Segoe UI", 8)
$chkDynInt.ForeColor= $C_MUTED
$chkDynInt.Size     = New-Object System.Drawing.Size(30, 24)
$chkDynInt.Location = New-Object System.Drawing.Point(312, 192)
$chkDynInt.Anchor   = $TR
$chkDynInt.Cursor   = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($chkDynInt)

$chkDynInt.Add_CheckedChanged({
    $script:DynIntEnabled   = $chkDynInt.Checked
    $script:DynStableStreak = 0
    $script:DynLastTier     = 1
    if (-not $chkDynInt.Checked -and $script:Running) {
        $script:DynIntCurrent   = [double]$numInt.Value
        $pingTimer.Interval     = [int]($numInt.Value * 1000)
        $lblStatus.Text         = "Monitoring: $($txtHost.Text)"
    }
})

# ─── Stat cards via TableLayoutPanel ───────────────────────────────────────────
$statsTable       = New-Object System.Windows.Forms.TableLayoutPanel
$statsTable.ColumnCount = 4; $statsTable.RowCount = 1
$statsTable.Size  = New-Object System.Drawing.Size(326, 70)
$statsTable.Location = New-Object System.Drawing.Point(12, 220)
$statsTable.Anchor   = $TLR
$statsTable.BackColor= $C_BG
$statsTable.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$statsTable.Padding = New-Object System.Windows.Forms.Padding(0)
[void]$statsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 29)))
[void]$statsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 29)))
[void]$statsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 27)))
[void]$statsTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 15)))
[void]$statsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

function New-StatCardPanel($labelText, $unitText) {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.BackColor   = $C_SURFACE2
    $panel.Dock        = [System.Windows.Forms.DockStyle]::Fill
    $panel.Margin      = New-Object System.Windows.Forms.Padding(0, 0, 4, 0)
    $panel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $labelText; $lbl.Font = $FONT_SMALL; $lbl.ForeColor = $C_MUTED
    $lbl.Dock = [System.Windows.Forms.DockStyle]::Top; $lbl.Height = 18
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::BottomCenter
    $panel.Controls.Add($lbl)

    $val = New-Object System.Windows.Forms.Label
    $val.Text = $DASH; $val.Font = $FONT_MONO_B; $val.ForeColor = $C_TEXT
    $val.Dock = [System.Windows.Forms.DockStyle]::Fill
    $val.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $panel.Controls.Add($val)

    $unit = New-Object System.Windows.Forms.Label
    $unit.Text = $unitText; $unit.Font = $FONT_SMALL; $unit.ForeColor = $C_MUTED
    $unit.Dock = [System.Windows.Forms.DockStyle]::Bottom; $unit.Height = 16
    $unit.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
    $panel.Controls.Add($unit)

    return @{ Panel=$panel; Val=$val }
}

$cardCur    = New-StatCardPanel "PING"    "ms"
$cardJitter = New-StatCardPanel "JITTER"  "ms"
$cardLoss   = New-StatCardPanel "LOSS" "%"
$cardScore  = New-StatCardPanel "SCORE"   ""
$cardScore.Val.Font = $FONT_SCORE
$cardScore.Val.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
$cardScore.Panel.Margin = New-Object System.Windows.Forms.Padding(0)

$statsTable.Controls.Add($cardCur.Panel,    0, 0)
$statsTable.Controls.Add($cardJitter.Panel, 1, 0)
$statsTable.Controls.Add($cardLoss.Panel,   2, 0)
$statsTable.Controls.Add($cardScore.Panel,  3, 0)
$form.Controls.Add($statsTable)

$lblScoreDesc         = New-Object System.Windows.Forms.Label
$lblScoreDesc.Text    = ""
$lblScoreDesc.Font    = $FONT_SMALL; $lblScoreDesc.ForeColor = $C_MUTED
$lblScoreDesc.Size    = New-Object System.Drawing.Size(326, 14)
$lblScoreDesc.Location= New-Object System.Drawing.Point(12, 293)
$lblScoreDesc.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblScoreDesc.Anchor  = $TLR
$form.Controls.Add($lblScoreDesc)

# ─── Chart ─────────────────────────────────────────────────────────────────────
$chartPanel           = New-Object System.Windows.Forms.Panel
$chartPanel.BackColor = [System.Drawing.Color]::FromArgb(10, 10, 14)
$chartPanel.Size      = New-Object System.Drawing.Size(326, 80)
$chartPanel.Location  = New-Object System.Drawing.Point(12, 309)
$chartPanel.Anchor    = $TLR
$chartPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$lblChart             = New-Object System.Windows.Forms.Label
$lblChart.Text        = "PING HISTORY  (last $($script:MaxHist) measurements)"
$lblChart.Font        = $FONT_SMALL; $lblChart.ForeColor = $C_MUTED
$lblChart.BackColor   = [System.Drawing.Color]::Transparent
$lblChart.Dock        = [System.Windows.Forms.DockStyle]::Top
$lblChart.Height      = 16; $lblChart.Padding = New-Object System.Windows.Forms.Padding(4,2,0,0)
$chartPanel.Controls.Add($lblChart)

$chartCanvas          = New-Object System.Windows.Forms.Panel
$chartCanvas.BackColor= [System.Drawing.Color]::FromArgb(10, 10, 14)
$chartCanvas.Dock     = [System.Windows.Forms.DockStyle]::Fill

$chartCanvas.Add_Paint({
    param($s, $e)
    $g = $e.Graphics; $cw = $chartCanvas.Width; $ch = $chartCanvas.Height; $n = $script:MaxHist
    $valid = @($script:History | Where-Object { -not $_.Timeout } | ForEach-Object { $_.Ms })
    $maxV  = if ($valid.Count -gt 0) { [Math]::Max(($valid | Measure-Object -Maximum).Maximum, 100) } else { 100 }
    for ($i = 0; $i -lt $script:History.Count; $i++) {
        $entry = $script:History[$i]
        $bx = [int]($i * ($cw / $n))
        $bw = [Math]::Max(1, [int](($i + 1) * ($cw / $n)) - $bx - 1)
        if ($entry.Timeout) {
            $br = New-Object System.Drawing.SolidBrush($C_TIMEOUT)
            $g.FillRectangle($br, $bx, 0, $bw, $ch); $br.Dispose()
        } else {
            $bh = [Math]::Max(2, [int](($entry.Ms / $maxV) * ($ch - 4)))
            $by = $ch - $bh
            $d  = New-Object System.Drawing.SolidBrush((Get-PingColorDim $entry.Ms))
            $t  = New-Object System.Drawing.SolidBrush((Get-PingColor    $entry.Ms))
            $g.FillRectangle($d, $bx, $by, $bw, $bh)
            $g.FillRectangle($t, $bx, $by, $bw, 2)
            $d.Dispose(); $t.Dispose()
        }
    }
})
$chartPanel.Controls.Add($chartCanvas)
$form.Controls.Add($chartPanel)

# ─── Legend ────────────────────────────────────────────────────────────────────
$legendPanel          = New-Object System.Windows.Forms.Panel
$legendPanel.BackColor= $C_BG
$legendPanel.Size     = New-Object System.Drawing.Size(326, 18)
$legendPanel.Location = New-Object System.Drawing.Point(12, 391)
$legendPanel.Anchor   = $TLR

function Add-LegendItem($panel, $lx, $color, $text) {
    $dot = New-Object System.Windows.Forms.Panel
    $dot.BackColor = $color; $dot.Size = New-Object System.Drawing.Size(8, 8)
    $dot.Location  = New-Object System.Drawing.Point($lx, 5)
    $panel.Controls.Add($dot)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text; $lbl.Font = $FONT_SMALL; $lbl.ForeColor = $C_MUTED
    $lbl.Size = New-Object System.Drawing.Size(72, 14)
    $lblX = $lx + 11
    $lbl.Location = New-Object System.Drawing.Point($lblX, 2)
    $panel.Controls.Add($lbl)
}
Add-LegendItem $legendPanel   0 $C_GOOD   "< 80 ms"
Add-LegendItem $legendPanel  82 $C_WARN   "80-200 ms"
Add-LegendItem $legendPanel 168 $C_BAD    "> 200 ms"
Add-LegendItem $legendPanel 254 $C_TO_FG  "Timeout"
$form.Controls.Add($legendPanel)

# ─── Start/Stop button ─────────────────────────────────────────────────────────
$btnToggle        = New-Object System.Windows.Forms.Button
$btnToggle.Text   = "▶  Start"
$btnToggle.Font   = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnToggle.ForeColor = $C_GOOD
$btnToggle.BackColor = [System.Drawing.Color]::FromArgb(12, 40, 18)
$btnToggle.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnToggle.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 90, 40)
$btnToggle.FlatAppearance.BorderSize  = 1
$btnToggle.Size   = New-Object System.Drawing.Size(326, 30)
$btnToggle.Location = New-Object System.Drawing.Point(12, 412)
$btnToggle.Anchor = $TLR; $btnToggle.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnToggle)

# ─── Log ───────────────────────────────────────────────────────────────────────
$lblLog           = New-Object System.Windows.Forms.Label
$lblLog.Text      = "LOG"; $lblLog.Font = $FONT_SMALL; $lblLog.ForeColor = $C_MUTED
$lblLog.Size      = New-Object System.Drawing.Size(326, 14); $lblLog.Location = New-Object System.Drawing.Point(12, 445)
$lblLog.Anchor    = $TLR
$form.Controls.Add($lblLog)

$logBox           = New-Object System.Windows.Forms.RichTextBox
$logBox.BackColor = [System.Drawing.Color]::FromArgb(10, 10, 14)
$logBox.ForeColor = $C_TEXT; $logBox.Font = $FONT_MONO; $logBox.ReadOnly = $true
$logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logBox.Size      = New-Object System.Drawing.Size(326, 198)
$logBox.Location  = New-Object System.Drawing.Point(12, 460)
$logBox.Anchor    = $TLRB
$logBox.ScrollBars= [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$logBox.WordWrap  = $false
$form.Controls.Add($logBox)

# ─── Status bar ────────────────────────────────────────────────────────────────
$sep2 = New-Object System.Windows.Forms.Panel
$sep2.BackColor = $C_BORDER; $sep2.Size = New-Object System.Drawing.Size(350, 1)
$sep2.Location  = New-Object System.Drawing.Point(0, 662); $sep2.Anchor = $BLR
$form.Controls.Add($sep2)

$lblStatus        = New-Object System.Windows.Forms.Label
$lblStatus.Text   = "Ready"; $lblStatus.Font = $FONT_SMALL; $lblStatus.ForeColor = $C_MUTED
$lblStatus.Size   = New-Object System.Drawing.Size(210, 18); $lblStatus.Location = New-Object System.Drawing.Point(12, 665)
$lblStatus.Anchor = $BL
$form.Controls.Add($lblStatus)

$lblMinMax        = New-Object System.Windows.Forms.Label
$lblMinMax.Text   = "Min: -  /  Max: - ms"; $lblMinMax.Font = $FONT_SMALL; $lblMinMax.ForeColor = $C_MUTED
$lblMinMax.Size   = New-Object System.Drawing.Size(160, 18); $lblMinMax.Location = New-Object System.Drawing.Point(178, 665)
$lblMinMax.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $lblMinMax.Anchor = $BR
$form.Controls.Add($lblMinMax)

# ─── System tray icon ──────────────────────────────────────────────────────────
$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip

$trayMenuShow = New-Object System.Windows.Forms.ToolStripMenuItem
$trayMenuShow.Text = "Show"
$trayMenuShow.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$trayMenuShow.Add_Click({
    $form.Show()
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $form.Activate()
})
[void]$trayMenu.Items.Add($trayMenuShow)
[void]$trayMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

$trayMenuExit = New-Object System.Windows.Forms.ToolStripMenuItem
$trayMenuExit.Text = "Exit"
$trayMenuExit.Add_Click({ $form.Close() })
[void]$trayMenu.Items.Add($trayMenuExit)

$trayIcon                  = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.ContextMenuStrip = $trayMenu
$trayIcon.Text             = "Ping Monitor"
$trayIcon.Icon             = New-ScoreIcon "?" $C_MUTED
$trayIcon.Visible          = $true

$trayIcon.Add_DoubleClick({
    $form.Show()
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $form.Activate()
})

$form.Add_Resize({
    if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
        $form.Hide()
    }
})

# ─── Ping timer ───────────────────────────────────────────────────────────────
$pingTimer          = New-Object System.Windows.Forms.Timer
$pingTimer.Interval = 2000

# ─── Write log ─────────────────────────────────────────────────────────────────
function Write-PingLog($ms, $jitter) {
    $time    = (Get-Date).ToString("HH:mm:ss")
    $isTO    = ($null -eq $ms)
    $msText  = if ($isTO) { "TIMEOUT" } else { $ms.ToString().PadLeft(5) + " ms" }
    $jitText = if ($null -eq $jitter -or $isTO) { "        " } else { " j:" + $jitter.ToString().PadLeft(3) + "ms" }

    $logBox.SelectionStart = $logBox.TextLength; $logBox.SelectionLength = 0
    $logBox.SelectionColor = $C_MUTED;           $logBox.AppendText("$time  ")
    $logBox.SelectionColor = if ($isTO) { $C_TO_FG } elseif ($ms -lt 80) { $C_GOOD } elseif ($ms -lt 200) { $C_WARN } else { $C_BAD }
    $logBox.AppendText($msText)
    $logBox.SelectionColor = $C_MUTED
    $logBox.AppendText($jitText + "  " + $txtHost.Text + "`n")

    if ($isTO) {
        $script:ConsecTimeouts++
        if ($script:ConsecTimeouts -ge 3 -and ($script:ConsecTimeouts % 3 -eq 0)) {
            $logBox.SelectionStart = $logBox.TextLength; $logBox.SelectionLength = 0
            $logBox.SelectionColor = $C_MUTED
            $lossNow = if ($script:Sent -gt 0) { [Math]::Round(($script:Lost/$script:Sent)*100) } else { 0 }
            $dbg = "         [DBG] $($script:ConsecTimeouts)x Timeout | Sent=$($script:Sent) Lost=$($script:Lost) Loss=${lossNow}% | Pending=$($script:PingPending)`n"
            $logBox.AppendText($dbg)
        }
    } else {
        if ($script:ConsecTimeouts -ge 3) {
            $logBox.SelectionStart = $logBox.TextLength; $logBox.SelectionLength = 0
            $logBox.SelectionColor = $C_GOOD
            $logBox.AppendText("         [DBG] Connection restored after $($script:ConsecTimeouts)x timeout`n")
        }
        $script:ConsecTimeouts = 0
    }

    $logBox.SelectionStart = $logBox.TextLength; $logBox.ScrollToCaret()

    if ($logBox.Lines.Count -gt $script:MaxLog) {
        $keep     = [Math]::Max($script:MaxLog - 20, 60)
        $lines    = $logBox.Lines
        $newText  = ($lines | Select-Object -Last $keep) -join "`n"
        $logBox.Text = $newText + "`n"
        $logBox.SelectionStart = $logBox.TextLength
        $logBox.ScrollToCaret()
    }
}

# ─── UI nach Messung aktualisieren ────────────────────────────────────────────
function Update-Stats($ms) {
    $script:Sent++
    if ($null -eq $ms) { $script:Lost++ }

    $script:History.Add([PSCustomObject]@{ Ms=$ms; Timeout=($null -eq $ms) })
    if ($script:History.Count -gt $script:MaxHist) { $script:History.RemoveAt(0) }

    $cardCur.Val.Text      = if ($null -eq $ms) { $DASH } else { "$ms" }
    $cardCur.Val.ForeColor = Get-PingColor $ms

    $jitter = Get-Jitter
    if ($null -ne $jitter) {
        $cardJitter.Val.Text      = "$jitter"
        $cardJitter.Val.ForeColor = if ($jitter -lt 15) { $C_GOOD } elseif ($jitter -lt 40) { $C_WARN } else { $C_BAD }
    } else { $cardJitter.Val.Text = $DASH; $cardJitter.Val.ForeColor = $C_TEXT }

    $lossP = if ($script:Sent -gt 0) { [Math]::Round(($script:Lost / $script:Sent) * 100) } else { 0 }
    $cardLoss.Val.Text      = "$lossP"
    $cardLoss.Val.ForeColor = if ($lossP -eq 0) { $C_GOOD } elseif ($lossP -lt 5) { $C_WARN } else { $C_BAD }

    $valid = @($script:History | Where-Object { -not $_.Timeout } | ForEach-Object { $_.Ms })
    if ($valid.Count -gt 0) {
        $mn = ($valid | Measure-Object -Minimum).Minimum
        $mx = ($valid | Measure-Object -Maximum).Maximum
        $lblMinMax.Text = "Min: $mn  /  Max: $mx ms"
    }

    $q = Get-QualityScore
    $cardScore.Val.Text        = $q.Grade
    $cardScore.Val.ForeColor   = $q.Color
    $lblScoreDesc.Text         = $q.Desc
    $lblScoreDesc.ForeColor    = $q.Color

    if ($q.Grade -ne $script:LastGrade) {
        $script:LastGrade = $q.Grade
        $newIcon  = New-ScoreIcon $q.Grade $q.Color
        $oldIcon  = $trayIcon.Icon
        $trayIcon.Icon = $newIcon
        if ($null -ne $oldIcon) { try { $oldIcon.Dispose() } catch {} }
    }
    $lossP2 = if ($script:Sent -gt 0) { [Math]::Round(($script:Lost / $script:Sent) * 100) } else { 0 }
    $curMs  = if ($null -eq $ms) { "Timeout" } else { "${ms}ms" }
    $intText = if ($script:DynIntEnabled) { "$($script:DynIntCurrent)s↕" } else { "$([int]$numInt.Value)s" }
    $trayIcon.Text = "Ping Monitor  |  $($q.Grade)  |  $curMs  |  Loss: ${lossP2}%  |  Int: $intText"

    Write-PingLog $ms $jitter
    $chartCanvas.Invalidate()
    $script:PingPending = $false

    # Adjust ping frequency if dynamic mode is enabled
    Update-DynamicInterval $q.Grade
}

# ─── Ping via .NET Ping.SendAsync ──────────────────────────────────────────────
function Send-Ping {
    if (-not $script:Running) { return }

    if ($script:PingPending) {
        $elapsed = ([datetime]::UtcNow - $script:PingSentAt).TotalSeconds
        if ($elapsed -gt 4) {
            $logBox.SelectionStart = $logBox.TextLength; $logBox.SelectionLength = 0
            $logBox.SelectionColor = $C_TO_FG
            $logBox.AppendText("         [WDG] PingPending seit ${elapsed}s haengt — erzwinge Reset`n")
            $logBox.SelectionStart = $logBox.TextLength; $logBox.ScrollToCaret()
            try { $script:Pinger.SendAsyncCancel() } catch {}
            $script:PingPending = $false
            Update-Stats $null
        }
        return
    }

    $script:PingPending = $true
    $script:PingSentAt  = [datetime]::UtcNow

    $dest    = $txtHost.Text.Trim()
    $pinger  = $script:Pinger

    try {
        $pinger.SendAsync($dest, 2000, $script:PingBuffer, $script:PingOpts, $null)
    } catch {
        $logBox.SelectionStart = $logBox.TextLength; $logBox.SelectionLength = 0
        $logBox.SelectionColor = $C_TO_FG
        $logBox.AppendText("         [ERR] SendAsync: $($_.Exception.Message)`n")
        $logBox.SelectionStart = $logBox.TextLength; $logBox.ScrollToCaret()
        $script:PingPending = $false
        Update-Stats $null
    }
}

# ─── Register PingCompleted event ──────────────────────────────────────────────
$pingCompletedHandler = {
    param($evSender, $evArgs)
    if (-not $script:Running) { $script:PingPending = $false; return }
    $ms = $null
    try {
        if ($evArgs.Cancelled) { $script:PingPending = $false; return }

        $reply = $evArgs.Reply
        if ($null -ne $evArgs.Error) {
            $logBox.SelectionStart = $logBox.TextLength; $logBox.SelectionLength = 0
            $logBox.SelectionColor = $C_TO_FG
            $logBox.AppendText("         [ERR] $($evArgs.Error.Message)`n")
            $logBox.SelectionStart = $logBox.TextLength; $logBox.ScrollToCaret()
        } elseif ($null -ne $reply -and $reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
            $ms = [int]$reply.RoundtripTime
            if ($ms -eq 0) { $ms = 1 }
        }
    } catch {
        $logBox.SelectionStart = $logBox.TextLength; $logBox.SelectionLength = 0
        $logBox.SelectionColor = $C_TO_FG
        $logBox.AppendText("         [EXC] $($_.Exception.Message)`n")
        $logBox.SelectionStart = $logBox.TextLength; $logBox.ScrollToCaret()
    }
    Update-Stats $ms
}

$script:Pinger.add_PingCompleted($pingCompletedHandler)

$pingTimer.Add_Tick({ Send-Ping })

# ─── Tooltips ──────────────────────────────────────────────────────────────────
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 8000
$toolTip.InitialDelay = 400
$toolTip.ReshowDelay  = 200
$toolTip.BackColor    = [System.Drawing.Color]::FromArgb(30, 30, 36)
$toolTip.ForeColor    = [System.Drawing.Color]::FromArgb(220, 220, 228)

$toolTip.SetToolTip($cardCur.Panel,    "Round-trip time to target in milliseconds")
$toolTip.SetToolTip($cardJitter.Panel, "Variation between consecutive pings — lower is more stable")
$toolTip.SetToolTip($cardLoss.Panel,   "Packet loss percentage over the last $($script:MaxHist) measurements")
$toolTip.SetToolTip($cardScore.Panel,  "Connection quality: A (excellent) to D (poor), X = no connection")
$toolTip.SetToolTip($chkDynInt,        "Auto-adjust ping interval based on network quality:`nStable → slower polling  |  Unstable → faster polling")
$toolTip.SetToolTip($numInt,           "Base ping interval in seconds (1–60)")
$toolTip.SetToolTip($txtHost,          "Target IP address or hostname to ping")
$toolTip.SetToolTip($chartPanel,       "Ping history — last $($script:MaxHist) measurements")
$toolTip.SetToolTip($btnTheme,         "Toggle dark / light theme")
$toolTip.SetToolTip($trkOpacity,       "Window transparency (20–100%)")
$toolTip.SetToolTip($netPanel,         "Network adapter info — refreshes every 15 seconds")

# Theme einmalig beim Start anwenden
Apply-Theme

# ─── Start/Stop button ─────────────────────────────────────────────────────────
$btnToggle.Add_Click({
    if (-not $script:Running) {
        $script:Running     = $true
        $script:Sent        = 0
        $script:Lost        = 0
        $script:PingPending  = $false
        $script:BwSessionRx  = 0
        $script:BwSessionTx  = 0
        $script:BwLastTime   = [datetime]::MinValue
        $script:History.Clear()
        $logBox.Clear()

        foreach ($c in @($cardCur, $cardJitter, $cardLoss, $cardScore)) {
            $c.Val.Text = $DASH; $c.Val.ForeColor = $C_TEXT
        }
        $lblScoreDesc.Text = ""; $lblMinMax.Text = "Min: -  /  Max: - ms"
        if ($script:DynIntEnabled) {
            $lblStatus.Text = "Monitoring: $($txtHost.Text)  ·  $($numInt.Value)s ="
        } else {
            $lblStatus.Text = "Monitoring: $($txtHost.Text)"
        }
        $lblStatus.ForeColor = $C_GOOD

        $btnToggle.Text = "■  Stop"; $btnToggle.ForeColor = $C_BAD
        $btnToggle.BackColor = [System.Drawing.Color]::FromArgb(50, 10, 10)
        $btnToggle.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(100, 20, 20)

        $txtHost.Enabled = $false; $numInt.Enabled = $false
        $script:DynIntCurrent   = [double]$numInt.Value
        $script:DynStableStreak = 0
        $script:DynLastTier     = 1
        $pingTimer.Interval = [int]($numInt.Value * 1000)
        $pingTimer.Start()
        Send-Ping

    } else {
        $script:Running = $false
        $pingTimer.Stop()
        try { $script:Pinger.SendAsyncCancel() } catch {}
        $script:PingPending = $false

        $btnToggle.Text = "▶  Start"; $btnToggle.ForeColor = $C_GOOD
        $btnToggle.BackColor = [System.Drawing.Color]::FromArgb(12, 40, 18)
        $btnToggle.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 90, 40)

        $txtHost.Enabled = $true; $numInt.Enabled = $true
        $script:DynIntCurrent   = [double]$numInt.Value
        $script:DynStableStreak = 0
        $lblStatus.Text = "Stopped"; $lblStatus.ForeColor = $C_MUTED

        $script:LastGrade = ""
        $oldIcon = $trayIcon.Icon
        $trayIcon.Icon = New-ScoreIcon "?" $C_MUTED
        $trayIcon.Text = "Ping Monitor  |  Stopped"
        try { $oldIcon.Dispose() } catch {}
    }
})

# ─── Cleanup ───────────────────────────────────────────────────────────────────
$form.Add_FormClosing({
    $script:Running = $false
    $pingTimer.Stop()
    $netChangeTimer.Stop()
    try { Unregister-Event -SourceIdentifier "PingMon.NetworkChanged" -ErrorAction SilentlyContinue } catch {}
    try { $script:Pinger.SendAsyncCancel() } catch {}
    if ($null -ne $script:GeoJob) {
        try { Remove-Job -Job $script:GeoJob -Force -ErrorAction SilentlyContinue } catch {}
    }
    if ($null -ne $script:geoCheckTimer) {
        try { $script:geoCheckTimer.Stop() } catch {}
    }
    if ($null -ne $script:geoRecheck) {
        try { $script:geoRecheck.Stop() } catch {}
    }
    if ($null -ne $script:netRefreshDelay) {
        try { $script:netRefreshDelay.Stop() } catch {}
    }
    if ($null -ne $script:startupTimer) {
        try { $script:startupTimer.Stop(); $script:startupTimer.Dispose() } catch {}
        $script:startupTimer = $null
    }
})
$form.Add_FormClosed({
    $pingTimer.Dispose()
    $bwTimer.Stop(); $bwTimer.Dispose()
    $netChangeTimer.Dispose()
    $netRefreshTimer.Stop(); $netRefreshTimer.Dispose()
    try { $script:Pinger.remove_PingCompleted($pingCompletedHandler) } catch {}
    $script:Pinger.Dispose()
    try { $trayIcon.Visible = $false } catch {}
    try { $trayIcon.Icon.Dispose() } catch {}
    $trayIcon.Dispose()
    $FONT_MONO.Dispose(); $FONT_MONO_B.Dispose(); $FONT_SCORE.Dispose()
    $FONT_SMALL.Dispose(); $FONT_UI.Dispose()
    $toolTip.Dispose()
})

[System.Windows.Forms.Application]::Run($form)
