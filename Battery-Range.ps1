<#
.SYNOPSIS
    Battery Range - Smart battery charging manager for Windows laptops

.DESCRIPTION
    This PowerShell script manages laptop battery charging within a specified
    state of charge (SoC) range using a Tasmota-compatible smart plug.
    It helps extend battery lifespan by avoiding constant full charges.

    Features:
    - Automatic battery management within configurable SoC ranges
    - Normal and High SoC range profiles
    - One-time charge/discharge to specific target
    - Manual charger control override
    - System tray integration with context menu
    - Failure notifications with manual action prompts
    - Intelligent standby handling with configurable SoC threshold
    - Optional Wi-Fi Hotspot management for smart plug connectivity

.NOTES
    File Name      : Battery-Range.ps1
    Prerequisite   : PowerShell 5.1+, Tasmota-compatible smart plug

.LINK
    https://github.com/CosimoDamianoForina/Battery-Range
#>

<#
MIT License

Copyright (c) 2025 Cosimo Damiano Forina

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

##############################################################################
# Default Configuration ######################################################
##############################################################################

[CmdletBinding()]
param (
    [Parameter(HelpMessage = "IP address of the smart plug")]
    [string]$SmartplugIP = "192.168.137.101",

    [Parameter(HelpMessage = "Timeout in seconds for the smart plug")]
    [ValidateRange(1, 5)]
    [int]$SmartplugTimeoutSeconds = 2,

    [Parameter(HelpMessage = "Timeout in seconds for confirming charging state change")]
    [ValidateRange(1, 10)]
    [int]$ConfirmationTimeoutSeconds = 5,

    [Parameter(HelpMessage = "Battery check interval in seconds")]
    [ValidateRange(15, 300)]
    [int]$CheckIntervalSeconds = 30,

    [Parameter(HelpMessage = "Maximum SoC for Auto mode")]
    [ValidateRange(20, 100)]
    [int]$MaxBatteryLevel = 45,

    [Parameter(HelpMessage = "Minimum SoC for Auto mode")]
    [ValidateRange(19, 99)]
    [int]$MinBatteryLevel = 35,

    [Parameter(HelpMessage = "Maximum SoC for Auto High mode")]
    [ValidateRange(20, 100)]
    [int]$MaxBatteryLevelHigh = 80,

    [Parameter(HelpMessage = "Minimum SoC for Auto High mode")]
    [ValidateRange(19, 99)]
    [int]$MinBatteryLevelHigh = 70,

    [Parameter(HelpMessage = "Maximum SoC to turn on charger before standby (-1 to disable)")]
    [ValidateRange(-1, 30)]
    [int]$StandbyOnSoCThreshold = 15,

    [Parameter(HelpMessage = "Minimum SoC to turn off charger before standby (101 to disable)")]
    [ValidateRange(20, 101)]
    [int]$StandbyOffSoCThreshold = 30,

    [Parameter(HelpMessage = "Manage Windows Wi-Fi Hotspot for smart plug connectivity")]
    [bool]$ManageWiFiHotspot = $false,

    [Parameter(HelpMessage = "Timeout in seconds for activating Windows Wi-Fi Hotspot")]
    [ValidateRange(5, 120)]
    [int]$HotspotTimeoutSeconds = 30,

    [Parameter(HelpMessage = "Maximum number of attempts to check the smart plug is reachable")]
    [ValidateRange(2, 50)]
    [int]$SmartplugConnectionMaxAttempts = 10,

    [Parameter(HelpMessage = "Maximum number of attempts to communicate with the smartplug")]
    [ValidateRange(1, 10)]
    [int]$SmartplugCommunicationMaxAttempts = 3
)

##############################################################################

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Validate SoC thresholds
try {
    if ($MinBatteryLevel -ge $MaxBatteryLevel) {
        throw "MinBatteryLevel ($MinBatteryLevel) must be less than MaxBatteryLevel ($MaxBatteryLevel)"
    }
    if ($MinBatteryLevelHigh -ge $MaxBatteryLevelHigh) {
        throw "MinBatteryLevelHigh ($MinBatteryLevelHigh) must be less than MaxBatteryLevelHigh ($MaxBatteryLevelHigh)"
    }
    if ($StandbyOnSoCThreshold -ge $StandbyOffSoCThreshold) {
        throw "StandbyOnSoCThreshold ($StandbyOnSoCThreshold) must be less than StandbyOffSoCThreshold ($StandbyOffSoCThreshold)"
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        $_.Exception.Message,
        "Battery Range Configuration Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return
}

# Single Instance Check
$mutexName = "Global\BatteryRangePowerShellScript"
$script:mutex = $null
$script:hasHandle = $false
try {
    $script:mutex = New-Object System.Threading.Mutex($false, $mutexName)

    try {
        $script:hasHandle = $script:mutex.WaitOne(0, $false)
    }
    catch [System.Threading.AbandonedMutexException] {
        # Previous instance crashed - we can take over
        $script:hasHandle = $true
    }

    if (-not $script:hasHandle) {
        # Another instance is already running
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "Battery Range is already running!",
            "Battery Range",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null

        # Dispose mutex before exiting
        $script:mutex.Dispose()
        return
    }
}
catch {
    Write-Host "Error: Could not create mutex." -ForegroundColor Red
    return
}

# Set DPI awareness
if (-not ('DPI' -as [type])) {
    Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public class DPI {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
'@
}
[DPI]::SetProcessDPIAware() | out-null

# Load WinRT assemblies for Mobile Hotspot control (if needed)
if ($ManageWiFiHotspot) {
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        [void][Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager, Windows.Networking.NetworkOperators, ContentType=WindowsRuntime]
        [void][Windows.Networking.Connectivity.NetworkInformation, Windows.Networking.Connectivity, ContentType=WindowsRuntime]
    }
    catch {
        Write-Warning "Mobile Hotspot management requires Windows 10+. Disabling feature."
        $ManageWiFiHotspot = $false
    }
}

# Add icon helper for destroying icon handles
if (-not ('IconHelper' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class IconHelper {
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
'@
}

# Add power status helper for fast AC/battery detection
if (-not ('PowerStatus' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class PowerStatus {
    [StructLayout(LayoutKind.Sequential)]
    public struct SYSTEM_POWER_STATUS {
        public byte ACLineStatus;        // 0=Offline, 1=Online, 255=Unknown
        public byte BatteryFlag;         // Battery charge status flags
        public byte BatteryLifePercent;  // 0-100, or 255 if unknown
        public byte SystemStatusFlag;    // 0=Battery saver off, 1=on
        public int BatteryLifeTime;      // Seconds remaining, -1 if unknown
        public int BatteryFullLifeTime;  // Seconds for full battery, -1 if unknown
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetSystemPowerStatus(out SYSTEM_POWER_STATUS lpSystemPowerStatus);

    public static SYSTEM_POWER_STATUS GetStatus() {
        SYSTEM_POWER_STATUS status;
        if (!GetSystemPowerStatus(out status)) {
            throw new System.ComponentModel.Win32Exception();
        }
        return status;
    }
}
'@
}

function Invoke-WinRTAsync {
    param($AsyncTask, [Type]$ResultType)

    $asTaskMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $_.Name -eq 'AsTask' -and
            $_.GetParameters().Count -eq 1 -and
            $_.GetParameters()[0].ParameterType.Name -match 'IAsyncOperation`1'
        } | Select-Object -First 1

    if (-not $asTaskMethod) {
        throw "Could not find AsTask method for WinRT async operations"
    }

    $task = $asTaskMethod.MakeGenericMethod($ResultType).Invoke($null, @($AsyncTask))
    $task.Wait() | Out-Null
    return $task.Result
}

function Enable-MobileHotspot {
    $timestamp = Get-Date -Format "HH:mm:ss"

    try {
        $profile = [Windows.Networking.Connectivity.NetworkInformation]::GetInternetConnectionProfile()

        if (-not $profile) {
            Write-Host "[$timestamp] WARNING: No internet connection available for hotspot" -ForegroundColor Yellow
            return $false
        }

        $manager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager]::CreateFromConnectionProfile($profile)
        $state = $manager.TetheringOperationalState

        if ($state -eq [Windows.Networking.NetworkOperators.TetheringOperationalState]::On) {
            Write-Host "[$timestamp] Hotspot is already running" -ForegroundColor Green
            return $true
        }

        Write-Host "[$timestamp] Starting Mobile Hotspot..." -ForegroundColor White

        $result = Invoke-WinRTAsync -AsyncTask $manager.StartTetheringAsync() -ResultType ([Windows.Networking.NetworkOperators.NetworkOperatorTetheringOperationResult])

        if ($result.Status -eq [Windows.Networking.NetworkOperators.TetheringOperationStatus]::Success) {
            Write-Host "[$timestamp] Mobile Hotspot activated" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "[$timestamp] Failed to start hotspot: $($result.Status)" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "[$timestamp] Error enabling hotspot: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Initialize-Hotspot {
    if (-not $ManageWiFiHotspot) { return }

    $pollIntervalMs = 1000
    $maxIterations = [math]::Ceiling(($HotspotTimeoutSeconds * 1000) / $pollIntervalMs)

    for ($i = 1; $i -le $maxIterations; $i++) {
        if (Enable-MobileHotspot) {
            Write-Host "Checking smart plug availability..." -ForegroundColor Cyan
            $available = Confirm-SmartplugAvailable
            if ($available) {
                Write-Host "Smartplug detected" -ForegroundColor Green
            }
            else {
                Write-Host "WARNING: Could not connect to the smartplug after $SmartplugConnectionMaxAttempts attempts" -ForegroundColor Red
            }
            return
        }

        if ($i -lt $maxIterations) {
            Write-Host "Hotspot activation failed. Retrying in $pollIntervalMs ms... (attempt $i/$maxIterations)" -ForegroundColor Yellow
        }

        Start-Sleep -Milliseconds $pollIntervalMs
    }

    Write-Host "WARNING: Could not enable hotspot after $maxIterations attempts" -ForegroundColor Red
}

function SendWindowsNotification {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Title,
        [Parameter(Mandatory)]
        [string]$Message
    )

    try {
        # WinRT toast APIs can throw if not supported / app not registered for toasts.
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null

        $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(
            [Windows.UI.Notifications.ToastTemplateType]::ToastText02
        )

        $toastXml = [xml]$template.GetXml()
        $textNodes = $toastXml.GetElementsByTagName("text")
        $textNodes.Item(0).AppendChild($toastXml.CreateTextNode($Title)) | Out-Null
        $textNodes.Item(1).AppendChild($toastXml.CreateTextNode($Message)) | Out-Null

        $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $xml.LoadXml($toastXml.OuterXml)

        $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
        $toast.Tag = "BatteryRangeNotification"
        $toast.Group = $toast.Tag
        $toast.ExpirationTime = [DateTimeOffset]::Now.AddSeconds($CheckIntervalSeconds)

        # Note: On some systems toasts require proper desktop app registration/shortcut.
        $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Battery Range")
        $notifier.Show($toast)
    }
    catch {
        # Log to console
        Write-Warning "$Title - $Message"
    }
}

function Get-BatteryStatus {
    try {
        $status = [PowerStatus]::GetStatus()

        # ACLineStatus: 0 = Offline (on battery), 1 = Online (on AC), 255 = Unknown
        $onAC = switch ($status.ACLineStatus) {
            0       { $false }  # On battery
            1       { $true }   # On AC power
            default {
                Write-Warning "AC line status unknown - assuming on AC"
                $true
            }
        }

        # BatteryLifePercent: 0-100, or 255 if unknown/no battery
        $batteryCharge = $status.BatteryLifePercent
        if ($batteryCharge -ge 101) {
            Write-Warning "Battery percentage unknown (no battery?) - assuming 0%"
            $batteryCharge = 0
        }

        $script:lastBatteryStatus = @{
            BatteryCharge = $batteryCharge
            OnAC = $onAC
        }
    }
    catch {
        Write-Warning "Failed to get power status: $($_.Exception.Message)"
        # Return defaults
        $script:lastBatteryStatus = @{
            BatteryCharge = 0
            OnAC = $true
        }
    }

    Update-TrayIcon

    return $script:lastBatteryStatus
}

function Get-SmartplugPower {
    $url = "http://$SmartplugIP/cm?cmnd=Power"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec $SmartplugTimeoutSeconds
        return $response.POWER
    }
    catch {
        return ""
    }
}

function Set-SmartplugPower {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet("On", "Off")]
        [string]$State
    )

    Initialize-Hotspot

    $url = "http://$SmartplugIP/cm?cmnd=Power%20$State"

    for ($attempt = 1; $attempt -le $SmartplugCommunicationMaxAttempts; $attempt++) {
        try {
            $response = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec $SmartplugTimeoutSeconds
            $actualState = $response.POWER
            if ($actualState -ieq $State) {  # Case-insensitive
                Write-Host "Smart plug turned $State successfully." -ForegroundColor Green
                return $true
            } else {
                Write-Host "Smart plug responded but state is '$actualState' instead of '$State'" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "Attempt $attempt/$SmartplugCommunicationMaxAttempts failed: $($_.Exception.Message)" -ForegroundColor Yellow
            if ($attempt -lt $SmartplugCommunicationMaxAttempts) {
                Start-Sleep -Milliseconds 500
            }
        }
    }

    Write-Host "Failed to control smart plug after $SmartplugCommunicationMaxAttempts attempts" -ForegroundColor Red
    return $false
}

function Get-ModeDisplayName {
    param (
        [string]$Mode
    )

    switch ($Mode) {
        "Auto"         { "Auto" }
        "AutoHigh"     { "Auto High" }
        "ChargerOn"    { "Charger On" }
        "ChargerOff"   { "Charger Off" }
        "ToTargetOnce" { "To $($script:targetSoC)% Once" }
        default        { $Mode }
    }
}

function Draw-BatteryIcon {
    param (
        [int]$BatteryPercent,
        [bool]$OnAC,
        [string]$Mode
    )

    $bitmap = $null
    $graphics = $null

    try {
        # Create a 16x16 bitmap (standard tray icon size)
        $bitmap = New-Object System.Drawing.Bitmap(16, 16)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

        # Mode-specific border colors
        $borderColor = switch ($Mode) {
            "Auto"         { [System.Drawing.Color]::DodgerBlue }
            "AutoHigh"     { [System.Drawing.Color]::Magenta }
            "ChargerOn"    { [System.Drawing.Color]::LimeGreen }
            "ChargerOff"   { [System.Drawing.Color]::OrangeRed }
            "ToTargetOnce" { [System.Drawing.Color]::Gold }
            default        { [System.Drawing.Color]::Gray }
        }

        # Fill with black background
        $graphics.Clear([System.Drawing.Color]::Black)

        # Draw border using filled rectangles
        $borderBrush = New-Object System.Drawing.SolidBrush($borderColor)
        try {
            $graphics.FillRectangle($borderBrush, 0, 0, 16, 1)   # Top
            $graphics.FillRectangle($borderBrush, 0, 15, 16, 1)  # Bottom
            $graphics.FillRectangle($borderBrush, 0, 0, 1, 16)   # Left
            $graphics.FillRectangle($borderBrush, 15, 0, 1, 16)  # Right
        }
        finally {
            $borderBrush.Dispose()
        }

        # Text color: green if on AC, white if on battery
        $textColor = if ($OnAC) {
            [System.Drawing.Color]::Lime
        } else {
            [System.Drawing.Color]::White
        }

        $textBrush = New-Object System.Drawing.SolidBrush($textColor)
        # Use smaller font for 3-digit numbers (100)
        $fontSize = if ($BatteryPercent -gt 99) { 4 } else { 6 }
        $font = New-Object System.Drawing.Font("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold)
        $format = New-Object System.Drawing.StringFormat

        try {
            $format.Alignment = [System.Drawing.StringAlignment]::Center
            $format.LineAlignment = [System.Drawing.StringAlignment]::Center

            $text = $BatteryPercent.ToString()
            # Rectangle for text
            $rect = New-Object System.Drawing.RectangleF(0, 0, 16, 16)
            $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::SingleBitPerPixelGridFit
            $graphics.DrawString($text, $font, $textBrush, $rect, $format)
        }
        finally {
            $format.Dispose()
            $font.Dispose()
            $textBrush.Dispose()
        }

        # Convert bitmap to icon
        $hIcon = $bitmap.GetHicon()
        $tempIcon = $null
        try {
            $tempIcon = [System.Drawing.Icon]::FromHandle($hIcon)
            $icon = [System.Drawing.Icon]$tempIcon.Clone()
        }
        finally {
            [void][IconHelper]::DestroyIcon($hIcon)
            if ($tempIcon) { $tempIcon.Dispose() }
        }

        return $icon
    }
    finally {
        if ($graphics) { $graphics.Dispose() }
        if ($bitmap) { $bitmap.Dispose() }
    }
}

function Update-TrayIcon {
    # Skip if tray icon not initialized yet
    if (-not $script:trayIcon) { return }
    if (-not $script:lastBatteryStatus) { return }

    # Check if anything changed since last render
    $currentCharge = $script:lastBatteryStatus.BatteryCharge
    $currentOnAC = $script:lastBatteryStatus.OnAC
    $currentMode = $script:currentMode

    if ($script:lastRenderedBatteryStatus -and
        $script:lastRenderedBatteryStatus.BatteryCharge -eq $currentCharge -and
        $script:lastRenderedBatteryStatus.OnAC -eq $currentOnAC -and
        $script:lastRenderedMode -eq $currentMode -and
        $script:lastRenderedTargetSoC -eq $script:targetSoC) {
        return  # Nothing changed, skip update
    }

    $newIcon = $null
    try {
        $newIcon = Draw-BatteryIcon -BatteryPercent $currentCharge -OnAC $currentOnAC -Mode $currentMode

        # Assign new icon and dispose old one
        $oldIcon = $script:trayIcon.Icon
        $script:trayIcon.Icon = $newIcon
        if ($oldIcon) { $oldIcon.Dispose() }
    }
    catch {
        if ($newIcon) { $newIcon.Dispose() }
        Write-Warning "Failed to update tray icon: $($_.Exception.Message)"
    }

    $powerText = if ($currentOnAC) { "On AC" } else { "On Battery" }
    $modeDisplay = Get-ModeDisplayName -Mode $currentMode

    $script:trayIcon.Text = "$currentCharge% - $powerText`nMode: $modeDisplay"

    # Store the rendered state
    $script:lastRenderedBatteryStatus = @{
        BatteryCharge = $currentCharge
        OnAC = $currentOnAC
    }
    $script:lastRenderedMode = $currentMode
    $script:lastRenderedTargetSoC = $script:targetSoC
}

function Confirm-ACState {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [bool]$ExpectedOnAC
    )

    $pollIntervalMs = 100
    $maxIterations = [math]::Ceiling(($ConfirmationTimeoutSeconds * 1000) / $pollIntervalMs)

    for ($i = 0; $i -lt $maxIterations; $i++) {
        Start-Sleep -Milliseconds $pollIntervalMs
        $status = Get-BatteryStatus
        if ($status.OnAC -eq $ExpectedOnAC) {
            return $true
        }
    }

    return $false
}

function Confirm-SmartplugAvailable {
    for ($i = 0; $i -lt $SmartplugConnectionMaxAttempts; $i++) {
        $power = Get-SmartplugPower
        if ($power -ne "") {
            return $true
        }
        if ($i -lt ($SmartplugConnectionMaxAttempts - 1)) {
            Start-Sleep -Milliseconds 500  # Small backoff
        }
    }
    return $false
}

function Invoke-AutoCharging {
    param (
        [int]$BatteryCharge,
        [bool]$OnAC,
        [int]$MinLevel,
        [int]$MaxLevel
    )

    # Battery LOW and NOT on AC → Need to START charging
    if ($BatteryCharge -le $MinLevel -and -not $OnAC) {
        Write-Host "Battery low ($BatteryCharge%)! Attempting to turn ON smart plug" -ForegroundColor Yellow
        $plugSuccess = Set-SmartplugPower -State "On"
        if (-not $plugSuccess) {
            SendWindowsNotification -Title "Battery Low: $BatteryCharge%" -Message "Smart plug unreachable. Please plug in the charger manually!"
        }
        else {
            $confirmed = Confirm-ACState -ExpectedOnAC $true
            if (-not $confirmed) {
                SendWindowsNotification -Title "AC Power Not Detected" -Message "Smart plug responded but AC not detected. Please check the charger connection!"
            }
        }
    }
    # Battery HIGH and ON AC → Need to STOP charging
    elseif ($BatteryCharge -ge $MaxLevel -and $OnAC) {
        Write-Host "Battery sufficiently charged ($BatteryCharge%)! Attempting to turn OFF smart plug" -ForegroundColor Yellow
        $plugSuccess = Set-SmartplugPower -State "Off"
        if (-not $plugSuccess) {
            SendWindowsNotification -Title "Battery High: $BatteryCharge%" -Message "Smart plug unreachable. Please unplug the charger manually!"
        }
        else {
            $confirmed = Confirm-ACState -ExpectedOnAC $false
            if (-not $confirmed) {
                SendWindowsNotification -Title "Still On AC Power" -Message "Smart plug responded but still on AC. Please check the charger connection!"
            }
        }
    }
    else {
        Write-Host "No action needed" -ForegroundColor Gray
    }
}

function Invoke-ManualChargerControl {
    param (
        [ValidateSet("On", "Off")]
        [string]$DesiredState,
        [bool]$OnAC
    )

    # Check if action is needed
    $needsAction = ($DesiredState -eq "On" -and -not $OnAC) -or ($DesiredState -eq "Off" -and $OnAC)

    if (-not $needsAction) {
        Write-Host "No action needed" -ForegroundColor Gray
        return
    }

    Write-Host "Manual mode: Enforcing charger $DesiredState" -ForegroundColor Yellow
    $plugSuccess = Set-SmartplugPower -State $DesiredState

    if (-not $plugSuccess) {
        if ($DesiredState -eq "On") {
            SendWindowsNotification -Title "Manual Charging Mode" -Message "Smart plug unreachable. Please plug in the charger manually!"
        }
        else {
            SendWindowsNotification -Title "Manual Discharge Mode" -Message "Smart plug unreachable. Please unplug the charger manually!"
        }
    }
    else {
        $expectedOnAC = ($DesiredState -eq "On")
        $confirmed = Confirm-ACState -ExpectedOnAC $expectedOnAC
        if (-not $confirmed) {
            if ($DesiredState -eq "On") {
                SendWindowsNotification -Title "AC Not Detected" -Message "Smart plug responded but AC not detected. Please check the charger connection!"
            }
            else {
                SendWindowsNotification -Title "Still On AC" -Message "Smart plug responded but still on AC. Please check the charger connection!"
            }
        }
    }
}

function Invoke-ToTargetOnce {
    param (
        [int]$BatteryCharge,
        [bool]$OnAC,
        [int]$TargetSoC
    )

    if ($BatteryCharge -lt $TargetSoC) {
        # Below target - need to charge
        Invoke-ManualChargerControl -DesiredState "On" -OnAC $OnAC
    }
    elseif ($BatteryCharge -gt $TargetSoC) {
        # Above target - need to discharge
        Invoke-ManualChargerControl -DesiredState "Off" -OnAC $OnAC
    }
    else {
        # Exactly at target - done!
        Write-Host "Target $TargetSoC% reached! Restoring previous mode: $($script:previousMode)" -ForegroundColor Green
        SendWindowsNotification -Title "Target Reached" -Message "Battery at $TargetSoC%. Returning to $(Get-ModeDisplayName -Mode $script:previousMode) mode."
        Set-Mode -NewMode $script:previousMode
    }
}

function Invoke-BatteryCheck {
    $batteryStatus = Get-BatteryStatus
    $batteryCharge = $batteryStatus.BatteryCharge
    $onAC = $batteryStatus.OnAC

    $powerSource = if ($onAC) { "AC" } else { "Battery" }
    $modeDisplay = Get-ModeDisplayName -Mode $script:currentMode
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] - Battery: $batteryCharge% | Power: $powerSource | Mode: $modeDisplay"

    switch ($script:currentMode) {
        "Auto" {
            Invoke-AutoCharging -BatteryCharge $batteryCharge -OnAC $onAC `
                -MinLevel $MinBatteryLevel -MaxLevel $MaxBatteryLevel
        }
        "AutoHigh" {
            Invoke-AutoCharging -BatteryCharge $batteryCharge -OnAC $onAC `
                -MinLevel $MinBatteryLevelHigh -MaxLevel $MaxBatteryLevelHigh
        }
        "ChargerOn" {
            Invoke-ManualChargerControl -DesiredState "On" -OnAC $onAC
        }
        "ChargerOff" {
            Invoke-ManualChargerControl -DesiredState "Off" -OnAC $onAC
        }
        "ToTargetOnce" {
            Invoke-ToTargetOnce -BatteryCharge $batteryCharge -OnAC $onAC -TargetSoC $script:targetSoC
        }
    }
}

function Set-Mode {
    param (
        [ValidateSet("Auto", "AutoHigh", "ChargerOn", "ChargerOff", "ToTargetOnce")]
        [string]$NewMode,
        [int]$TargetSoC = 50
    )

    # Store previous mode if switching to ToTargetOnce (but not if already in ToTargetOnce)
    if ($NewMode -eq "ToTargetOnce") {
        if ($script:currentMode -ne "ToTargetOnce") {
            $script:previousMode = $script:currentMode
        }
        $script:targetSoC = $TargetSoC
        # Update menu text to show actual target
        $script:menuToTargetOnce.Text = "To $TargetSoC% Once"
    }
    else {
        # Reset menu text when leaving ToTargetOnce mode
        $script:menuToTargetOnce.Text = "To X% Once..."
    }

    $script:currentMode = $NewMode

    # Update checkbox states (only one can be checked at a time)
    $script:menuAuto.Checked = ($NewMode -eq "Auto")
    $script:menuAutoHigh.Checked = ($NewMode -eq "AutoHigh")
    $script:menuChargerOn.Checked = ($NewMode -eq "ChargerOn")
    $script:menuChargerOff.Checked = ($NewMode -eq "ChargerOff")
    $script:menuToTargetOnce.Checked = ($NewMode -eq "ToTargetOnce")

    $modeDisplay = Get-ModeDisplayName -Mode $NewMode
    Write-Host "Mode changed to: $modeDisplay" -ForegroundColor Green

    # Restart timer to apply changes immediately
    Restart-Timer
}

function Show-TargetSoCDialog {
    $form = New-Object System.Windows.Forms.Form
    $selectedValue = $null
    try {
        $form.Text = "Battery Range"
        $form.Size = New-Object System.Drawing.Size(280, 150)
        $form.StartPosition = "CenterScreen"
        $form.FormBorderStyle = "FixedDialog"
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false
        $form.TopMost = $true

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(15, 15)
        $label.Size = New-Object System.Drawing.Size(240, 20)
        $label.Text = "Enter target battery percentage:"
        $form.Controls.Add($label)

        $numericUpDown = New-Object System.Windows.Forms.NumericUpDown
        $numericUpDown.Location = New-Object System.Drawing.Point(15, 40)
        $numericUpDown.Size = New-Object System.Drawing.Size(80, 25)
        $numericUpDown.Minimum = 20
        $numericUpDown.Maximum = 100
        # Default to current battery level or 50
        $defaultValue = if ($script:lastBatteryStatus) { $script:lastBatteryStatus.BatteryCharge } else { 50 }
        $numericUpDown.Value = [Math]::Max(20, [Math]::Min(100, $defaultValue))
        $form.Controls.Add($numericUpDown)

        $percentLabel = New-Object System.Windows.Forms.Label
        $percentLabel.Location = New-Object System.Drawing.Point(100, 42)
        $percentLabel.Size = New-Object System.Drawing.Size(30, 20)
        $percentLabel.Text = "%"
        $form.Controls.Add($percentLabel)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(70, 75)
        $okButton.Size = New-Object System.Drawing.Size(75, 25)
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $okButton
        $form.Controls.Add($okButton)

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(155, 75)
        $cancelButton.Size = New-Object System.Drawing.Size(75, 25)
        $cancelButton.Text = "Cancel"
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $cancelButton
        $form.Controls.Add($cancelButton)

        # Focus the numeric input
        $form.Add_Shown({ $numericUpDown.Select() })

        # Get the result
        $result = $form.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedValue = [int]$numericUpDown.Value
        }
    }
    finally {
        $form.Dispose()
    }
    return $selectedValue
}

function Initialize-TrayIcon {
    # Create NotifyIcon
    $script:trayIcon = New-Object System.Windows.Forms.NotifyIcon

    # Create context menu
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

    # Auto menu item
    $script:menuAuto = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:menuAuto.Text = "Auto ($MinBatteryLevel%-$MaxBatteryLevel%)"
    $script:menuAuto.Checked = $true
    $script:menuAuto.Add_Click({
        Set-Mode "Auto"
    })

    # Auto High menu item
    $script:menuAutoHigh = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:menuAutoHigh.Text = "Auto High ($MinBatteryLevelHigh%-$MaxBatteryLevelHigh%)"
    $script:menuAutoHigh.Checked = $false
    $script:menuAutoHigh.Add_Click({
        Set-Mode "AutoHigh"
    })

    # Separator 1
    $separator1 = New-Object System.Windows.Forms.ToolStripSeparator

    # To X% Once menu item
    $script:menuToTargetOnce = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:menuToTargetOnce.Text = "To X% Once..."
    $script:menuToTargetOnce.Checked = $false
    $script:menuToTargetOnce.Add_Click({
        $targetSoC = Show-TargetSoCDialog
        if ($null -ne $targetSoC) {
            Set-Mode -NewMode "ToTargetOnce" -TargetSoC $targetSoC
        }
    })

    # Separator 2
    $separator2 = New-Object System.Windows.Forms.ToolStripSeparator

    # Charger On menu item
    $script:menuChargerOn = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:menuChargerOn.Text = "Charger On"
    $script:menuChargerOn.Checked = $false
    $script:menuChargerOn.Add_Click({
        Set-Mode "ChargerOn"
    })

    # Charger Off menu item
    $script:menuChargerOff = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:menuChargerOff.Text = "Charger Off"
    $script:menuChargerOff.Checked = $false
    $script:menuChargerOff.Add_Click({
        Set-Mode "ChargerOff"
    })

    # Separator 3
    $separator3 = New-Object System.Windows.Forms.ToolStripSeparator

    # Exit menu item
    $exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitMenuItem.Text = "Exit"

    # Add items to context menu
    $contextMenu.Items.Add($script:menuAuto) | Out-Null
    $contextMenu.Items.Add($script:menuAutoHigh) | Out-Null
    $contextMenu.Items.Add($separator1) | Out-Null
    $contextMenu.Items.Add($script:menuToTargetOnce) | Out-Null
    $contextMenu.Items.Add($separator2) | Out-Null
    $contextMenu.Items.Add($script:menuChargerOn) | Out-Null
    $contextMenu.Items.Add($script:menuChargerOff) | Out-Null
    $contextMenu.Items.Add($separator3) | Out-Null
    $contextMenu.Items.Add($exitMenuItem) | Out-Null

    $script:trayIcon.ContextMenuStrip = $contextMenu

    # Event: Exit clicked
    $exitMenuItem.Add_Click({
        Write-Host "Exiting - Enabling charging before shutdown" -ForegroundColor Yellow
        Set-SmartplugPower -State "On"

        [System.Windows.Forms.Application]::Exit()
    })

    # Set initial dynamic icon based on current battery status
    Update-TrayIcon

    # Show the tray icon
    $script:trayIcon.Visible = $true
}

function Initialize-Timer {
    $script:timer = New-Object System.Windows.Forms.Timer
    $script:timer.Interval = $CheckIntervalSeconds * 1000

    $script:timer.Add_Tick({
        Invoke-BatteryCheck
    })
}

function Restart-Timer {
    if (-not $script:timer) { return }

    # Stop the timer
    $script:timer.Stop()

    # Run initial battery check
    Invoke-BatteryCheck

    # Start the timer
    $script:timer.Start()
}

function Invoke-OnUIThread {
    param(
        [Parameter(Mandatory)]
        [ScriptBlock]$ScriptBlock
    )

    $control = $script:trayIcon.ContextMenuStrip

    if (-not $control) {
        Write-Warning "UI control not available for thread marshaling"
        return
    }

    try {
        if ($control.InvokeRequired) {
            # We're on a background thread - marshal to UI thread
            $control.Invoke([Action]$ScriptBlock)
        }
        else {
            # Already on UI thread
            & $ScriptBlock
        }
    }
    catch [System.ObjectDisposedException] {
        # Control was disposed (app shutting down)
        Write-Warning "Control disposed during invoke"
    }
    catch [System.InvalidOperationException] {
        # Handle was not created or other invoke issues
        Write-Warning "Could not invoke on UI thread: $($_.Exception.Message)"
    }
}

function Initialize-PowerEventHandler {
    # Create event handler for power mode changes (suspend/resume)
    $script:powerModeHandler = [Microsoft.Win32.PowerModeChangedEventHandler]{
        param($sender, $e)

        Invoke-OnUIThread -ScriptBlock {
            $timestamp = Get-Date -Format "HH:mm:ss"
            $mode = $e.Mode

            switch ($mode) {
                'Suspend' {
                    Write-Host "[$timestamp] System entering standby mode" -ForegroundColor Yellow

                    # Stop timer
                    if ($script:timer) {
                        $script:timer.Stop()
                    }

                    # Get current SoC from last known status
                    $currentSoC = if ($script:lastBatteryStatus) {
                        $script:lastBatteryStatus.BatteryCharge
                    } else { 50 }

                    if ($currentSoC -ge $StandbyOffSoCThreshold) {
                        Write-Host "[$timestamp] Battery at $currentSoC% (>= $StandbyOffSoCThreshold%) - Ensuring charger is off" -ForegroundColor Yellow

                        # Turn off smart plug
                        Set-SmartplugPower -State "Off" | out-null
                    }
                    elseif ($currentSoC -le $StandbyOnSoCThreshold) {
                        Write-Host "[$timestamp] Battery at $currentSoC% (<= $StandbyOnSoCThreshold%) - Ensuring charger is on" -ForegroundColor Yellow

                        # Turn on smart plug
                        Set-SmartplugPower -State "On" | out-null
                    }
                    else {
                        Write-Host "[$timestamp] Battery at $currentSoC% - Keeping charger state" -ForegroundColor Gray
                    }
                }
                'Resume' {
                    Write-Host "[$timestamp] System waking from standby" -ForegroundColor Cyan

                    # Update the tray icon
                    Get-BatteryStatus

                    # If requested, enable hotspot and check smartplug availability
                    Initialize-Hotspot

                    # Restart timer (runs immediate battery check)
                    Restart-Timer
                }
            }
        }
    }

    # Register the event handler
    [Microsoft.Win32.SystemEvents]::add_PowerModeChanged($script:powerModeHandler)
    Write-Host "Power event handler registered" -ForegroundColor Gray
}

# Main execution
$script:timer = $null
$script:trayIcon = $null
$script:powerModeHandler = $null
try {
    Write-Host "--- Battery Range ---" -ForegroundColor Cyan
    Write-Host "Right-click the tray icon for options" -ForegroundColor Cyan
    Write-Host "Modes: Auto ($MinBatteryLevel%-$MaxBatteryLevel%), Auto High ($MinBatteryLevelHigh%-$MaxBatteryLevelHigh%), Manual On/Off, To X% Once" -ForegroundColor Cyan
    if ($ManageWiFiHotspot) {
        Write-Host "Wi-Fi Hotspot management: Enabled" -ForegroundColor Cyan
    }
    Write-Host "Standby SoC thresholds: $StandbyOnSoCThreshold%, $StandbyOffSoCThreshold%" -ForegroundColor Cyan

    # Global state - Modes: "Auto", "AutoHigh", "ChargerOn", "ChargerOff", "ToTargetOnce"
    $script:currentMode = "Auto"
    $script:previousMode = "Auto"
    $script:targetSoC = 50

    # Enable hotspot before starting
    Initialize-Hotspot

    # Initialize and start timer
    Initialize-Timer
    Restart-Timer

    # Initialize tray icon
    Initialize-TrayIcon

    # Register power event handler for suspend/resume
    Initialize-PowerEventHandler

    # Run the Windows Forms application message loop
    [System.Windows.Forms.Application]::Run()
}
finally {
    # Cleanup on exit
    if ($script:powerModeHandler) {
        [Microsoft.Win32.SystemEvents]::remove_PowerModeChanged($script:powerModeHandler)
        Write-Host "Power event handler unregistered" -ForegroundColor Gray
    }
    if ($script:timer) {
        $script:timer.Stop()
        $script:timer.Dispose()
    }
    if ($script:trayIcon) {
        $script:trayIcon.Visible = $false
        if ($script:trayIcon.Icon) { $script:trayIcon.Icon.Dispose() }
        if ($script:trayIcon.ContextMenuStrip) { $script:trayIcon.ContextMenuStrip.Dispose() }
        $script:trayIcon.Dispose()
    }
    if ($script:mutex) {
        if ($script:hasHandle) {
            $script:mutex.ReleaseMutex()
        }
        $script:mutex.Dispose()
    }

    Write-Host "Battery Range stopped" -ForegroundColor Yellow
}
