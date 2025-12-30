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
    - Manual charger control override
    - System tray integration with context menu
    - Smart plug failure notifications with manual action prompts

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
    [Parameter(HelpMessage = "IP address of the Tasmota smart plug")]
    [string]$TasmotaPlugIP = "192.168.137.101",

    [Parameter(HelpMessage = "Timeout in seconds for the Tasmota smart plug")]
    [ValidateRange(1, 5)]
    [int]$TasmotaTimeoutSeconds = 2,

    [Parameter(HelpMessage = "Delay in milliseconds before confirming the status")]
    [ValidateRange(100, 5000)]
    [int]$ConfirmationDelayMilliseconds = 1000,

    [Parameter(HelpMessage = "Battery check interval in seconds")]
    [ValidateRange(10, 300)]
    [int]$CheckIntervalSeconds = 30,

    [Parameter(HelpMessage = "Maximum battery level for normal Auto mode")]
    [ValidateRange(20, 100)]
    [int]$MaxBatteryLevel = 45,

    [Parameter(HelpMessage = "Minimum battery level for normal Auto mode")]
    [ValidateRange(10, 99)]
    [int]$MinBatteryLevel = 35,

    [Parameter(HelpMessage = "Maximum battery level for Auto High mode")]
    [ValidateRange(20, 100)]
    [int]$MaxBatteryLevelHigh = 80,

    [Parameter(HelpMessage = "Minimum battery level for Auto High mode")]
    [ValidateRange(10, 99)]
    [int]$MinBatteryLevelHigh = 70
)

##############################################################################

# Validate that min < max for both ranges
if ($MinBatteryLevel -ge $MaxBatteryLevel) {
    throw "MinBatteryLevel ($MinBatteryLevel) must be less than MaxBatteryLevel ($MaxBatteryLevel)"
}
if ($MinBatteryLevelHigh -ge $MaxBatteryLevelHigh) {
    throw "MinBatteryLevelHigh ($MinBatteryLevelHigh) must be less than MaxBatteryLevelHigh ($MaxBatteryLevelHigh)"
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
        exit
    }
}
catch {
    Write-Host "Error: Could not create mutex." -ForegroundColor Red
    exit
}

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Add icon extraction capability from system DLLs
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;

public class IconExtractor {
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    public static Icon Extract(string filePath, int iconIndex) {
        IntPtr hIcon = ExtractIcon(IntPtr.Zero, filePath, iconIndex);
        if (hIcon == IntPtr.Zero) return null;

        try {
            Icon icon = Icon.FromHandle(hIcon);
            return (Icon)icon.Clone();
        }
        finally {
            DestroyIcon(hIcon);
        }
    }
}
"@ -ReferencedAssemblies System.Drawing

# Global state - Modes: "Auto", "AutoHigh", "ChargerOn", "ChargerOff"
$script:currentMode = "Auto"

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
    $battery = Get-CimInstance -ClassName Win32_Battery
    $estimatedCharge = $battery.estimatedChargeRemaining
    $batteryStatus = $battery.BatteryStatus

    # 1=Discharging, 2=Unknown/AC, 6=Charging, 7=ChargingHigh, 8=ChargingLow, 9=ChargingCritical
    # Note: Status 2 often means "Plugged in, not charging" (e.g. at 100% or threshold limit)
    $chargingStatuses = @(2, 6, 7, 8, 9)
    $isCharging = $batteryStatus -in $chargingStatuses

    return @{
        Charge     = $estimatedCharge
        IsCharging = $isCharging
    }
}

function Set-TasmotaPlug {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet("On", "Off")]
        [string]$State
    )

    $url = "http://$TasmotaPlugIP/cm?cmnd=Power%20$State"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec $TasmotaTimeoutSeconds
        Write-Host "Smart plug turned $State successfully. Response: $($response.POWER)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to control smart plug: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Invoke-AutoCharging {
    param (
        [int]$Charge,
        [bool]$IsCharging,
        [int]$MinLevel,
        [int]$MaxLevel
    )

    # Battery LOW and NOT charging → Need to START charging
    if ($Charge -lt $MinLevel -and -not $IsCharging) {
        Write-Host "Battery low! Attempting to turn ON smart plug..." -ForegroundColor Yellow
        $plugSuccess = Set-TasmotaPlug -State "On"
        if (-not $plugSuccess) {
            SendWindowsNotification -Title "Battery Low: $Charge%" -Message "Smart plug unreachable. Please plug in the charger manually!"
        }
        else {
            Write-Host "Smart plug activated - charging should start automatically." -ForegroundColor Cyan
            # Wait and verify charging started
            Start-Sleep -Milliseconds $ConfirmationDelayMilliseconds
            $newStatus = Get-BatteryStatus
            if (-not $newStatus.IsCharging) {
                SendWindowsNotification -Title "Charging Not Started" -Message "Smart plug responded but charging didn't start. Please check the charger connection!"
            }
        }
    }
    # Battery HIGH and CHARGING → Need to STOP charging
    elseif ($Charge -gt $MaxLevel -and $IsCharging) {
        Write-Host "Battery sufficiently charged! Attempting to turn OFF smart plug..." -ForegroundColor Yellow
        $plugSuccess = Set-TasmotaPlug -State "Off"
        if (-not $plugSuccess) {
            SendWindowsNotification -Title "Battery High: $Charge%" -Message "Smart plug unreachable. Please unplug the charger manually!"
        }
        else {
            Write-Host "Smart plug deactivated - charging should stop automatically." -ForegroundColor Cyan
            # Wait and verify charging stopped
            Start-Sleep -Milliseconds $ConfirmationDelayMilliseconds
            $newStatus = Get-BatteryStatus
            if ($newStatus.IsCharging) {
                SendWindowsNotification -Title "Charging Not Stopped" -Message "Smart plug responded but charging continues. Please check the charger connection!"
            }
        }
    }
    else {
        Write-Host "Range $MinLevel%-$MaxLevel%. No action needed." -ForegroundColor Gray
    }
}

function Invoke-ManualChargerControl {
    param (
        [ValidateSet("On", "Off")]
        [string]$DesiredState,
        [bool]$IsCharging
    )

    # Check if action is needed
    $needsAction = ($DesiredState -eq "On" -and -not $IsCharging) -or ($DesiredState -eq "Off" -and $IsCharging)

    if (-not $needsAction) {
        Write-Host "Manual mode: Charger already in desired state ($DesiredState)." -ForegroundColor Gray
        return
    }

    Write-Host "Manual mode: Enforcing charger $DesiredState..." -ForegroundColor Yellow
    $plugSuccess = Set-TasmotaPlug -State $DesiredState

    if (-not $plugSuccess) {
        if ($DesiredState -eq "On") {
            SendWindowsNotification -Title "Manual Charging Mode" -Message "Smart plug unreachable. Please plug in the charger manually!"
        }
        else {
            SendWindowsNotification -Title "Manual Discharge Mode" -Message "Smart plug unreachable. Please unplug the charger manually!"
        }
    }
    else {
        Write-Host "Smart plug $DesiredState command sent successfully." -ForegroundColor Cyan
        # Wait and verify state changed
        Start-Sleep -Milliseconds $ConfirmationDelayMilliseconds
        $newStatus = Get-BatteryStatus
        $stateCorrect = ($DesiredState -eq "On" -and $newStatus.IsCharging) -or ($DesiredState -eq "Off" -and -not $newStatus.IsCharging)
        if (-not $stateCorrect) {
            if ($DesiredState -eq "On") {
                SendWindowsNotification -Title "Charging Not Started" -Message "Smart plug responded but charging didn't start. Please check the charger connection!"
            }
            else {
                SendWindowsNotification -Title "Charging Not Stopped" -Message "Smart plug responded but charging continues. Please check the charger connection!"
            }
        }
    }
}

function Check-Battery {
    $batteryStatus = Get-BatteryStatus
    $estimatedCharge = $batteryStatus.Charge
    $isCharging = $batteryStatus.IsCharging

    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Battery: $estimatedCharge% | Charging: $isCharging | Mode: $($script:currentMode)"

    switch ($script:currentMode) {
        "Auto" {
            Invoke-AutoCharging -Charge $estimatedCharge -IsCharging $isCharging `
                -MinLevel $MinBatteryLevel -MaxLevel $MaxBatteryLevel
        }
        "AutoHigh" {
            Invoke-AutoCharging -Charge $estimatedCharge -IsCharging $isCharging `
                -MinLevel $MinBatteryLevelHigh -MaxLevel $MaxBatteryLevelHigh
        }
        "ChargerOn" {
            Invoke-ManualChargerControl -DesiredState "On" -IsCharging $isCharging
        }
        "ChargerOff" {
            Invoke-ManualChargerControl -DesiredState "Off" -IsCharging $isCharging
        }
    }
}

function Set-Mode {
    param (
        [ValidateSet("Auto", "AutoHigh", "ChargerOn", "ChargerOff")]
        [string]$NewMode
    )

    $script:currentMode = $NewMode

    # Update checkbox states (only one can be checked at a time)
    $script:menuAuto.Checked = ($NewMode -eq "Auto")
    $script:menuAutoHigh.Checked = ($NewMode -eq "AutoHigh")
    $script:menuChargerOn.Checked = ($NewMode -eq "ChargerOn")
    $script:menuChargerOff.Checked = ($NewMode -eq "ChargerOff")

    Write-Host "Mode changed to: $NewMode" -ForegroundColor Green

    # Restart timer to apply changes immediately
    Restart-Timer
}

function Initialize-TrayIcon {
    # Create NotifyIcon
    $script:trayIcon = New-Object System.Windows.Forms.NotifyIcon

    # Try to get a battery/power icon from system DLLs
    try {
        $icon = [IconExtractor]::Extract("$env:SystemRoot\System32\powercpl.dll", 2)
        if ($icon) {
            $script:trayIcon.Icon = $icon
        }
        else {
            # Last resort: use PowerShell icon
            $script:trayIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $PID).Path)
        }
    }
    catch {
        # Fallback to PowerShell icon
        $script:trayIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $PID).Path)
    }

    $script:trayIcon.Text = "Battery Range"

    # Create context menu
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

    # Auto menu item
    $script:menuAuto = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:menuAuto.Text = "Auto"
    $script:menuAuto.Checked = $true
    $script:menuAuto.Add_Click({
        Set-Mode "Auto"
    })

    # Auto High menu item
    $script:menuAutoHigh = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:menuAutoHigh.Text = "Auto high"
    $script:menuAutoHigh.Checked = $false
    $script:menuAutoHigh.Add_Click({
        Set-Mode "AutoHigh"
    })

    # Separator 1
    $separator1 = New-Object System.Windows.Forms.ToolStripSeparator

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

    # Separator 2
    $separator2 = New-Object System.Windows.Forms.ToolStripSeparator

    # Exit menu item
    $exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitMenuItem.Text = "Exit"

    # Add items to context menu
    $contextMenu.Items.Add($script:menuAuto) | Out-Null
    $contextMenu.Items.Add($script:menuAutoHigh) | Out-Null
    $contextMenu.Items.Add($separator1) | Out-Null
    $contextMenu.Items.Add($script:menuChargerOn) | Out-Null
    $contextMenu.Items.Add($script:menuChargerOff) | Out-Null
    $contextMenu.Items.Add($separator2) | Out-Null
    $contextMenu.Items.Add($exitMenuItem) | Out-Null

    $script:trayIcon.ContextMenuStrip = $contextMenu

    # Event: Exit clicked
    $exitMenuItem.Add_Click({
        Write-Host "Exiting - Enabling charging before shutdown..." -ForegroundColor Yellow
        Set-TasmotaPlug -State "On"

        [System.Windows.Forms.Application]::Exit()
    })

    # Show the tray icon
    $script:trayIcon.Visible = $true
}

function Initialize-Timer {
    $script:timer = New-Object System.Windows.Forms.Timer
    $script:timer.Interval = $CheckIntervalSeconds * 1000

    $script:timer.Add_Tick({
        Check-Battery
    })
}

function Restart-Timer {
    if (-not $script:timer) { return }

    # Stop the timer
    $script:timer.Stop()

    # Run initial battery check
    Check-Battery

    # Start the timer
    $script:timer.Start()
}

# Main execution
try {
    Write-Host "Starting Battery Range..." -ForegroundColor Cyan
    Write-Host "Right-click the tray icon for options." -ForegroundColor Cyan
    Write-Host "Modes: Auto ($MinBatteryLevel%-$MaxBatteryLevel%), Auto High ($MinBatteryLevelHigh%-$MaxBatteryLevelHigh%), Manual On/Off" -ForegroundColor Cyan

    # Initialize and start timer
    Initialize-Timer
    Restart-Timer

    # Initialize tray icon
    Initialize-TrayIcon

    # Run the Windows Forms application message loop
    [System.Windows.Forms.Application]::Run()
}
finally {
    # Cleanup on exit
    if ($script:timer) {
        $script:timer.Stop()
        $script:timer.Dispose()
    }
    if ($script:trayIcon) {
        $script:trayIcon.Visible = $false
        $script:trayIcon.Dispose()
    }

    # Release the mutex
    if ($script:hasHandle -and $script:mutex) {
        $script:mutex.ReleaseMutex()
    }
    if ($script:mutex) {
        $script:mutex.Dispose()
    }

    Write-Host "Battery Range stopped." -ForegroundColor Yellow
}
