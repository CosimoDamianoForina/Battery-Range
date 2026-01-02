<p align="center">
<img width="640" height="320" alt="Battery-Range" src="https://github.com/user-attachments/assets/998a9075-3fa1-4dab-bc0f-19a3f1e8af5b" />
</p>

# Battery-Range

A smart battery charging manager for Windows laptops that helps extend battery lifespan by maintaining charge within optimal State of Charge (SoC) ranges using a Tasmota-compatible smart plug.

## Why?

Lithium-ion batteries degrade faster when kept at 100% charge or frequently deep-discharged. Research suggests maintaining batteries between 20-80% (or even narrower ranges like 40-60%) can significantly extend their lifespan. Battery Range automates this by controlling your laptop charger through a smart plug.

## Features

- **Automatic Battery Management** - Maintains battery within configurable SoC ranges
- **Dual Range Profiles** - Normal (35-45%) and High (70-80%) modes for different needs
- **One-Time Target** - Charge/discharge to a specific target once
- **Manual Override** - Force charger on/off when needed
- **System Tray Integration** - Unobtrusive operation with right-click context menu
- **Failure Notifications** - Alerts you to manually plug/unplug if smart plug is unreachable
- **Single Instance** - Prevents multiple copies from running simultaneously

## Requirements

- Windows 10/11
- PowerShell 5.1 or later
- Tasmota-compatible Wi-Fi smart plug (see [Devices with Factory Flashed Tasmota](https://templates.blakadder.com/preflashed.html))
- Network connectivity between laptop and smart plug (regular Wi-Fi network or Windows Mobile Hotspot)

## Quick Start

1. Set up your Tasmota smart plug (see [Smart Plug Setup](#tasmota-smart-plug-setup))
2. Download `Battery-Range.ps1`
3. Run the script:
   ```powershell
   .\Battery-Range.ps1
   ```
4. Right-click the system tray icon to select modes

## Parameters

| Parameter | Default | Range | Description |
|-----------|---------|-------|-------------|
| `-TasmotaTimeoutSeconds` | 2 | 1-5 | Timeout in seconds for the Tasmota smart plug |
| `-ConfirmationTimeoutSeconds` | 5 | 1-10 | Timeout for confirming charging state change |
| `-CheckIntervalSeconds` | 30 | 15-300 | How often to check battery status |
| `-MaxBatteryLevel` | 45 | 20-100 | Maximum battery % for Auto mode |
| `-MinBatteryLevel` | 35 | 10-99 | Minimum battery % for Auto mode |
| `-MaxBatteryLevelHigh` | 80 | 20-100 | Maximum battery % for Auto High mode |
| `-MinBatteryLevelHigh` | 70 | 10-99 | Minimum battery % for Auto High mode |

**Min values must be less than their corresponding Max values.**

### Example with Custom Settings

```powershell
.\Battery-Range.ps1 -TasmotaPlugIP "192.168.1.100" -MinBatteryLevel 40 -MaxBatteryLevel 60
```

## Modes

| Mode | Description |
|------|-------------|
| **Auto** | Maintains battery between `MinBatteryLevel` and `MaxBatteryLevel` (default: 35-45%) |
| **Auto High** | Maintains battery between `MinBatteryLevelHigh` and `MaxBatteryLevelHigh` (default: 70-80%) |
| **To X% Once...** | Brings the battery to a specific target and then returns to the previous mode |
| **Charger On** | Forces charger on continuously (useful before travel) |
| **Charger Off** | Forces charger off continuously |

## System Tray Icon

The tray icon provides at-a-glance status:

| Element | Meaning |
|---------|---------|
| **Number** | Current battery percentage |
| **Text color: Green** | Laptop is on AC power |
| **Text color: White** | Laptop is on battery |
| **Border: Blue** | Auto mode |
| **Border: Magenta** | Auto High mode |
| **Border: Golden** | To X% Once mode |
| **Border: Green** | Charger On mode |
| **Border: Orange** | Charger Off mode |

## Tasmota Smart Plug Setup

### Flashing Tasmota

If your smart plug doesn't already have Tasmota, you'll need to flash it. See the [Tasmota documentation](https://tasmota.github.io/docs/) for device-specific instructions.

### Configuring Static IP for Windows Mobile Hotspot

If you're using your laptop's Mobile Hotspot to connect the smart plug (useful for portable setups), configure a static IP on the plug:

1. Connect to your Tasmota plug's web interface
2. Go to **Console**
3. Enter the following command:

```
Backlog IPAddress1 192.168.137.101; IPAddress2 192.168.137.1; IPAddress3 255.255.255.0; IPAddress4 192.168.137.1; Restart 1
```

This sets:
- **IPAddress1**: Static IP for the plug (`192.168.137.101`)
- **IPAddress2**: Gateway (Windows hotspot is always `192.168.137.1`)
- **IPAddress3**: Subnet mask (`255.255.255.0`)
- **IPAddress4**: DNS server (using the gateway)

The plug will restart with the new network settings.

### Configuring for Regular Wi-Fi Network

For a regular home network, use your router's DHCP reservation feature to assign a static IP, or configure via Tasmota console:

```
Backlog IPAddress1 192.168.1.101; IPAddress2 192.168.1.1; IPAddress3 255.255.255.0; IPAddress4 192.168.1.1; Restart 1
```

Adjust the IP addresses to match your network configuration.

## Running at Startup

1. Press `Win + R`, type `shell:startup`, press Enter
2. Create a shortcut to run the script:
   - Right-click → New → Shortcut
   - Target: `powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File "[path to the script]\Battery-Range.ps1"`
   - Name: `Battery-Range`

## How It Works

```
┌─────────────────────────────────────────────────────────┐
│                    Battery-Range                        │
├─────────────────────────────────────────────────────────┤
│                                                         │
│   Check Battery Level (every N seconds)                 │
│              │                                          │
│              ▼                                          │
│   ┌─────────────────────┐                               │
│   │  Battery ≤ Min%     │──Yes──► Turn ON Smart Plug    │
│   │  AND discharging?   │                               │
│   └─────────────────────┘                               │
│              │ No                                       │
│              ▼                                          │
│   ┌─────────────────────┐                               │
│   │  Battery ≥ Max%     │──Yes──► Turn OFF Smart Plug   │
│   │  AND charging?      │                               │
│   └─────────────────────┘                               │
│              │ No                                       │
│              ▼                                          │
│        Do Nothing                                       │
│                                                         │
└─────────────────────────────────────────────────────────┘
```

## Robustness Notes

- **On Exit**: The script automatically turns the charger ON when exiting to prevent accidental battery drain
- **Power Outages**: If the smart plug loses power, it will restore its previous state when power returns (default Tasmota behavior)
- **Plug Failures**: If the smart plug is unreachable, a Windows notification prompts manual action
- **Wiring Issues**: If the battery charging status is inconsistent with the smart plug status, a notification prompts the user to check the charger connection 

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Script won't start | Ensure execution policy allows scripts: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` |
| Can't connect to plug | Verify IP address and that laptop and plug are on same network |
| Notifications not appearing | Check Windows notification settings for PowerShell |
| Battery not detected | Ensure running on a laptop with battery installed |

## License

MIT License - See [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

## Acknowledgments

- [Tasmota](https://github.com/arendst/tasmota) - Open source firmware for ESP8266/ESP32 devices


