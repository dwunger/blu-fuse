# Bluetooth Device Auto-Disconnect Script

## Overview
This Python script automatically disconnects a specified Bluetooth device after a period of inactivity, determined by the absence of audio playback. It's designed to run in environments where it's beneficial to automatically manage Bluetooth connections, such as freeing up a device for other connections or saving the device's battery life.

## Requirements
- Python 3.6 or higher
- Windows 10 or higher
- Bluetooth Command Line Tools installed (for device disconnection)
- `winrt` Python package (for monitoring media playback)

## Installation
1. **Install Bluetooth Command Line Tools:**
   - Download and install from [BluetoothInstaller.com](https://bluetoothinstaller.com/bluetooth-command-line-tools/).
2. **Python Dependencies:**
   - Install required Python packages by running `pip install winrt`.

## Usage
1. **List Bluetooth Devices:**
   - To list all paired Bluetooth devices, run the script with the `--list-devices` flag.
     ```
     python fuse.py --list-devices
     ```
     ```output
     . . .
     15 : EDIFIER W820NB Plus
     16 : EDIFIER W820NB Plus Avrcp Transport
     17 : EDIFIER W820NB Plus Avrcp Transport
     ```
     Use the common or base name of a device:
     ```
     EDIFIER W820NB Plus
     ```

2. **Monitor and Disconnect Device:**
   - Specify the device name, inactivity threshold, scan interval, and path to `btcom` executable.
     ```
     python fuse.py -n "COMMON DEVICE NAME" --threshold THRESHOLD_IN_MINUTES --scan-interval SCAN_INTERVAL_IN_MINUTES --btcom-path "PATH TO BTCOM"
     ```
   - **Parameters:**
     - `-n`, `--name`: Name of the Bluetooth device to monitor.
     - `--threshold`: Inactivity threshold in minutes before disconnecting.
     - `--scan-interval`: Interval in minutes to check for audio activity.
     - `--btcom-path`: Path to the `btcom` executable.

3. **Default Configuration (Optional):**
   - Uncomment and adjust the default configuration block inside the script for a predetermined setup.

## Notes
- Ensure the Bluetooth device name exactly matches the name listed under Bluetooth devices in Windows.
- The `btcom` path typically looks like `C:\Program Files (x86)\Bluetooth Command Line Tools\bin\btcom`.

## Troubleshooting
- If the device fails to disconnect, verify the correct device name and `btcom` path.
- Ensure Bluetooth Command Line Tools are correctly installed and accessible.

## License
This script is provided "as is", without warranty of any kind.
