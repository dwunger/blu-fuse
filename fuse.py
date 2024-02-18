import subprocess
import argparse
import time
import asyncio
import winrt.windows.media.control as wmc
import sys
import os

async def get_media_session():
    # Request the Global System Media Transport Controls Session Manager
    sessions = await wmc.GlobalSystemMediaTransportControlsSessionManager.request_async()
    # Get the current media session
    session = sessions.get_current_session()
    return session

def media_is_playing(session):
    # If no session is found, return False as media is not playing
    if session is None:
        return False
    
    # Get the current playback status of the session
    playback_status = session.get_playback_info().playback_status
    
    # Compare the playback status to the 'Playing' status
    return playback_status == wmc.GlobalSystemMediaTransportControlsSessionPlaybackStatus.PLAYING

async def is_device_playing_audio(device_name):
    # Function to check if the specified device is currently playing audio
    session = await get_media_session()
    return media_is_playing(session)

def list_bluetooth_devices():
    list_devices_cmd = 'Get-PnpDevice | Where-Object {$_.Class -eq "Bluetooth"} | Format-List *'
    try:
        devices_output = subprocess.check_output(['powershell', '-Command', list_devices_cmd], text=True)
        name_list = []
        print("Listing friendly names of all Bluetooth devices:")
        for line in devices_output.splitlines():
            if "FriendlyName" in line:
                name_list.append(" ".join(line.split()[2:]))
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while listing devices: {e}")
    name_list.sort()
    for i, device in enumerate(name_list):
        print(f"{i} : {device}") 


def disconnect_bluetooth_device(device_name, DEBUG = False):
    '''
    (Not working)
    Prototype to remove bttools dependency
    '''
    # PowerShell command to list all Bluetooth devices
    list_devices_cmd = 'Get-PnpDevice | Where-Object {$_.Class -eq "Bluetooth"} | Format-List *'
    # Note: Changed command to list all properties of Bluetooth devices for better debugging

    try:
        # Execute the command to list all Bluetooth devices and their properties
        devices_output = subprocess.check_output(['powershell', '-Command', list_devices_cmd], text=True)
        if DEBUG:
            print("Debug - Listing all Bluetooth devices and their properties:")
            print(devices_output)  # Debug print to see the output of the device listing

            # Debug print to check if the device name is in the output
            if device_name in devices_output:
                print(f"Debug - Device '{device_name}' found in the output.")
            else:
                print(f"Debug - Device '{device_name}' NOT found in the output. Check device name spelling and Bluetooth connection status.")

        # Assuming we need to find a more specific property to identify the device
        # Example: Look for the FriendlyName property in the output
        found_device = False
        for line in devices_output.splitlines():
            if "FriendlyName" in line and device_name in line:
                print(f"Debug - Device '{device_name}' found with FriendlyName in the output.")
                found_device = True
                # Extract the InstanceId and construct the command to remove the device (if needed)
                # Note: You'll need to parse the output properly to extract the InstanceId

        if not found_device and DEBUG:
            print(f"Debug - Device '{device_name}' not found with FriendlyName. Check if device is paired and connected.")

    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e.output}")


def disconnect_bluetooth_device(device_name, btcom_path, DEBUG=True):
    # Source:
    # https://bluetoothinstaller.com/bluetooth-command-line-tools/btcom.html
    
    # Command to disable the hands-free service (HFP) and the audio sink service (A2DP)
    disable_hfp_cmd = '"{}" -n "{}" -r -s111e'.format(btcom_path, device_name)
    disable_a2dp_cmd = '"{}" -n "{}" -r -s110b'.format(btcom_path, device_name)


    try:
        # Execute the command to disable the HFP service
        if DEBUG:
            print(f"Disabling HFP service for device '{device_name}'...")
        subprocess.check_output(disable_hfp_cmd, shell=True)
        
        # Execute the command to disable the A2DP service
        if DEBUG:
            print(f"Disabling A2DP service for device '{device_name}'...")
        subprocess.check_output(disable_a2dp_cmd, shell=True)

        if DEBUG:
            print(f"Device '{device_name}' disconnected successfully.")

    except subprocess.CalledProcessError as e:
        print(f"An error occurred while disconnecting '{device_name}': {e}")


async def main(device_name, threshold, scan_interval, btcom_path):
    no_audio_duration = 0  # Duration in minutes that the device has had no audio

    while True:
        if await is_device_playing_audio(device_name):
            no_audio_duration = 0  # Reset the counter if audio is playing
        else:
            no_audio_duration += scan_interval

        if no_audio_duration >= threshold:
            disconnect_bluetooth_device(device_name, btcom_path)
            # Should the script exit on device disconnect?
            # break

        await asyncio.sleep(scan_interval * 60)  # Wait for the next scan interval

def install_winrt():
    try:
        # Attempt to install the winrt package using pip
        subprocess.check_call([sys.executable, "-m", "pip", "install", "winrt"])
        print("winrt package installed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install winrt package: {e}")

def check_python_version_and_install_winrt():
    if sys.version_info.major == 3 and sys.version_info.minor == 7:
        print("Python 3.7 detected. Attempting to install winrt package...")
        install_winrt()
    else:
        print(f"Python {sys.version_info.major}.{sys.version_info.minor} detected. No action required for winrt package installation.")
        
def generate_startup_script(python_script_path):
    startup_folder = os.path.expandvars(r'%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup')
    script_name = os.path.basename(python_script_path)
    vbs_script_path = os.path.join(startup_folder, f"{script_name}.vbs")

    vbs_content = f"""Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "{sys.executable}" & Chr(34) & " " & Chr(34) & "{python_script_path}" & Chr(34), 0
Set WshShell = Nothing"""

    with open(vbs_script_path, 'w') as vbs_file:
        vbs_file.write(vbs_content)

    print(f"Silent startup script generated and placed in: {vbs_script_path}")
    

def generate_powershell_startup_script(python_script_path):
    startup_folder = os.path.expandvars(r'%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup')
    ps1_script_path = os.path.join(startup_folder, "StartMyPythonScript.ps1")

    # Pre-format the paths to include the necessary escape characters
    python_exe_escaped = sys.executable.replace('\\', '\\\\')
    script_path_escaped = python_script_path.replace('\\', '\\\\')

    # Use the pre-formatted paths in the PowerShell content
    ps1_content = f"""
    $pythonExe = "{python_exe_escaped}"
    $scriptPath = "{script_path_escaped}"
    Start-Process -FilePath $pythonExe -ArgumentList "$scriptPath" -WindowStyle Hidden
    """

    with open(ps1_script_path, 'w') as ps1_file:
        ps1_file.write(ps1_content)

    print(f"PowerShell startup script generated and placed in: {ps1_script_path}")
    
def generate_batch_startup_script(python_script_path):
    startup_folder = os.path.expandvars(r'%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup')
    batch_script_path = os.path.join(startup_folder, "StartMyPythonScript.bat")

    batch_content = f"""
@echo off
"{sys.executable}" "{python_script_path}"
    """.strip()

    with open(batch_script_path, 'w') as batch_file:
        batch_file.write(batch_content)

    print(f"Batch startup script generated and placed in: {batch_script_path}")

if __name__ == "__main__":
    # Check Python version and attempt to install winrt if necessary
    check_python_version_and_install_winrt()
    
    parser = argparse.ArgumentParser(description="Auto disconnect Bluetooth device on inactivity.")
    parser.add_argument("-n", "--name", type=str, required=False, default=None, help="Name of the Bluetooth device")
    parser.add_argument("--threshold", type=int, required=False, default=None, help="Inactivity threshold in minutes")
    parser.add_argument("--scan-interval", type=int, required=False, default=None, help="Scan interval in minutes")
    parser.add_argument("--btcom-path", type=str, required=False, default=None, help="Path to btcom executable\nSee link for install:\n\thttps://bluetoothinstaller.com/bluetooth-command-line-tools/btcom.html")
    parser.add_argument("--list-devices", action="store_true", help="List all paired Bluetooth devices and exit")
    parser.add_argument("--generate-startup-script", action="store_true", help="Generate silent startup script and place it in the Windows Startup folder")
    parser.add_argument("--generate-powershell-startup-script", action="store_true", help="Test failed: Generate PowerShell startup script and place it in the Windows Startup folder")
    parser.add_argument("--generate-batch-startup-script", action="store_true", help="Test failed: Generate batch startup script and place it in the Windows Startup folder")

    args = parser.parse_args()
    if args.list_devices:
        list_bluetooth_devices()
        print("\nDevice enumeration complete. Exiting process...")
        exit()

    # Check for the generate-startup-script flag
    if args.generate_startup_script:
        # Assuming your Python script's path is known. Replace 'your_script.py' with your actual script filename or path.
        python_script_path = os.path.abspath(__file__)  # Gets the path of the current script
        generate_startup_script(python_script_path)
        print("\nStartup script generated. Exiting process...")
        exit()
        
    # Alt ption for PowerShell startup script
    if args.generate_powershell_startup_script:
        python_script_path = os.path.abspath(__file__)  # Gets the path of the current script
        generate_powershell_startup_script(python_script_path)
        print("\nPowerShell startup script generated. Exiting process...")
        exit()
        
    if args.generate_batch_startup_script:
        python_script_path = os.path.abspath(__file__)  # Gets the path of the current script
        generate_batch_startup_script(python_script_path)
        print("\nBatch startup script generated. Exiting process...")
        exit()

    # # Uncomment this block to hard-code a handler for a single device as the default behavior:
    # device_name = "EDIFIER W820NB Plus"
    # max_inactivity = 60 # minutes
    # scan_interval =   5 # minutes
    # btcom_path = "C:/Program Files (x86)/Bluetooth Command Line Tools/bin/btcom"
    
    # if(args.name == None):
    #     args.name = device_name
    # if (args.threshold == None):
    #     args.threshold = max_inactivity
    # if (args.scan_interval == None):
    #     args.scan_interval = scan_interval
    # if (args.btcom_path == None):
    #     args.btcom_path = btcom_path
        
    # Check if all required arguments are provided or set
    if not args.name or not args.threshold or not args.scan_interval or not args.btcom_path:
        print("Missing required arguments.")
        parser.print_help()
        exit()

    asyncio.run(main(args.name, args.threshold, args.scan_interval, args.btcom_path))


