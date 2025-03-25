import os
import time
import serial
import serial.tools.list_ports
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
from PIL import Image
import win32gui
import win32api
import win32ui
import win32con
import ctypes
import ctypes.wintypes
import numpy as np

last_known_apps = set()

def find_esp32_port(baud_rate=460800, retry_delay=2.0):
    """
    Continuously search for an available port that can be opened at 'baud_rate'.
    If no port is found or open fails, sleep 'retry_delay' seconds and retry.
    Returns a valid serial.Serial object once successful.
    """
    while True:
        ports = list(serial.tools.list_ports.comports())
        if not ports:
            print("No COM ports found. Retrying in", retry_delay, "seconds...")
            time.sleep(retry_delay)
            continue

        for port_info in ports:
            device = port_info.device  # e.g. 'COM3', 'COM4', etc.
            description = port_info.description  # e.g. 'USB-SERIAL CH340'
            # Optionally, inspect 'port_info.vid'/'port_info.pid' if you need specific hardware matching.
            try:
                # Try to open the port
                print(f"Attempting to open {device} ({description}) at {baud_rate} baud...")
                ser = serial.Serial(device, baud_rate, timeout=0.1)
                # If we make it this far without exception, we have a port open.
                print(f"Connected to {device} successfully.")
                return ser
            except Exception as e:
                # Could not open this port; move on to the next.
                print(f"Failed to open {device}: {e}")

        print(f"Could not find a valid port. Retrying in {retry_delay} seconds...")
        time.sleep(retry_delay)

def set_app_volume(app_name: str, volume: float):
    """
    Set the volume for a given app (by process name) to 'volume' [0.0..1.0].
    Uses PyCAW to find the matching session.
    """
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        process = session.Process
        if process:
            process_name = process.name()
            if (process_name.lower() == app_name.lower() or
               process_name.lower().removesuffix(".exe") == app_name.lower()):
                volume_control = session.SimpleAudioVolume
                volume_control.SetMasterVolume(volume, None)
                return
    print(f"Application '{app_name}' not found or not playing audio.")

def parse_line(ser):
    """
    Checks the serial buffer for data. If there's a line in the form 'app_name,volume',
    set that volume via set_app_volume, then send back the updated volume to the ESP32.
    Otherwise, just print the line.
    """
    if ser.in_waiting > 0:
        line = ser.readline().decode('utf-8', errors='ignore').strip()
        if not line:
            return

        # Attempt to parse a volume command
        if ',' in line:
            print(line)
            parts = line.split(',')
            if len(parts) == 2:
                app_name = parts[0].strip()
                try:
                    vol_percent = int(parts[1].strip())
                    volume = vol_percent / 100.0
                    if 0.0 <= volume <= 1.0:
                        # Set the volume
                        set_app_volume(app_name, volume)

                        # Read back the actual volume to confirm
                        sessions = AudioUtilities.GetAllSessions()
                        for session in sessions:
                            process = session.Process
                            if process:
                                process_name = process.name()
                                if (process_name.lower() == app_name.lower() or
                                    process_name.lower().removesuffix(".exe") == app_name.lower()):
                                    current_vol = session.SimpleAudioVolume.GetMasterVolume()
                                    confirmed_volume = int(current_vol * 100)
                                    # Send back updated volume
                                    message = f"{app_name},{confirmed_volume}\n"
                                    ser.write(message.encode())
                                    ser.flush()
                                    break
                except ValueError:
                    print(f"Invalid volume value: {parts[1]}")
        else:
            print(f"ESP32: {line}")

def get_app_volume(session):
    """Utility to get the float volume [0..1] for a PyCAW Session."""
    return session.SimpleAudioVolume.GetMasterVolume()

def icon_to_image_with_mask(hicon):
    """
    Extract an icon from a Windows HICON handle
    and return a Pillow Image in RGB.
    """
    width = win32api.GetSystemMetrics(win32con.SM_CXICON)
    height = win32api.GetSystemMetrics(win32con.SM_CYICON)

    hdc_screen = win32gui.GetDC(0)
    dc_screen = win32ui.CreateDCFromHandle(hdc_screen)

    color_dc = dc_screen.CreateCompatibleDC()
    color_bmp = win32ui.CreateBitmap()
    color_bmp.CreateCompatibleBitmap(dc_screen, width, height)
    color_dc.SelectObject(color_bmp)

    color_dc.FillSolidRect((0, 0, width, height), 0x000000)
    win32gui.DrawIconEx(color_dc.GetSafeHdc(), 0, 0, hicon, width, height, 0, None, win32con.DI_NORMAL)

    mask_dc = dc_screen.CreateCompatibleDC()
    mask_bmp = win32ui.CreateBitmap()
    mask_bmp.CreateCompatibleBitmap(dc_screen, width, height)
    mask_dc.SelectObject(mask_bmp)

    mask_dc.FillSolidRect((0, 0, width, height), 0xFFFFFF)
    win32gui.DrawIconEx(mask_dc.GetSafeHdc(), 0, 0, hicon, width, height, 0, None, win32con.DI_MASK)

    color_info = color_bmp.GetInfo()
    color_str = color_bmp.GetBitmapBits(True)
    color_im = Image.frombuffer('RGB',
                                (color_info['bmWidth'], color_info['bmHeight']),
                                color_str, 'raw', 'BGRX', 0, 1)

    # Clean up
    mask_dc.DeleteDC()
    color_dc.DeleteDC()
    dc_screen.DeleteDC()
    win32gui.ReleaseDC(0, hdc_screen)
    win32gui.DestroyIcon(hicon)

    return color_im

def extract_icon(exe_path, save_path):
    """
    Extracts the primary icon from a given .exe file and saves as PNG.
    """
    try:
        large_icons, _ = win32gui.ExtractIconEx(exe_path, 0)
        if large_icons:
            hicon = large_icons[0]
            image = icon_to_image_with_mask(hicon)
            image.save(save_path)
        else:
            print(f"No icon found for {exe_path}")
    except Exception as e:
        print(f"Error extracting icon from {exe_path}: {e}")

def fetch_app_icons():
    """
    Fetch icons for all active audio apps and store them in './app_icons'.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    destination_folder = os.path.join(script_dir, "app_icons")
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        process = session.Process
        if process:
            try:
                process_name = process.name().removesuffix(".exe")
                exe_path = process.exe()
                if exe_path and os.path.exists(exe_path):
                    icon_path = os.path.join(destination_folder, f"{process_name}.png")
                    extract_icon(exe_path, icon_path)
                else:
                    print(f"Executable path not found for {process_name}")
            except Exception as e:
                print(f"Error processing {process.name()}: {e}")

def rgb_to_rgb565(r, g, b):
    """
    Convert an RGB888 pixel to RGB565 (two-byte) format.
    """
    return ((r & 0xF8) << 8) | ((g & 0xFC) << 3) | (b >> 3)

def open_app_icon(sprite):
    """
    Open the given PNG icon, resize to 32x32, and convert to 2-byte RGB565 format.
    Returns a bytearray of length 2048.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_dir, "app_icons", sprite + ".png")

    image = Image.open(image_path)
    image = image.convert("RGB").resize((32, 32))
    sprite_data = bytearray()
    for y in range(32):
        for x in range(32):
            r, g, b = image.getpixel((x, y))
            pixel565 = rgb_to_rgb565(r, g, b)
            sprite_data.extend(pixel565.to_bytes(2, byteorder="little"))
    return sprite_data

def Chunk_send(ser, data):
    """
    Sends 'data' to the ESP32 in chunks to avoid buffer overflow.
    Using a bigger chunk_size (512) and shorter delay (0.01s).
    """
    chunk_size = 512
    total_bytes = len(data)
    idx = 0
    while idx < total_bytes:
        chunk = data[idx:idx+chunk_size]
        ser.write(chunk)
        ser.flush()
        idx += chunk_size
        # Adjust or remove the delay below if needed
        time.sleep(0.01)
    print("[PC] Chunks sent.")

def wait_for_ready_signal(ser):
    """
    Wait until ESP32 prints "READY_TO_RECEIVE"
    """
    while True:
        if ser.in_waiting > 0:
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            if line == "READY_TO_RECEIVE":
                print("[PC] ESP32 is ready to receive data.")
                return
        time.sleep(0.05)

def Handshake(ser):
    """
    Wait until ESP32 prints "READY", then respond "OK"
    """
    ser.reset_input_buffer()
    while True:
        if ser.in_waiting:
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            print("[PC] Received line:", line)
            if line == "READY":
                print("[PC] ESP32 is ready!")
                ser.write(b"OK\n")
                ser.flush()
                break
        time.sleep(0.05)

def Initialize_apps(ser):
    """
    1) Fetch the list of apps playing audio and their icons.
    2) Send "Initialising apps" command to ESP32.
    3) For each app, send "Start_app" + name + volume + icon data + "End_app".
    4) End with "Done".
    """
    sessions = AudioUtilities.GetAllSessions()
    fetch_app_icons()

    ser.write(b"Initialising apps\n")

    # 1) Wait until the ESP32 says "READY_TO_RECEIVE"
    wait_for_ready_signal(ser)

    for session in sessions:
        process = session.Process
        if not process:
            continue

        app_name = process.name().removesuffix(".exe")

        ser.write(b"Start_app\n")

        # Send app name
        ser.write((app_name + "\n").encode())

        # Send volume
        vol_percent = int(get_app_volume(session) * 100)
        volume_str = str(vol_percent) + "\n"
        ser.write(volume_str.encode())

        # Send icon data
        sprite_data = open_app_icon(app_name)
        Chunk_send(ser, sprite_data)

        # End this app's transmission
        ser.write(b"End_app\n")

    print("[PC] Done sending apps.")
    ser.write(b"Done\n")

def check_for_added_removed_apps(ser):
    """
    Detects new and removed apps, notifying the ESP32.
    """
    global last_known_apps

    # Get the latest set of active audio apps
    sessions = AudioUtilities.GetAllSessions()
    current_apps = set()
    app_data = {}  # Store volume and icon data

    for session in sessions:
        process = session.Process
        if process:
            app_name = process.name().removesuffix(".exe")
            current_apps.add(app_name)

            # Store volume level
            volume_level = int(get_app_volume(session) * 100)
            app_data[app_name] = volume_level

    # Detect removed apps (previously known, but now missing)
    removed_apps = last_known_apps - current_apps
    for app in removed_apps:
        print(f"[PC] App closed: {app}. Notifying ESP32...")
        ser.write(b"Remove_app\n")
        ser.write((app + "\n").encode())
        ser.flush()

    # Detect new apps (not in last_known_apps but now in current_apps)
    new_apps = current_apps - last_known_apps
    for app in new_apps:
        print(f"[PC] New app detected: {app}. Sending to ESP32...")

        ser.write(b"New_app\n")
        ser.write((app + "\n").encode())  # Send app name
        ser.write((str(app_data[app]) + "\n").encode())  # Send volume

        # Send app icon data
        sprite_data = open_app_icon(app)
        Chunk_send(ser, sprite_data)

        ser.write(b"End_app\n")  # Indicate end of app transmission

    # Update known apps list
    last_known_apps = current_apps

def main_loop():
    """
    One cycle of the main routine: handshake, initialize apps, parse lines, watch for added/removed apps.
    If the port is disconnected, let the caller handle it (raising an exception).
    """
    # 1) Find and open the port
    ser = find_esp32_port()  # Returns a serial.Serial object

    # 2) Perform handshake
    Handshake(ser)

    # 3) Initialize apps
    Initialize_apps(ser)

    # 4) Populate last_known_apps
    global last_known_apps
    last_known_apps.clear()
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        p = session.Process
        if p:
            last_known_apps.add(p.name().removesuffix(".exe"))

    # 5) Main loop
    last_check_time = time.time()
    check_interval = 2.0  # seconds

    while True:
        parse_line(ser)  # Handle incoming serial data

        current_time = time.time()
        if current_time - last_check_time >= check_interval:
            check_for_added_removed_apps(ser)
            last_check_time = current_time

        time.sleep(0.00005)

def run_forever():
    """
    Runs the main_loop forever, automatically restarting if the port is disconnected.
    """
    while True:
        try:
            main_loop()
        except (serial.SerialException, OSError) as ex:
            # Port disconnected or other I/O error.
            print(f"Serial error: {ex}. Restarting...")
            time.sleep(2.0)
            continue  # back to find_esp32_port()
        except KeyboardInterrupt:
            print("Program terminated by user.")
            break  # exit altogether

if __name__ == "__main__":
    run_forever()
