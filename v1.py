from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import os
from PIL import Image
import win32gui
import win32api
import win32ui
import win32con
import serial
import time
import ctypes
import ctypes.wintypes
import numpy as np

def init():
    """
    Initialize the serial connection to the ESP32.
    Increase baud rate to 115200 for better throughput.
    """
    port = 'COM4'
    baud_rate = 460800  # Increase from 57600
    try:
        ser = serial.Serial(port, baud_rate, timeout=0.1)  # shorter timeout
        return ser
    except Exception as e:
        print(f"Error opening serial port: {e}")
        return None

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
                                    #print(f"[PC -> ESP32] Synced volume: {message.strip()}")
                                    break
                    else:
                        print(f"Volume out of range: {vol_percent}")
                except ValueError:
                    print(f"Invalid volume value: {parts[1]}")
            else:
                print(f"Invalid format. Expected 'app_name,volume'. Got: {line}")
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

if __name__ == "__main__":
    ser = init()
    if not ser:
        print("Failed to open serial port. Exiting.")
        exit(1)

    try:
        # Step 1: Handshake
        Handshake(ser)

        # Step 2: Send apps & icons
        Initialize_apps(ser)

        # Step 3: Continuously parse lines from ESP32
        print("[PC] Listening for volume-change commands from ESP32...")
        while True:
            parse_line(ser)
            # lower the sleep for better responsiveness
            time.sleep(0.00005)

    except KeyboardInterrupt:
        print("Program terminated by user.")
    finally:
        ser.close()
