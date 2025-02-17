from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
from comtypes import CLSCTX_ALL
import os
from PIL import Image
import win32gui, win32api, win32ui, win32con
import serial
import time
import ctypes
import ctypes.wintypes
import numpy as np

def init():
    port = 'COM4'
    baud_rate = 57600
    try:
        ser = serial.Serial(port, baud_rate, timeout=1)
        return ser
    except Exception as e:
        print(f"Error opening serial port: {e}")
        return None

def read_serial():
    """Continuously reads and prints serial data from the ESP32."""
    while True:
        if ser.in_waiting:  # Check if data is available
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            print(f"ESP32: {line}")

def set_app_volume(app_name: str, volume: float):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        process = session.Process
        if process:
            process_name = process.name()
            # Compare with and without ".exe"
            if (process_name.lower() == app_name.lower() or
               process_name.lower().removesuffix(".exe") == app_name.lower()):
                volume_control = session.SimpleAudioVolume
                volume_control.SetMasterVolume(volume, None)
                return
    print(f"Application '{app_name}' not found or not playing audio.")

def Receive_app_volume():
    try:
        while True:
            if ser.in_waiting:
                try:
                    line = ser.readline().decode('utf-8').strip()
                    print(f"Received: {line}")
                except Exception as e:
                    print(f"Error decoding serial data: {e}")
                    continue

                # Expect command in "app_name,volume" format
                if ',' in line:
                    parts = line.split(',')
                    if len(parts) == 2:
                        app_name = parts[0].strip()
                        try:
                            # Assume incoming volume is given as integer percentage (0-100)
                            vol_percent = int(parts[1].strip())
                            volume = vol_percent / 100.0
                            if 0.0 <= volume <= 1.0:
                                set_app_volume(app_name, volume)
                                # Now find the session and get the actual volume
                                sessions = AudioUtilities.GetAllSessions()
                                for session in sessions:
                                    process = session.Process
                                    if process:
                                        process_name = process.name()
                                        if (process_name.lower() == app_name.lower() or
                                            process_name.lower().removesuffix(".exe") == app_name.lower()):
                                            current_vol = session.SimpleAudioVolume.GetMasterVolume()
                                            # Convert back to percentage (integer)
                                            confirmed_volume = int(current_vol * 100)
                                            # Send back the updated volume in the same format: "app_name,volume"
                                            message = f"{app_name},{confirmed_volume}\n"
                                            ser.write(message.encode())
                                            ser.flush()
                                            print(f"Sent sync: {message.strip()}")
                                            break
                            else:
                                print("Volume value out of range (0.0 to 1.0).")
                        except ValueError:
                            print(f"Invalid volume value: {parts[1]}")
                    else:
                        print("Invalid command format. Expected 'app_name,volume'.")
                else:
                    print("Received data does not match the expected command format.")
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("Program terminated by user.")
    finally:
        ser.close()

def get_app_volume(session):
    return session.SimpleAudioVolume.GetMasterVolume()

def icon_to_image_with_mask(hicon):
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

    mask_info = mask_bmp.GetInfo()
    mask_str = mask_bmp.GetBitmapBits(True)
    mask_im = Image.frombuffer('RGB',
                               (mask_info['bmWidth'], mask_info['bmHeight']),
                               mask_str, 'raw', 'BGRX', 0, 1)

    mask_dc.DeleteDC()
    color_dc.DeleteDC()
    dc_screen.DeleteDC()
    win32gui.ReleaseDC(0, hdc_screen)
    win32gui.DestroyIcon(hicon)

    pix_color = color_im.load()
    pix_mask  = mask_im.load()
    for y in range(height):
        for x in range(width):
            r_m, g_m, b_m = pix_mask[x, y]
            # Adjust your transparency logic as needed here
            pass
    return color_im

def extract_icon(exe_path, save_path):
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
    return ((r & 0xF8) << 8) | ((g & 0xFC) << 3) | (b >> 3)

def open_app_icon(sprite):
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

def Chunk_send(data):
    chunk_size = 240
    total_bytes = len(data)
    for i in range(0, total_bytes, chunk_size):
        chunk = data[i:i+chunk_size]
        ser.write(chunk)
        ser.flush()
        time.sleep(0.05)
    print("Chunks sent")

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
        time.sleep(0.1)

def Handshake():
    """
    Wait until ESP32 prints "READY", then respond "OK"
    """
    ser.reset_input_buffer()
    while True:
        if ser.in_waiting:
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            print("Received line:", line)
            if line == "READY":
                print("ESP32 is ready!")
                ser.write(b"OK\n")
                ser.flush()
                break
        time.sleep(0.1)

def Initialize_apps():

    sessions = AudioUtilities.GetAllSessions()
    fetch_app_icons()
    ser.write(b"Initialising apps\n")
    # Step 1: Wait until ESP32 says "READY_TO_RECEIVE"
    wait_for_ready_signal(ser)

    # Step 2: Send "Initialising apps"
    

    first_iteration = True
    total_processes = len(sessions)
    index = 0

    for session in sessions:
        app = session.Process
        index += 1

        # If at the end of the sessions list, stop
        if total_processes == index+1:
            break

        if app is None:
            continue

        if first_iteration:
            first_iteration = False
        elif app.name() == app_name + ".exe":
            continue

        ser.write(b"Start_app\n")

        app_name = app.name().removesuffix(".exe")
        ser.write((app_name + "\n").encode())

        volume_str = str(int(get_app_volume(session)*100)) + "\n"
        ser.write(volume_str.encode())

        sprite_data = open_app_icon(app_name)
        Chunk_send(sprite_data)

        ser.write(b"End_app\n")

    print("Done sending apps")
    ser.write(b"Done\n")

# def parse_line():
#     if (ser.read()==""):
#         #do this
        

if __name__ == "__main__":
    ser = init()

    Handshake()        # Step 1: Do handshake
    Initialize_apps()  # Step 2: Send apps & icons

    # Step 3: Print everything from ESP32
    read_serial()
    #while (True): parse line pour les changement de volume

    ser.close()
