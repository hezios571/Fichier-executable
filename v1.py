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
    # Le COM port doit être disponible (ne pas ouvrir l'ESP32 sur un autre programme)
    port = 'COM4'
    baud_rate = 57600
    try:
        ser = serial.Serial(port, baud_rate, timeout=1)
    except Exception as e:
        print(f"Error opening serial port: {e}")
        return
    return ser

def set_app_volume(app_name: str, volume: float):
    """
    Change le volume d'une application donnée.

    Paramètres:
      - app_name (str): Nom de l'application
      - volume (float): Niveau de volume (0.0 à 1.0)
    """
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        process = session.Process
        if process:
            process_name = process.name()
            # Compare avec et sans ".exe"
            if (process_name.lower() == app_name.lower() or
               process_name.lower().removesuffix(".exe") == app_name.lower()):
                volume_control = session.SimpleAudioVolume
                volume_control.SetMasterVolume(volume, None)
                print(f"Volume de {process_name} défini à {volume * 100:.0f}%")
                return

    print(f"Application '{app_name}' not found or not playing audio.")

def icon_to_image_with_mask(hicon):
    """
    Convertit un handle d'icône Windows (hicon) en objet PIL.Image,
    en gérant la transparence via un masque.
    """
    width = win32api.GetSystemMetrics(win32con.SM_CXICON)
    height = win32api.GetSystemMetrics(win32con.SM_CYICON)

    # DC écran
    hdc_screen = win32gui.GetDC(0)
    dc_screen = win32ui.CreateDCFromHandle(hdc_screen)

    # DC couleur
    color_dc = dc_screen.CreateCompatibleDC()
    color_bmp = win32ui.CreateBitmap()
    color_bmp.CreateCompatibleBitmap(dc_screen, width, height)
    color_dc.SelectObject(color_bmp)

    # Remplit le fond en noir (0x000000)
    color_dc.FillSolidRect((0, 0, width, height), 0x000000)
    # Dessine l'icône en mode normal
    win32gui.DrawIconEx(color_dc.GetSafeHdc(), 0, 0, hicon, width, height, 0, None, win32con.DI_NORMAL)

    # DC pour le masque
    mask_dc = dc_screen.CreateCompatibleDC()
    mask_bmp = win32ui.CreateBitmap()
    mask_bmp.CreateCompatibleBitmap(dc_screen, width, height)
    mask_dc.SelectObject(mask_bmp)

    # Remplit le fond en blanc
    mask_dc.FillSolidRect((0, 0, width, height), 0xFFFFFF)
    # Dessine le masque
    win32gui.DrawIconEx(mask_dc.GetSafeHdc(), 0, 0, hicon, width, height, 0, None, win32con.DI_MASK)

    # Convertit DC → PIL
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

    # Libère GDI
    mask_dc.DeleteDC()
    color_dc.DeleteDC()
    dc_screen.DeleteDC()
    win32gui.ReleaseDC(0, hdc_screen)
    win32gui.DestroyIcon(hicon)

    # Fusion pixel par pixel
    pix_color = color_im.load()
    pix_mask  = mask_im.load()
    for y in range(height):
        for x in range(width):
            r_m, g_m, b_m = pix_mask[x, y]
            # Ajustez selon la logique de masque désirée
            if r_m < 128 and g_m < 128 and b_m < 128:
                # Considéré transparent
                pass
            else:
                # Considéré opaque
                pass

    return color_im

def extract_icon(exe_path, save_path):
    """
    Extrait la première icône large d'un exécutable et la sauvegarde en image PNG.
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
    Extrait les icônes pour les applications audio actives et les
    sauvegarde dans un dossier 'app_icons' au même endroit que le script.
    """
    # Chemin du dossier 'app_icons' dans le même répertoire que ce script
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
    # Convertit un triplet (R, G, B) en format 16 bits 565
    return ((r & 0xF8) << 8) | ((g & 0xFC) << 3) | (b >> 3)

def open_app_icon(sprite):
    """
    Ouvre l'icône "sprite" depuis 'app_icons',
    la convertit en RGB565, et renvoie les données sous forme de bytes.
    """
    # Chemin du dossier 'app_icons' dans le même répertoire que le script V1
    script_dir = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_dir, "app_icons", sprite + ".png")

    image = Image.open(image_path)
    image = image.convert("RGB").resize((32, 32))
    sprite_data = bytearray()
    for y in range(32):
        for x in range(32):
            r, g, b = image.getpixel((x, y))
            pixel565 = rgb_to_rgb565(r, g, b)
            sprite_data.extend(pixel565.to_bytes(2, byteorder="little")) #little endian pour que l'application affiche correctement
    return sprite_data

def Receive_app_volume():
    """
    Boucle principale pour lire les commandes reçues de l'ESP32.
    Format attendu: "app_name,volume"
    """
    try:
        while True:
            if ser.in_waiting:  # s'il y a du data dans le buffer
                try:
                    line = ser.readline().decode('utf-8').strip()
                    print(f"Received: {line}")
                except Exception as e:
                    print(f"Error decoding serial data: {e}")
                    continue

                if ',' in line:  # commande pour modifier le son
                    parts = line.split(',')
                    if len(parts) == 2:
                        app_name = parts[0].strip()
                        try:
                            volume = float(parts[1].strip())
                            if 0.0 <= volume <= 1.0:
                                set_app_volume(app_name, volume)
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

def Chunk_send(data):
    """
    Découpe les données de l'icône en blocs de taille 'chunk_size'
    et les envoie à l'ESP32.
    """
    chunk_size = 240
    total_bytes = len(data)
    for i in range(0, total_bytes, chunk_size):
        chunk = data[i:i+chunk_size]
        n = ser.write(chunk)
        ser.flush()
        time.sleep(0.05)
    print("Chunks sent")

def wait_for_ready_signal(ser):
    """
    Attend le signal "READY_TO_RECEIVE_SPRITE" depuis l'ESP32.
    """
    while True:
        if ser.in_waiting > 0:
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            if line == "READY_TO_RECEIVE_SPRITE":
                print("[PC] ESP32 is ready to receive sprite data.")
                return
        time.sleep(0.1)

def Handshake():
    """
    Effectue un handshake initial avec l'ESP32:
      - Attend un message "READY"
      - Répond "OK"
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

def send_app_icons():
    sessions = AudioUtilities.GetAllSessions()
    fetch_app_icons()
    wait_for_ready_signal(ser)
    print("Sending app icons...")
    first_iteration = True
    for session in sessions:
        app = session.Process
        if app==None:
            break
        else:
            if first_iteration:
                first_iteration = False  
            elif app.name() == app_name + ".exe":
                continue
            app_name=app.name().removesuffix(".exe")
            sprite_data = open_app_icon(app_name)
            Chunk_send(sprite_data)
    print("Done sending icons")




if __name__ == "__main__":
    ser = init()
    Handshake()
    send_app_icons()
    ser.close()
