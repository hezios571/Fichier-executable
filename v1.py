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
    #Le COM4 port doit être disponible donc il ne faut pas ouvrir le serial monitor de esp32
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
    Cette fonction peremt de changer le volume d'une fonction en mettant en paramètres:
    1. Nom de l'application
    2. Volume désiré

    La fonction affiche un erreur si l'application est introuvable
    """
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        process = session.Process
        if process:
            process_name = process.name()  # Garde le nom original
            
            # On compare avec et sans ".exe" pour éviter les erreurs
            if process_name.lower() == app_name.lower() or process_name.lower().removesuffix(".exe") == app_name.lower():
                volume_control = session.SimpleAudioVolume
                volume_control.SetMasterVolume(volume, None)
                print(f"Volume de {process_name} défini à {volume * 100:.0f}%")
                return
            
    print(f"Application '{app_name}' not found or not playing audio.")

def icon_to_image_with_mask(hicon):
    """
    Récupère l'icône Windows en 2 étapes:
    1) L'image couleur normale (DI_NORMAL)
    2) Le masque (DI_MASK)
    Puis fusionne pour marquer en magenta les zones réellement transparentes.
    """
    # 1. Dimensions standard de l'icône Windows
    width = win32api.GetSystemMetrics(win32con.SM_CXICON)
    height = win32api.GetSystemMetrics(win32con.SM_CYICON)

    # 2. Prépare un DC "couleur"
    hdc_screen = win32gui.GetDC(0)
    dc_screen = win32ui.CreateDCFromHandle(hdc_screen)

    color_dc = dc_screen.CreateCompatibleDC()
    color_bmp = win32ui.CreateBitmap()
    color_bmp.CreateCompatibleBitmap(dc_screen, width, height)
    color_dc.SelectObject(color_bmp)

    # Remplit le fond en noir(0xFFFFFF) qui va match le fond noir
    color_dc.FillSolidRect((0, 0, width, height), 0x000000)

    # Dessine l'icône en mode normal (DI_NORMAL)
    win32gui.DrawIconEx(color_dc.GetSafeHdc(), 0, 0, hicon, width, height, 0, None, win32con.DI_NORMAL)

    # 3. Prépare un DC pour le masque
    mask_dc = dc_screen.CreateCompatibleDC()
    mask_bmp = win32ui.CreateBitmap()
    mask_bmp.CreateCompatibleBitmap(dc_screen, width, height)
    mask_dc.SelectObject(mask_bmp)

    # Remplit le fond en blanc (par exemple)
    mask_dc.FillSolidRect((0, 0, width, height), 0xFFFFFF)

    # Dessine le masque (DI_MASK)
    win32gui.DrawIconEx(mask_dc.GetSafeHdc(), 0, 0, hicon, width, height, 0, None, win32con.DI_MASK)

    # 4. Convertit les deux DC en images PIL
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

    # 5. Libère les ressources GDI
    mask_dc.DeleteDC()
    color_dc.DeleteDC()
    dc_screen.DeleteDC()
    win32gui.ReleaseDC(0, hdc_screen)
    win32gui.DestroyIcon(hicon)

    # 6. Fusion pixel par pixel (zone noire dans mask = transparent, zone blanche = opaque)
    #    - Souvent, DI_MASK met en noir les zones transparentes, et blanc les zones opaques,
    #      mais parfois c’est l’inverse. Il faut tester.
    pix_color = color_im.load()
    pix_mask  = mask_im.load()
    for y in range(height):
        for x in range(width):
            # Récupère la valeur RGB du masque
            r_m, g_m, b_m = pix_mask[x, y]
            # Si le masque est noir => transparent => on laisse magenta
            # Si le masque est blanc => c'est la zone opaque => on garde la couleur
            # Inverse si nécessaire en fonction du rendu que vous observez
            if r_m < 128 and g_m < 128 and b_m < 128:
                # Ici, on considère pixel transparent => on garde magenta
                pass
            else:
                # Zone opaque => on garde le pixel de color_im tel quel
                pass

    return color_im

def extract_icon(exe_path, save_path):
    try:
        large_icons, _ = win32gui.ExtractIconEx(exe_path, 0)
        if large_icons:
            hicon = large_icons[0]
            # Convertit l'icône en PIL avec masque
            image = icon_to_image_with_mask(hicon)
            image.save(save_path)
        else:
            print(f"No icon found for {exe_path}")
    except Exception as e:
        print(f"Error extracting icon from {exe_path}: {e}")

def fetch_app_icons():
    """
    Extracts the icons for active audio sessions and saves them to a folder on the desktop.
    """
    
    destination_folder = r"C:\Users\Jayro\Desktop\app_icons"
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    #print(f"Icons will be saved in: {destination_folder}")

    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        process = session.Process
        if process:
            try:
                process_name = process.name().removesuffix(".exe")
                exe_path = process.exe()  # Get the path to the executable
                if exe_path and os.path.exists(exe_path):
                    icon_path = os.path.join(destination_folder, f"{process_name}.png")
                    extract_icon(exe_path, icon_path)
                else:
                    print(f"Executable path not found for {process_name}")
            except Exception as e:
                print(f"Error processing {process.name()}: {e}")

def rgb_to_rgb565(r, g, b):
    # Standard 565 arrangement: R in bits [15..11], G [10..5], B [4..0]
    # For an int in Python, that means:
    return ((r & 0xF8) << 8) | ((g & 0xFC) << 3) | (b >> 3)

def open_app_icon():
    image_path = r"C:\Users\Jayro\Desktop\app_icons\Razer Synapse 3.png"
    image = Image.open(image_path)
    image = image.convert("RGB").resize((32, 32))
    sprite_data = bytearray()
    for y in range(32):
        for x in range(32):
            r, g, b = image.getpixel((x, y))
            pixel565 = rgb_to_rgb565(r, g, b)
            # Store this 16-bit value as 2 bytes in LITTLE-ENDIAN order:
            sprite_data.extend(pixel565.to_bytes(2, byteorder="little"))
    return sprite_data

def Receive_app_volume():
    try:
        while True:
            if ser.in_waiting: #si ser.in.waiting=/0, il y a du data dans le buffer
                try:
                    line = ser.readline().decode('utf-8').strip() #Lis un ligne de bytes jusqu'au newline et enlève les espaces vides
                    print(f"Received: {line}")
                except Exception as e:
                    print(f"Error decoding serial data: {e}")
                    continue

                if ',' in line: #s'il y a une virgule, c'est une commande pour modifier le son
                    parts = line.split(',') #split la ligne en deux partie, avant et après la virgule
                    if len(parts) == 2: #s'assure qu'il y a juste deux parties (appname et volume)
                        app_name = parts[0].strip()
                        try:
                            volume = float(parts[1].strip())
                            
                            if 0.0 <= volume <= 1.0: #s'assurer que le volume est dans le range accepté
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

def Chunk_send(sprite_data):

    chunk_size = 240  # Send in chunks that are smaller than or equal to the RX buffer size
    total_bytes = len(sprite_data) 
    for i in range(0, total_bytes, chunk_size):
                sprite_data = open_app_icon()
                
                chunk = sprite_data[i:i+chunk_size]
                n = ser.write(chunk)
                ser.flush()
                print(f"Chunk {i//chunk_size + 1}: {n} bytes written")
                time.sleep(0.05)  # A short delay between chunks
    print("Done sending app icon")
                
def wait_for_ready_signal(ser):
    
    while True:
        if ser.in_waiting > 0:
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            if line == "READY_TO_RECEIVE_SPRITE":
                print("[PC] ESP32 is ready to receive sprite data.")
                return
        time.sleep(0.1)

def Handshake():
    ser.reset_input_buffer() 
    while True:
        if ser.in_waiting:
            line = ser.readline().decode('utf-8', errors='ignore').strip()
            print("Received line:", line)
            if line == "READY":
                print("ESP32 is ready!")
                ser.write(b"OK\n")
                ser.flush()  # Ensure the data is sent immediately
                break
        time.sleep(0.1)

if __name__ == "__main__":
    ser=init()
    Handshake()
    wait_for_ready_signal(ser)
    print("Sending app icons")    
    sprite_data = open_app_icon() #temporaire
    fetch_app_icons()
    
    Chunk_send(sprite_data)

    ser.close()
    
        