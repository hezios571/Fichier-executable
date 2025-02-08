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

port = 'COM4'
baud_rate = 115200

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

def icon_to_image(hicon):
    # Get system icon dimensions
    width = win32api.GetSystemMetrics(win32con.SM_CXICON)
    height = win32api.GetSystemMetrics(win32con.SM_CYICON)
    
    # Get the device context for the entire screen
    hdc = win32gui.GetDC(0)
    
    # Create a device context from this handle
    dc = win32ui.CreateDCFromHandle(hdc)
    
    # Create a memory device context
    memdc = dc.CreateCompatibleDC()
    
    # Create a bitmap object
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(dc, width, height)
    
    # Select the bitmap object into the memory DC
    memdc.SelectObject(bmp)
    
    # Fill the background with a transparent color (optional)
    memdc.FillSolidRect((0, 0, width, height), 0xFFFFFF)  # White background; change if needed
    
    # Draw the icon into the memory DC
    win32gui.DrawIconEx(memdc.GetSafeHdc(), 0, 0, hicon, width, height, 0, None, win32con.DI_NORMAL)
    
    # Get bitmap information and bits
    bmpinfo = bmp.GetInfo()
    bmpstr = bmp.GetBitmapBits(True)
    
    # Create a PIL image from the bitmap bits
    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1
    )
    
    # Clean up: release DCs and delete objects
    memdc.DeleteDC()
    dc.DeleteDC()
    win32gui.ReleaseDC(0, hdc)
    win32gui.DestroyIcon(hicon)  # Destroy the icon handle
    
    return im

def extract_icon(exe_path, save_path):
    try:
        # Extract the icon handles (large and small)
        large_icons, _ = win32gui.ExtractIconEx(exe_path, 0)
        if large_icons:
            hicon = large_icons[0]
            # Convert the icon handle to a PIL image
            image = icon_to_image(hicon)
            image.save(save_path)
            #print(f"Icon saved: {save_path}")
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

def rgb_to_rgb565(r, g, b):#sert a la conversion du fichier png compressé
    return ((r & 0xF8) << 8) | ((g & 0xFC) << 3) | (b >> 3)

def open_app_icon():
    image_path = r"C:\Users\Jayro\Desktop\app_icons\Razer Synapse 3.png"
    image = Image.open(image_path)
    image = image.convert("RGB").resize((32, 32))
    sprite_data = bytearray()
    # Process each pixel in the 32x32 image
    for y in range(32):
        for x in range(32):
            r, g, b = image.getpixel((x, y))
            pixel565 = rgb_to_rgb565(r, g, b)
            # Convert the 16-bit value to 2 bytes (big-endian) and append to our data
            sprite_data.extend(pixel565.to_bytes(2, byteorder="big"))
            
    return sprite_data

def Receive_app_volume():
#Le COM4 port doit être disponible donc il ne faut pas ouvrir le serial monitor de esp32
    try:
        ser = serial.Serial(port, baud_rate, timeout=1) #attend 1 seconde, s'il n'y a pas de data sur le port = timeout
        print(f"Connected to {port} at {baud_rate} baud.")
    except Exception as e:
        print(f"Error opening serial port: {e}")
        return
    """
    Possibilité d'utiliser des event interrupts pour rendre le programme moins lourd sur le CPU (pySerial-asyncio)
    """
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
            #sert a envoyer les app icons
                    
            #permet de ne pas utiliser trop de ressources du CPU, doit probablement être modifié pour des interrupts?
            time.sleep(0.1)
            print("sent app icon")
            ser.write(open_app_icon())
            ser.close()  
    except KeyboardInterrupt:
        print("Program terminated by user.")
    finally:
        ser.close()

if __name__ == "__main__":
    while True:
        fetch_app_icons()
        port = 'COM4'
        baud_rate = 57600
        ser = serial.Serial(port, baud_rate, timeout=1)
        sprite_data = open_app_icon()
        ser.write(sprite_data)
        ser.close() 
        print("sent app icon")
        time.sleep(2)