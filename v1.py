from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
from comtypes import CLSCTX_ALL
import os
from PIL import Image
import win32gui, win32api
import serial
import time

def list_audio_sessions():
    """
    Cette fonction vérifie les applications qui peuvent émettre du son et montre leur volume associé.
    Celle-ci ne retourne rien

    Pas certain si cette fonction pourras être utile
    """
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        process = session.Process
        if process:
           
            process_name = process.name().removesuffix(".exe")  
          
            volume = session.SimpleAudioVolume
            print(f"App: {process_name} | Volume: {volume.GetMasterVolume():.2f}")

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

def extract_icon(exe_path, save_path):
    """
    Extrait l'icône d'un fichier .exe et la sauvegarde.

    :param exe_path: Chemin du fichier .exe.
    :param save_path: Chemin où sauvegarder l'image.
    """
    try:
        large, _ = win32gui.ExtractIconEx(exe_path, 0)
        if large:
            icon = large[0]
            ico_x = win32api.GetSystemMetrics(49)  # Taille icône standard
            image = Image.new("RGBA", (ico_x, ico_x))
            hdc = image.im.id
            win32gui.DrawIcon(hdc, 0, 0, icon)
            image.save(save_path)
            win32gui.DestroyIcon(icon)
            print(f" Icône sauvegardée : {save_path}")
        else:
            print(f" Pas d’icône trouvée pour {exe_path}")
    except Exception as e:
        print(f" Erreur en extrayant l'icône ({exe_path}): {e}")


def fetch_app_icons():
    """
    Trouve et extrait les icônes des applications audio actives,
    puis les enregistre sur le bureau dans le fichier app_icons.
    """
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    destination_folder = os.path.join(desktop_path, "app_icons")

    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    print(f" Icons will be saved in: {destination_folder}")

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
                    print(f"Chemin non trouvé pour {process_name}")

            except Exception as e:
                print(f"Impossible d'obtenir le chemin de {process_name}: {e}")

def main():
   
    port = 'COM4'
    baud_rate = 115200
#Le COM4 port doit être disponible donc il ne faut pas ouvrir le serial monitor de esp32
    try:
        ser = serial.Serial(port, baud_rate, timeout=1) #attend 1 seconde, s'il n'y a pas de data sur le port = timeout
        print(f"Connected to {port} at {baud_rate} baud.")
    except Exception as e:
        print(f"Error opening serial port: {e}")
        return

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
                    
            #permet de ne pas utiliser trop de ressources du CPU, doit probablement être modifié pour des interrupts?
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("Program terminated by user.")
    finally:
        ser.close()

if __name__ == "__main__":
    main()