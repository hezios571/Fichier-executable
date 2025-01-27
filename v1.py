from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
from comtypes import CLSCTX_ALL


def list_audio_sessions():
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        process = session.Process
        if process:
           
            process_name = process.name().removesuffix(".exe")  
          
            volume = session.SimpleAudioVolume
            print(f"App: {process_name} | Volume: {volume.GetMasterVolume():.2f}")
