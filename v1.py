from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
from comtypes import CLSCTX_ALL


def list_audio_sessions():
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        process = session.Process
        if process:
            # Remove '.exe' from the process name
            process_name = process.name().removesuffix(".exe")  # For Python 3.9+
            # For older Python versions:
            # process_name = process.name()[:-4] if process.name().lower().endswith(".exe") else process.name()
            volume = session.SimpleAudioVolume
            print(f"App: {process_name} | Volume: {volume.GetMasterVolume():.2f}")
