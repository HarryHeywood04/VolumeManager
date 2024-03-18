import subprocess
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


def getMonitorCount():
    command = subprocess.Popen('powershell.exe Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams', stdout=subprocess.PIPE)
    systemInfo = str(command.communicate())
    monitorCount = systemInfo.count("Active")
    return monitorCount


def setVolume(mute: bool):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    if mute:
        volume.SetMasterVolumeLevel(-45.0, None)
    else:
        volume.SetMasterVolumeLevel(0.0, None)


if getMonitorCount() == 1:
    setVolume(True)
else:
    setVolume(False)
