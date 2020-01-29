import win32gui, win32process
import psutil

from time import sleep
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume


def set_volume(option=None):
    sessions = AudioUtilities.GetAllSessions()  # Lists all of the running programs but it has sound output
    for session in sessions:
        if session.Process and session.Process.name() == "Spotify.exe":
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)

            if option == "open":
                volume.SetMasterVolume(1, None)  # Opens the sound
            elif option == "close":
                volume.SetMasterVolume(0, None)  # Closes the sound

            return volume.GetMasterVolume()  # Returns the volume value


def winEnumHandler(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        pid = win32process.GetWindowThreadProcessId(hwnd)

        # win32gui.GetWindowText(hwnd)  ##Returns title name
        #                               ##In Spotify's title name is name of the song we are listening or advertisement
        # psutil.Process(pid[1]).name() ##Returns process name

        program = dict()
        program["title_name"] = win32gui.GetWindowText(hwnd)
        program["exe_name"] = psutil.Process(pid[1]).name()
        programs.append(program)  # Appends all of the running programs in dict type


while True:
    programs = list()
    win32gui.EnumWindows(winEnumHandler, None)  # Lists all of the running programs

    for program in programs:
        if program["exe_name"] == "Spotify.exe" and program["title_name"] != "":
            spotify_song = program["title_name"]

            if spotify_song == "Advertisement":
                if set_volume() == 1.0:
                    set_volume("close")
                # print("Advertisement!")
            else:
                if set_volume() == 0.0:
                    set_volume("open")
    sleep(0.1)