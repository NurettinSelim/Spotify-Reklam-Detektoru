import win32gui, win32process
import psutil

from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume


def ses_ayarlama(islem=None):
    sessions = AudioUtilities.GetAllSessions()  # Bütün ses çıkışı olan programları listeliyor.
    for session in sessions:
        if session.Process and session.Process.name() == "Spotify.exe":
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)

            if islem == "ac":
                volume.SetMasterVolume(1, None)  # Sesi açar.
            elif islem == "kapat":
                volume.SetMasterVolume(0, None)  # Sesi kapar.

            return volume.GetMasterVolume()  # Şu andaki sesi döndürür.


def winEnumHandler(hwnd, ctx):
    if win32gui.IsWindowVisible(hwnd):
        pid = win32process.GetWindowThreadProcessId(hwnd)

        # win32gui.GetWindowText(hwnd)  ##Pencere adını veriyor.
        #                               ##Spotify penceresinin adında da dinlediğimiz şarkı yazıyor.
        # psutil.Process(pid[1]).name() ##Pencerenin hangi programa ait olduğunu veriyor.

        program = dict()
        program["title_name"] = win32gui.GetWindowText(hwnd)
        program["exe_name"] = psutil.Process(pid[1]).name()
        programs.append(program)  # Bütün programları sözlük şeklinde programs listesine ekliyor.


while True:
    programs = list()
    win32gui.EnumWindows(winEnumHandler, None)  # Çalışan tüm programları listeliyor.

    for program in programs:
        if program["exe_name"] == "Spotify.exe" and program["title_name"] != "":
            spotify_song = program["title_name"]

            if spotify_song == "Advertisement":
                if ses_ayarlama() == 1.0:
                    ses_ayarlama("kapat")
                # print("Reklam!")
            else:
                if ses_ayarlama() == 0.0:
                    ses_ayarlama("ac")
