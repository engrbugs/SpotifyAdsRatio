import time  # sleep
from pycaw.pycaw import AudioUtilities  # mute

import os
import win32con as wcon
import win32api as wapi
import win32gui as wgui
import win32process as wproc

version = 'b0.6.13'


def enum_windows_proc(wnd, param):
    pid = param.get("pid", None)
    data = param.get("data", None)
    if pid is None or wproc.GetWindowThreadProcessId(wnd)[1] == pid:
        text = wgui.GetWindowText(wnd)
        if text:
            style = wapi.GetWindowLong(wnd, wcon.GWL_STYLE)
            if style & wcon.WS_VISIBLE:
                if data is not None:
                    data.append((wnd, text))
                # else:
                # print("%08X - %s" % (wnd, text))


def enum_process_windows(pid=None):
    data = []
    param = {
        "pid": pid,
        "data": data,
    }
    wgui.EnumWindows(enum_windows_proc, param)
    return data


def _filter_processes(processes, search_name=None):
    if search_name is None:
        return processes
    filtered = []
    for pid, _ in processes:
        try:
            proc = wapi.OpenProcess(wcon.PROCESS_ALL_ACCESS, 0, pid)
        except:
            # print("Process {0:d} couldn't be opened: {1:}".format(pid, traceback.format_exc()))
            continue
        try:
            file_name = wproc.GetModuleFileNameEx(proc, None)
        except:
            # print("Error getting process name: {0:}".format(traceback.format_exc()))
            wapi.CloseHandle(proc)
            continue
        base_name = file_name.split(os.path.sep)[-1]
        if base_name.lower() == search_name.lower():
            filtered.append((pid, file_name))
        wapi.CloseHandle(proc)
    return tuple(filtered)


def enum_processes(process_name=None):
    procs = [(pid, None) for pid in wproc.EnumProcesses()]
    return _filter_processes(procs, search_name=process_name)


def check_window_text(*args):
    global globalI
    proc_name = args[0] if args else None
    procs = enum_processes(process_name=proc_name)
    for pid, name in procs:
        data = enum_process_windows(pid)
        if data:
            proc_text = "PId {0:d}{1:s}windows:".format(pid, " (File: [{0:s}]) ".format(name) if name else " ")
            # print(proc_text)
            for handle, text in data:
                print("{0:d}    {1:d}: [{2:s}]".format(globalI, handle, text))
                return text
    return ''


# variable
spotify_app_exe = 'Spotify.exe'
sleep_time = 1  # sec
spotify_muted = None
globalI = -1
maxGlobalI = 9
start_count_down = 3
paused_text = 'Spotify Free'
from_paused = True


print(f'Spotify Ads Ratio Counter is here by engrbugs {version}')


class AudioController(object):
    def __init__(self, process_name):
        self.process_name = process_name
        self.volume = self.process_volume()

    def mute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                interface.SetMute(1, None)
                print(self.process_name, 'has been muted.')  # debug

    def unmute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                interface.SetMute(0, None)
                print(self.process_name, 'has been unmuted.')  # debug

    def process_volume(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                print('Volume:', interface.GetMasterVolume())  # debug
                return interface.GetMasterVolume()

    def set_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # only set volume in the range 0.0 to 1.0
                self.volume = min(1.0, max(0.0, decibels))
                interface.SetMasterVolume(self.volume, None)
                print('Volume set to', self.volume)  # debug

    def decrease_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                try:
                    # 0.0 is the min value, reduce by decibels
                    self.volume = max(0.0, self.volume - decibels)
                    interface.SetMasterVolume(self.volume, None)
                    print('Volume reduced to', self.volume)  # debug
                except:
                    # 0.0 is the min value, reduce by decibels
                    self.volume = .7
                    interface.SetMasterVolume(self.volume, None)
                    print('Volume reduced to', self.volume)  # debug

    def increase_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                try:
                    # 1.0 is the max value, raise by decibels
                    self.volume = min(1.0, self.volume + decibels)
                    interface.SetMasterVolume(self.volume, None)
                    print('Volume raised to', self.volume)  # debug
                except:
                    # 1.0 is the max value, raise by decibels
                    self.volume = 1
                    interface.SetMasterVolume(self.volume, None)
                    print('Volume raised to', self.volume)  # debug


audio_controller = AudioController(spotify_app_exe)


def main():
    global from_paused
    global audio_controller
    global spotify_muted

    for i in range(start_count_down):
        print('Spotify Ad Muter will initialize in ' + str(start_count_down - i) + '...')
        time.sleep(1)

    app_initialize()

    while True:
        global globalI
        global maxGlobalI
        globalI += 1
        time.sleep(sleep_time)

        title = check_window_text(spotify_app_exe)  # New

        if globalI >= maxGlobalI:
            globalI = -1

        #  if "Advertisement" in titles:  # Spotify window text is named as Advertisement
        if "Advertisement" in title:  # Spotify window text is named as Advertisement
            print('got Advertisement')
            print(audio_controller.process_volume())
            from_paused = None
            if not spotify_muted:
                fade_out()
        elif paused_text in title:
            print('got Spotify Free')
            print(audio_controller.process_volume())
            from_paused = True
            if not spotify_muted:
                fade_out()
        #  elif "Spotify" in titles:  # Named 'Spotify' when ads playing, else it's 'Spotify Free'
        elif "Spotify" in title:
            print('got Spotify')
            print(audio_controller.process_volume())
            from_paused = None
            if not spotify_muted:
                fade_out()
        #  elif "Spotify Free" in titles:  # Named 'Spotify' when ads playing, else it's 'Spotify Free'

        elif "-" in title:
            if spotify_muted:
                fade_in()
        else:
            print('No - seen!')
            print(audio_controller.process_volume())
            from_paused = None
            if not spotify_muted:
                fade_out()


def app_initialize():
    audio_controller.decrease_volume(0.11)
    time.sleep(.1)
    audio_controller.decrease_volume(0.11)
    time.sleep(.1)
    audio_controller.increase_volume(0.11)
    time.sleep(.1)
    audio_controller.increase_volume(0.11)
    time.sleep(.1)
    audio_controller.set_volume(1)
    audio_controller.unmute()


def fade_out():
    global spotify_muted
    spotify_muted = True
    # audio_controller.decrease_volume(0.3)
    # time.sleep(.2)
    # audio_controller.decrease_volume(0.11)
    # time.sleep(.1)
    # audio_controller.decrease_volume(0.11)
    # time.sleep(.1)
    # audio_controller.decrease_volume(0.11)
    # time.sleep(.1)
    # audio_controller.decrease_volume(0.11)
    # time.sleep(.1)
    audio_controller.set_volume(0)
    audio_controller.mute()


def fade_in():
    global from_paused
    global spotify_muted
    spotify_muted = None
    if not from_paused:
        audio_controller.increase_volume(0.11)
        time.sleep(.1)
        audio_controller.increase_volume(0.28)
        time.sleep(.1)
        audio_controller.increase_volume(0.33)
        time.sleep(.1)
    audio_controller.set_volume(1)
    audio_controller.unmute()


if __name__ == "__main__":
    main()
