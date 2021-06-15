import time  # sleep
from pycaw.pycaw import AudioUtilities  # mute

import os
import win32con as wcon
import win32api as wapi
import win32gui as wgui
import win32process as wproc


# New Curses GUI
import curses
import traceback
import datetime

# Constant
VERSION = 'b0.6.18'
WINDOW_TITLE = 'Spotify Ads Ratio Counter'
WINDOW_AUTHOR = ' by engrbugs'

SPOTIFY_APP_EXE = 'Spotify.exe'
SLEEP_TIME_SONG = 1  # sec
SLEEP_TIME_ADVERTISEMENT = 1  # sec
PAUSED_TEXT = 'Spotify Free'
START_COUNT_DOWN = 3

# Variable
spotify_muted = None
from_paused = True
volume_now = 1.0


class ProgressAnimation:
    counter = 0
    ANIM_LENGTH = 11
    X_COORD = 0
    Y_COORD = 0

    # CHAR
    OPENING_CHAR = '['
    CLOSING_CHAR = ']'
    FULL_UNICODE_CHAR = '█'
    HALF_UNICODE_CHAR = '▒'

    # COLORS
    COLOR_GREEN = 0xA
    COLOR_BLUE = 0xB
    COLOR_YELLOW = 0xE
    COLOR_WHITE = 0x7
    COLOR_GRAY = 0x8
    COLOR_FULLWHITE = 0xF

    global open_bracket_progress_win
    global close_bracket_progress_win
    global body_progress_win

    def __init__(self, xcoord, ycoord):
        # self.X_COORD = xcoord
        # self.Y_COORD = ycoord
        curses.init_pair(1, self.COLOR_GRAY, curses.COLOR_BLACK)  # darkest non-black color on black bg
        curses.init_pair(2, self.COLOR_FULLWHITE, curses.COLOR_BLACK)  # full white on black bg
        curses.init_pair(3, self.COLOR_WHITE, self.COLOR_GRAY)  # full white on black bg
        curses.init_pair(4, self.COLOR_GREEN, curses.COLOR_BLACK)  # green on black bg
        curses.init_pair(5, self.COLOR_BLUE, curses.COLOR_BLUE)  # blue on black bg, used to highlight max
        curses.init_pair(6, self.COLOR_YELLOW, curses.COLOR_BLACK)
        curses.init_pair(7, curses.COLOR_RED, curses.COLOR_BLACK)
        self.open_bracket_progress_win = curses.newwin(2, len(self.OPENING_CHAR), ycoord, xcoord)
        self.body_progress_win = curses.newwin(2, self.ANIM_LENGTH, ycoord, xcoord + len(self.OPENING_CHAR))
        self.close_bracket_progress_win = curses.newwin(2, len(self.CLOSING_CHAR), ycoord, xcoord
                                                        + self.ANIM_LENGTH + len(self.OPENING_CHAR))
        self.open_bracket_progress_win.addstr('[')
        self.body_progress_win.addstr(self.FULL_UNICODE_CHAR * self.ANIM_LENGTH, curses.color_pair(1))
        self.close_bracket_progress_win.addstr(']')

    def refresh(self):
        self.open_bracket_progress_win.touchwin()
        self.open_bracket_progress_win.refresh()
        self.close_bracket_progress_win.touchwin()
        self.close_bracket_progress_win.refresh()
        self.body_progress_win.touchwin()
        self.body_progress_win.refresh()

    def next(self):
        # Do not forget to call refresh after you call this func.
        self.body_progress_win.clear()

        self.counter += 1
        body_str = (self.FULL_UNICODE_CHAR * (abs(self.counter) - 1)) + self.HALF_UNICODE_CHAR + \
                   (self.FULL_UNICODE_CHAR * (self.ANIM_LENGTH - abs(self.counter)))
        print(body_str)
        self.body_progress_win.addstr(body_str, curses.color_pair(1))

        # Paint color
        self.body_progress_win.chgat(0, abs(self.counter) - 1, 1, curses.color_pair(5))

        if self.counter == self.ANIM_LENGTH:
            self.counter *= -1
        elif self.counter == -2:
            self.counter = 0


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
    proc_name = args[0] if args else None
    procs = enum_processes(process_name=proc_name)
    for pid, name in procs:
        data = enum_process_windows(pid)
        if data:
            proc_text = "PId {0:d}{1:s}windows:".format(pid, " (File: [{0:s}]) ".format(name) if name else " ")
            # print(proc_text)
            for handle, text in data:
                print("{0:d}: [{1:s}]".format(handle, text))
                return text
    return ''


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
                global volume_now
                volume_now = 0.0
                print(self.process_name, 'has been muted.')  # debug

    def unmute(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                interface.SetMute(0, None)
                global volume_now
                volume_now = 1.0
                print(self.process_name, 'has been unmuted.')  # debug

    def process_volume(self):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                global volume_now
                volume_now = interface.GetMasterVolume()
                return interface.GetMasterVolume()

    def set_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                # only set volume in the range 0.0 to 1.0
                self.volume = min(1.0, max(0.0, decibels))
                interface.SetMasterVolume(self.volume, None)
                global volume_now
                volume_now = self.volume

    def decrease_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                try:
                    # 0.0 is the min value, reduce by decibels
                    self.volume = max(0.0, self.volume - decibels)
                    interface.SetMasterVolume(self.volume, None)
                    global volume_now
                    volume_now = self.volume
                except:
                    # 0.0 is the min value, reduce by decibels
                    self.volume = .7
                    interface.SetMasterVolume(self.volume, None)
                    # global volume_now
                    volume_now = self.volume

    def increase_volume(self, decibels):
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            interface = session.SimpleAudioVolume
            if session.Process and session.Process.name() == self.process_name:
                try:
                    # 1.0 is the max value, raise by decibels
                    self.volume = min(1.0, self.volume + decibels)
                    interface.SetMasterVolume(self.volume, None)
                    global volume_now
                    volume_now = self.volume
                except:
                    # 1.0 is the max value, raise by decibels
                    self.volume = 1
                    interface.SetMasterVolume(self.volume, None)
                    # global volume_now
                    volume_now = self.volume


def main():
    global from_paused
    global audio_controller
    global spotify_muted
    global SPOTIFY_APP_EXE

    audio_controller = AudioController(SPOTIFY_APP_EXE)
    # for i in range(start_count_down):
    #     print('Spotify Ad Muter will initialize in ' + str(start_count_down - i) + '...')
    #     time.sleep(1)

    app_initialize()

    # New Curses GUI
    stdscr = curses.initscr()  # initialize curses screen
    height, width = stdscr.getmaxyx()
    curses.start_color()
    curses.noecho()  # turn off auto echoing of keypress on to screen
    curses.cbreak()  # enter break mode where pressing Enter key
    curses.curs_set(0)
    #  after keystroke is not required for it to register
    stdscr.keypad(True)  # enable special Key values such as curses.KEY_LEFT etc

    # Paint the screen
    progress_animation = ProgressAnimation(5, 2)
    stdscr.border(0)

    tempy = 2  # row of Main Title
    tempx = int((width - len(WINDOW_TITLE) - len(WINDOW_AUTHOR)) / 2)
    stdscr.addstr(tempy, tempx, WINDOW_TITLE, curses.A_BOLD)
    stdscr.addstr(tempy, tempx + len(WINDOW_TITLE), WINDOW_AUTHOR, curses.A_NORMAL)

    tempy = 1  # row of version
    tempx = width - len(VERSION) - 1
    stdscr.addstr(tempy, tempx, VERSION)

    tempy = 4  # row of Spotify Window Text:
    tempx = 5
    stdscr.addstr(tempy, tempx, 'SPOTIFY window text:', curses.A_BOLD)

    tempy = 5  # row of Spotify Window Text:
    tempx = 10
    textbox_current_window_text = curses.newwin(2, width - (2 * tempx), tempy, tempx)
    #stdscr.addstr(tempy, tempx, 'SPOTIFY window text:', curses.A_BOLD)

    tempy = 7  # row of Spotify Window Text:
    tempx = 5
    stdscr.addstr(tempy, tempx, 'Volume:')

    tempy = 7  # row of Volume Textbox:
    tempx = 13
    textbox_volume = curses.newwin(2, 3, tempy, tempx)

    while True:
        time.sleep(SLEEP_TIME_SONG)

        title = check_window_text(SPOTIFY_APP_EXE)  # New

        #  if "Advertisement" in titles:  # Spotify window text is named as Advertisement
        if "Advertisement" in title:  # Spotify window text is named as Advertisement
            print('got Advertisement')
            print(audio_controller.process_volume())
            from_paused = None
            if not spotify_muted:
                fade_out()
        elif PAUSED_TEXT in title:
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
        # New Curses GUI
        textbox_current_window_text.clear()
        if title == '':
            title = SPOTIFY_APP_EXE + ' is not found.'
        textbox_current_window_text.addstr(title, curses.A_NORMAL)
        textbox_current_window_text.touchwin()
        textbox_current_window_text.refresh()
        textbox_volume.clear()
        textbox_volume.addstr(str('%.1f' % volume_now), curses.A_NORMAL if volume_now == 1 else curses.color_pair(7))
        textbox_volume.touchwin()
        textbox_volume.refresh()
        stdscr.refresh()
        progress_animation.next()
        progress_animation.refresh()


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
    audio_controller.__init__(SPOTIFY_APP_EXE)
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
