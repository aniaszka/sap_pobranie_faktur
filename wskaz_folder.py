import win32gui
import win32con
import pyautogui
import time
import logging.config
import logging
import configparser


def get_windows():
    def sort_windows(windows):
        sorted_windows = []

        # Find the first entry
        for window in windows:
            if window["hwnd_above"] == 0:
                sorted_windows.append(window)
                break
        else:
            raise (IndexError("Could not find first entry"))

        # Follow the trail
        while True:
            for window in windows:
                if sorted_windows[-1]["hwnd"] == window["hwnd_above"]:
                    sorted_windows.append(window)
                    break
            else:
                break

        # Remove hwnd_above
        for window in windows:
            del (window["hwnd_above"])

        return sorted_windows

    def enum_handler(hwnd, results):
        window_placement = win32gui.GetWindowPlacement(hwnd)
        results.append({
            "hwnd": hwnd,
            "hwnd_above": win32gui.GetWindow(hwnd, win32con.GW_HWNDPREV),  # Window handle to above window
            "title": win32gui.GetWindowText(hwnd),
            "visible": win32gui.IsWindowVisible(hwnd) == 1,
            "minimized": window_placement[1] == win32con.SW_SHOWMINIMIZED,
            "maximized": window_placement[1] == win32con.SW_SHOWMAXIMIZED,
            "rectangle": win32gui.GetWindowRect(hwnd)  # (left, top, right, bottom)
        })

    enumerated_windows = []
    win32gui.EnumWindows(enum_handler, enumerated_windows)
    return sort_windows(enumerated_windows)


def wklej_sciezke(faktury_sap_path):
    time.sleep(1)
    pyautogui.hotkey('shift', '\t')
    time.sleep(1)
    pyautogui.hotkey('shift', '\t')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.typewrite(faktury_sap_path)
    time.sleep(1)
    pyautogui.press('\t')
    time.sleep(1)
    pyautogui.press('\t')
    time.sleep(1)
    pyautogui.press('\n')
    time.sleep(1)



def wstaw_sciezke():
    logging_settings_file = 'logging.ini'
    logging.config.fileConfig(logging_settings_file, disable_existing_loggers=False)
    logger = logging.getLogger(__name__)

    while True:
        windows = get_windows()
        logger.debug('Tworzę listę otwartych okien windowsa.')
        try:
            moje_okno = next(
                (item for item in windows if item["title"] == "Przeglądanie w poszukiwaniu plików lub folderów"), False)

            #logger.debug('numer okna: ', moje_okno['hwnd'])

            time.sleep(2)
            win32gui.SetForegroundWindow(moje_okno['hwnd'])  # aktywuje okno
            time.sleep(2)
            win32gui.ShowWindow(moje_okno['hwnd'], win32con.SW_SHOWNORMAL)  # wyciąga na wierzch
            time.sleep(2)

            config = configparser.ConfigParser(interpolation=None)
            config.read('sap_pobranie_faktur_settings.ini')

            faktury_sap_path = config['server_path']['faktury_sap_path']  # string

            # TODO wywalić 3 poniższe wiersze
            print('Sprawdzam co się dzieje')
            print(faktury_sap_path)


            wklej_sciezke(faktury_sap_path)
            logger.debug('Folder wskazany!')
            break
        except:
            logger.debug('Nie ma okna do wskazania ścieżki. Próbuję dalej.')
            time.sleep(2)

if __name__ == "__main__":
    wstaw_sciezke()
