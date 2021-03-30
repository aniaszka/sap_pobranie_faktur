import win32gui
import win32con
import logging.config
import logging


def get_windows():
    logging_settings_file = 'logging.ini'
    logging.config.fileConfig(logging_settings_file, disable_existing_loggers=False)
    logger = logging.getLogger(__name__)
    logger.debug('Funkcja get_windows uruchomiona')
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

def close_window_by_part_name(tekst):
    logging.config.fileConfig('logging.ini', disable_existing_loggers=False)
    logger = logging.getLogger(__name__)
    logger.debug('Funkca close_window uruchomiona.')
    windows = get_windows()
    hwnd_to_close = []
    logger.debug('Tworzę listę okien do zamknięcia.')
    for window in windows:
        if tekst in window['title']:  # 'sap_attachments'
            print(window)
            flaga = True
            hwnd_to_close.append(window['hwnd'])
        else:
            flaga = False
    logger.debug(hwnd_to_close)
    for handle in hwnd_to_close:
        win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
    logger.debug('Okno lub okna zamknięte.')
    return len(hwnd_to_close) != 0



if __name__ == '__main__':
    tekst = 'Excel'
    x = close_window_by_part_name(tekst)
    print(x)