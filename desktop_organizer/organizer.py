import os
import subprocess
import json
from time import sleep

import win32gui
import win32api


def string_to_tuple(string: str, dtype=int) -> tuple:
    """Converts a suitable decoded string into an encoded tuple.

    :param string: Decoded tuple as a string
    :type string: str
    :param dtype: data type of tuple values, defaults to int
    :type dtype: data type, optional
    :return: Tuple representation of string
    :rtype: tuple
    """
    if not isinstance(string, str):
        raise TypeError("Source string hast to be of type str!")

    if dtype == str:
        string = "".join(string.split())
    values = [dtype(v) for v in string.strip(" ()[]").split(",") if v != ""]
    return tuple(values)


def get_screen_size():
    width = win32api.GetSystemMetrics(0)
    height = win32api.GetSystemMetrics(1)
    return width, height


def calc_pixel(value_tuple: tuple, dim: int) -> int:
    if value_tuple[dim] < 1:
        value = get_screen_size()[dim] * value_tuple[dim]
    else:
        value = value_tuple[dim]
    return int(value)


def list_all_windows():
    def _print_window(hwnd, _):
        rect = win32gui.GetWindowRect(hwnd)
        x = rect[0]
        y = rect[1]
        w = rect[2] - x
        h = rect[3] - y
        name = win32gui.GetWindowText(hwnd)
        print(f"Window '{name}' @ X:{x}, Y:{y}, W:{w}, H:{h}")

    win32gui.EnumWindows(_print_window, None)


class Program:
    def __init__(self, config) -> None:
        self.window_id = 0
        self.name = config["name"]
        self.location = config["location"]
        self.size = string_to_tuple(config["size"], float)
        self.pos = string_to_tuple(config["position"], float)
        self.delay = getattr(config, "delay", 0)

    def start(self) -> None:
        print(f"[INIT]: Start {self.name} from {self.location}")
        subprocess.call(self.location)
        sleep(self.delay)
        print("[ OK ]: Done")

    def get_or_create(self) -> None:
        if not self.find_window_id():
            self.start()
            for _ in range(RETRY):
                if self.find_window_id():
                    break
                sleep(1)

    def find_window_id(self) -> bool:
        def _check_name(hwnd, _):
            text = win32gui.GetWindowText(hwnd)
            if text is not None and self.name in text:
                self.window_id = hwnd

        win32gui.EnumWindows(_check_name, None)

        if self.window_id > 0:
            print(f"[ OK ]: Found window '{self.name}' with ID: {self.window_id}")
            return True
        print(f"[FAIL]: Could not find window with name '{self.name}'")
        return False

    def resize(self) -> None:
        x = calc_pixel(self.pos, 0)
        y = calc_pixel(self.pos, 1)
        w = calc_pixel(self.size, 0)
        h = calc_pixel(self.size, 1)
        print(f"[INFO]: Set {self.name} to X:{x}, Y:{y}, W:{w}, H:{h}")
        win32gui.MoveWindow(self.window_id, x, y, w, h, True)


def main():
    folder = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
    with open(folder + r"\organizer.json", encoding="utf-8") as f:
        settings = json.load(f)

    RETRY = settings["retry"]
    DEBUG = settings["debug"]
    GET_LIST = settings["get_list"]

    if GET_LIST:
        list_all_windows()

    for program_config in settings["programs"]:
        p = Program(program_config)
        p.get_or_create()
        p.resize()


if __name__ == "__main__":
    main()
