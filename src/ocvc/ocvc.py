import os
import webbrowser
from ctypes import windll

import keyboard
import PIL.Image
import pyperclip
import pystray
import requests
from win32gui import GetForegroundWindow, GetWindowText


def run(app_name: str):
    logo = r"..\..\img\logo.png"
    app_image = PIL.Image.open(logo)
    app = pystray.Icon(
        app_name,
        app_image,
        title=app_name,
        menu=tray_menu(),
    )

    app.run(setup)


def tray_menu() -> list[pystray.MenuItem]:
    menu_items = [pystray.MenuItem("GitHub", on_clicked),
                  pystray.MenuItem(current_version(), on_clicked)]
    if new_version_available():
        menu_items.append(pystray.MenuItem("download new version",
                                           on_clicked))
    menu_items.append(pystray.MenuItem("Close", on_clicked))
    return menu_items


def new_version_available() -> bool:
    cur_version = int_version_from_string(current_version())
    last_version = github_ocvc_last_version_number()
    return last_version > cur_version


def int_version_from_string(string_version: str) -> int:
    return int(string_version.replace("ver ", "").replace(".", ""))


def current_version() -> str:
    return "ver 1.3"


def github_ocvc_last_version_number() -> int:
    version = 0
    response = requests.get("https://api.github.com/repos/alei1180/ocvc/releases/latest")
    if response.status_code == 200:
        tag = response.json()["name"]
        version = int_version_from_string(tag)
    return version


def on_clicked(app_icon: pystray.Icon, item: pystray.MenuItem):
    if str(item) == "GitHub":
        webbrowser.open("https://github.com/alei1180/ocvc", new=2)
    elif str(item) == current_version():
        tag = str(item).replace(" ", "")
        webbrowser.open(f"https://github.com/alei1180/ocvc/releases/tag/{tag}", new=2)
        print(github_ocvc_last_version_number())
    elif str(item) == "download new version":
        webbrowser.open("https://github.com/alei1180/ocvc/releases/latest/", new=2)
        print(github_ocvc_last_version_number())
    elif str(item) == "Close":
        app_icon.visible = False
        app_icon.stop()


def setup(tray_icon: pystray.Icon):
    tray_icon.visible = True
    if tray_icon.visible:
        wait_press_hotkey("Shift + Alt + V")


def wait_press_hotkey(hotkey: str):
    keyboard.add_hotkey(hotkey, lambda: open_code_in_vsc())


def open_code_in_vsc():
    current_window = GetWindowText(GetForegroundWindow())
    if not configurator_window(current_window, "Конфигуратор"):
        return None

    copied_code = copy_clipboard()
    empty_clipboard()
    if not copied_code:
        return None

    file_name = file_name_from_window(current_window)
    if not file_name:
        return None

    with open(file_name, "w", encoding="utf8") as file:
        file.write(copied_code)

    os.system(f"code {file_name} --new-window --sync off")


def configurator_window(window: str, title: str) -> bool:
    return window.find(title) != -1


def copy_clipboard() -> str:
    copied_code = pyperclip.paste()
    copied_code = copied_code.replace("\r", "")
    return copied_code


def empty_clipboard():
    if windll.user32.OpenClipboard(None):
        windll.user32.EmptyClipboard()
        windll.user32.CloseClipboard()


def file_name_from_window(window_name: str) -> str:
    file_name = ""
    if not window_name:
        return file_name

    path_to_save = rf"{os.getenv('TEMP')}\mrg\\"
    if not os.path.exists(path_to_save):
        os.makedirs(path_to_save)

    cut_interval = window_name_cut_interval(window_name)
    file_name = window_name[cut_interval["begin"] : cut_interval["end"]]
    file_name = file_name.replace(":", "")
    file_name = file_name.replace("[Только для чтения]", "")
    file_name = file_name.strip()
    file_name = file_name.replace(" ", "_")
    add_postfix_module = file_name.upper().find("МОДУЛЬ") == -1
    if add_postfix_module:
        file_name = f"{file_name}_Модуль"
    file_name = f"{file_name}.bsl"
    file_name = f"{path_to_save}{file_name}"
    return file_name


def window_name_cut_interval(window_name: str) -> dict[str, int]:
    cut = {"begin": 0, "end": 0}

    external_file = (
        window_name.find(".epf") != -1 or window_name.find(".erf") != -1
    ) and window_name.find("\\")
    if external_file:
        cut["begin"] = window_name.rfind("\\")
        cut["end"] = window_name.rfind(" - Конфигуратор")
        return cut

    cut["begin"] = window_name.find("= ") + 1
    if window_name.find("Общий модуль") != -1:
        cut["end"] = window_name.find(":")
    else:
        cut["end"] = window_name.find(" -")
    return cut


if __name__ == "__main__":
    run("ocvc")
