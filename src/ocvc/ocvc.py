import os
from ctypes import windll

import keyboard
import PIL.Image
import pyperclip
import pystray
from win32gui import GetForegroundWindow, GetWindowText


def run(app_name: str) -> None:
    logo = r"..\..\img\logo.png"
    app_image = PIL.Image.open(logo)
    app = pystray.Icon(
        app_name,
        app_image,
        title=app_name,
        menu=pystray.Menu(pystray.MenuItem("Close", on_clicked)),
    )

    app.run(setup)


def on_clicked(app_icon: pystray.Icon, item: pystray.MenuItem) -> None:
    if str(item) == "Close":
        app_icon.visible = False
        app_icon.stop()


def setup(tray_icon: pystray.Icon) -> None:
    tray_icon.visible = True
    if tray_icon.visible:
        wait_press_hotkey("Shift + Alt + V")


def wait_press_hotkey(hotkey: str) -> None:
    keyboard.add_hotkey(hotkey, lambda: open_code_in_vsc())


def open_code_in_vsc() -> None:
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


def empty_clipboard() -> None:
    if windll.user32.OpenClipboard(None):
        windll.user32.EmptyClipboard()
        windll.user32.CloseClipboard()


def file_name_from_window(window_name: str) -> str:
    file_name = ""
    if not window_name:
        return file_name

    path_to_save = rf"{os.getenv('TEMP')}\mrg\\"

    cut_interval = window_name_cut_interval(window_name)
    file_name = window_name[cut_interval["begin"]:cut_interval["end"]]
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
