import os
import webbrowser
from ctypes import windll
from datetime import datetime

import keyboard
import PIL.Image
import pyperclip
import pystray
import requests
from loguru import logger
from win32gui import GetForegroundWindow, GetWindowText


log_file = rf"{os.path.expanduser("~")}\ocvc.log"
logger.add(log_file, level='INFO')


@logger.catch
def run(app_name: str):
    logo = r"..\..\img\logo.png"
    app_image = PIL.Image.open(logo)
    app = pystray.Icon(
        app_name,
        app_image,
        title=app_name,
        menu=tray_menu(),
    )
    logger.info('create tray menu')
    app.run(setup)


def tray_menu() -> list[pystray.MenuItem]:
    menu_items = [pystray.MenuItem("Paste to VS Code", on_clicked),
                  pystray.MenuItem("GitHub", on_clicked),
                  pystray.MenuItem(current_version(), on_clicked)]
    if new_version_available():
        menu_items.append(pystray.MenuItem("Download new version",
                                           on_clicked))
    menu_items.append(pystray.MenuItem("Show Log", on_clicked))
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
    url = "https://api.github.com/repos/alei1180/ocvc/releases/latest"
    response = requests.get(url)
    if response.status_code == 200:
        tag = response.json()["name"]
        version = int_version_from_string(tag)
    else:
        logger.debug(f"error response: {url}, status code: {response.status_code}")
    return version


def on_clicked(app_icon: pystray.Icon, item: pystray.MenuItem):
    if str(item) == "Paste to VS Code":
        logger.info(f"menu: {item}")
        open_code_in_vsc_from_menu()
    elif str(item) == "GitHub":
        logger.info(f"menu: {item}")
        url = "https://github.com/alei1180/ocvc"
        logger.info(f"open in web browser url: {url}")
        webbrowser.open(url, new=2)
    elif str(item) == current_version():
        logger.info(f"menu: {item}")
        tag = str(item).replace(" ", "")
        url = f"https://github.com/alei1180/ocvc/releases/tag/{tag}"
        logger.info(f"open in web browser url: {url}")
        webbrowser.open(url, new=2)
    elif str(item) == "Download new version":
        logger.info(f"menu: {item}")
        url = "https://github.com/alei1180/ocvc/releases/latest/"
        logger.info(f"open in web browser url: {url}")
        webbrowser.open(url, new=2)
    elif str(item) == "Show Log":
        logger.info(f"menu: {item}")
        logger.info(f"open in vscode log file {log_file}")
        os.system(f"code {log_file} --new-window --sync off")
    elif str(item) == "Close":
        logger.info(f"menu: {item}")
        app_icon.visible = False
        app_icon.stop()
        logger.info(f"close app: {current_version()}")


def open_code_in_vsc_from_menu():
    logger.info("copy code from clipboard in var")
    copied_code = copy_clipboard().strip()
    logger.info("clear clipboard")
    empty_clipboard()
    if not copied_code:
        logger.debug("clipboard is empty")
        return None
    file_name = file_name_from_menu()
    if not file_name:
        return None

    with open(file_name, "w", encoding="utf8") as file:
        logger.info(f"write code from var in file {file_name}")
        file.write(copied_code)

    logger.info(f"open in vscode file {file_name}")
    os.system(f"code {file_name} --new-window --sync off")


def file_name_from_menu() -> str:
    path_to_save = rf"{os.getenv('TEMP')}\mrg\\".replace("\\\\", "\\")
    if not os.path.exists(path_to_save):
        os.makedirs(path_to_save)
    file_name = f"{path_to_save}ocvc_paste_from_menu_{format_current_date()}.bsl"
    return file_name


def format_current_date() -> str:
    return datetime.now().strftime('%d%m%y%H%M%S')


def setup(tray_icon: pystray.Icon):
    tray_icon.visible = True
    if tray_icon.visible:
        wait_press_hotkey(hotkey())


def hotkey() -> str:
    return "Shift + Alt + V"


def wait_press_hotkey(hotkey: str):
    keyboard.add_hotkey(hotkey, lambda: open_code_in_vsc_from_hotkey())


def open_code_in_vsc_from_hotkey():
    current_window = GetWindowText(GetForegroundWindow())
    if not configurator_window(current_window, "Конфигуратор"):
        return None
    logger.info(f"press hotkey {hotkey()} in window Конфигуратор")

    logger.info("copy code from clipboard in var")
    copied_code = copy_clipboard().strip()
    logger.info("clear clipboard")
    empty_clipboard()
    if not copied_code:
        logger.info("clipboard is empty")
        return None

    file_name = file_name_from_window(current_window)
    if not file_name:
        logger.debug("file name is empty")
        return None

    with open(file_name, "w", encoding="utf8") as file:
        logger.info(f"write code from var in file {file_name}")
        file.write(copied_code)

    logger.info(f"open in vscode file {file_name}")
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

    path_to_save = rf"{os.getenv('TEMP')}\mrg\\".replace("\\\\", "\\")
    if not os.path.exists(path_to_save):
        os.makedirs(path_to_save)

    cut_interval = window_name_cut_interval(window_name)
    file_name = window_name[cut_interval["begin"] : cut_interval["end"]]
    file_name = file_name.replace("\\", "")
    file_name = file_name.replace(":", "")
    file_name = file_name.replace("[Только для чтения]", "")
    file_name = file_name.strip()
    file_name = file_name.replace(" ", "_")
    add_postfix_module = file_name.upper().find("МОДУЛЬ") == -1
    if add_postfix_module:
        file_name = f"{file_name}_Модуль"
    file_name = f"{path_to_save}ocvc_{file_name}_{format_current_date()}.bsl"
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
    logger.info(f"start app: {current_version()}")
    run("ocvc")
