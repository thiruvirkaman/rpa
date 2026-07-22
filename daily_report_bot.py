from pdb import main
import re
import sys
import time
from datetime import datetime
from pathlib import Path
import pyautogui
import pyperclip

#Disclaimer this code is for windows only and requires Chrome and Excel to be installed. It may not work on other platforms or with different browsers or spreadsheet applications.

CITY = "Bengaluru"
WEATHER_URL = f"https://wttr.in/{CITY}?format=3"
OUTPUT_DIR = Path(__file__).resolve().parent / "reports"

APP_START_SECONDS = 6
PAGE_LOAD_SECONDS = 8

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1.0


def wait(seconds: float, message: str) -> None:
    print(message)
    time.sleep(seconds)

#Preventing overwriting existing files by generating a unique file path if the specified path already exists
def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_stem(f"{path.stem}_{counter}")
        if not candidate.exists():
            return candidate
        counter += 1

#Function to open an application using the Windows Run dialog
def open_with_run(command: str) -> None:
    pyautogui.hotkey("win", "r")
    time.sleep(0.8)
    pyautogui.write(command, interval=0.04)
    pyautogui.press("enter")

#Function to fetch the weather from the Chrome browser and return it as a string
def fetch_weather_from_chrome() -> str:
    open_with_run("chrome")
    wait(APP_START_SECONDS, "1/5 Chrome opened")
    pyautogui.hotkey("ctrl", "l")
    pyautogui.write(WEATHER_URL, interval=0.02)
    pyautogui.press("enter")
    wait(PAGE_LOAD_SECONDS, f"2/5 Loaded weather for {CITY}")
    width, height = pyautogui.size()
    pyautogui.moveTo(width // 2, height // 2, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey("ctrl", "a")
    pyautogui.hotkey("ctrl", "c")
    time.sleep(1)
    weather = " ".join(pyperclip.paste().split())
    if not weather or CITY.lower() not in weather.lower():
        raise RuntimeError(
            "The weather value was not copied. Check the internet connection, "
            "close any Chrome pop-ups, and run the bot again."
        )
    return weather


def weather_comment(weather: str) -> str:
    match = re.search(r"(-?\d+)\s*°?C", weather, flags=re.IGNORECASE)
    if not match:
        return "Weather update recorded"

    temperature = int(match.group(1))
    if temperature >= 32:
        return "Hot day - stay hydrated"
    if temperature >= 24:
        return "Good for outdoor activities"
    if temperature >= 18:
        return "Cool weather - carry a jacket"
    return "Cold weather - dress warmly"


def create_excel_report(weather: str, report_path: Path) -> None:
    open_with_run("excel")
    wait(APP_START_SECONDS, "3/5 Excel opened")
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.hotkey("win", "up")
    time.sleep(1)
    pyautogui.hotkey("ctrl", "home")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [
        ["Date & Time", "Fetched Weather", "Comment"],
        [now, weather, weather_comment(weather)],
    ]

    table = "\r\n".join("\t".join(row) for row in rows)
    pyperclip.copy(table)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.hotkey("alt", "h")
    pyautogui.press("o")
    pyautogui.press("i")
    time.sleep(1)
    pyautogui.hotkey("ctrl", "home")
    pyautogui.press("f12")
    time.sleep(2)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(str(report_path), interval=0.01)
    pyautogui.press("enter")
    wait(3, "4/5 Workbook saved")
    pyautogui.hotkey("alt", "f4")


OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
date_stamp = datetime.now().strftime("%Y-%m-%d")
report_path = unique_path(OUTPUT_DIR / f"daily_report_{date_stamp}.xlsx")
screenshot_path = unique_path(OUTPUT_DIR / f"daily_report_{date_stamp}.png")
print("Starting in 2 seconds.")
time.sleep(2)

try:
    weather = fetch_weather_from_chrome()
    print(f"Copied: {weather}")
    create_excel_report(weather, report_path)
    time.sleep(2)
    pyautogui.screenshot(str(screenshot_path))
except pyautogui.FailSafeException:
    print("Bot stopped by the PyAutoGUI fail-safe.")
except Exception as exc:
        print(f"Bot failed: {exc}")
print("5/5 Screenshot saved")
print(f"Excel report: {report_path}")
print(f"Screenshot:   {screenshot_path}")

