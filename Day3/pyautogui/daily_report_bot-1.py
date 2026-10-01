
import pyautogui
import pyperclip
import time
import os
from datetime import datetime


# ============================================================
# DAILY REPORT BOT
# Chrome + PyAutoGUI + Microsoft Excel
# ============================================================

pyautogui.PAUSE = 0.7

# ------------------------------------------------------------
# 1. Generate date and time automatically
# ------------------------------------------------------------

now = datetime.now()

current_date = now.strftime("%Y-%m-%d")
current_datetime = now.strftime("%Y-%m-%d %H:%M:%S")

script_folder = os.path.dirname(os.path.abspath(__file__))

excel_filename = f"daily_report_{current_date}.xlsx"
screenshot_filename = f"daily_report_{current_date}.png"

excel_path = os.path.join(script_folder, excel_filename)
screenshot_path = os.path.join(script_folder, screenshot_filename)


# ------------------------------------------------------------
# 2. Open Chrome
# ------------------------------------------------------------

print("1. Opening Chrome...")

pyautogui.hotkey("win", "r")
time.sleep(1)

pyautogui.write("chrome")
pyautogui.press("enter")

time.sleep(5)


# ------------------------------------------------------------
# 3. Open public weather website
# ------------------------------------------------------------

print("2. Opening weather website...")

pyautogui.hotkey("ctrl", "l")

pyautogui.write(
    "https://wttr.in/Gummidipoondi?format=3",
    interval=0.01
)

pyautogui.press("enter")

time.sleep(7)


# ------------------------------------------------------------
# 4. Copy weather information
# ------------------------------------------------------------

print("3. Copying weather information...")

pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("ctrl", "c")

time.sleep(1)

weather_data = pyperclip.paste().strip()

if not weather_data:
    weather_data = "Weather information unavailable"

print("Weather data:", weather_data)

# ------------------------------------------------------------
# 5. Close Chrome
# ------------------------------------------------------------

print("4. Closing Chrome...")
pyautogui.hotkey("alt", "f4")
time.sleep(2)

# ------------------------------------------------------------
# 5. Create automatic comment
# ------------------------------------------------------------

comment = "Daily weather status recorded successfully."


# ------------------------------------------------------------
# 6. Open Microsoft Excel
# ------------------------------------------------------------

print("4. Opening Microsoft Excel...")

pyautogui.hotkey("win", "r")
time.sleep(1)

pyautogui.write("excel")
pyautogui.press("enter")

time.sleep(8)


# ------------------------------------------------------------
# 7. Create a new Excel workbook
# ------------------------------------------------------------

print("5. Creating new Excel workbook...")

# If Excel opens with the start screen,
# create a new blank workbook.
pyautogui.hotkey("ctrl", "n")

time.sleep(4)


# ------------------------------------------------------------
# 8. Enter report headers and data
# ------------------------------------------------------------

print("6. Entering report data...")

# Header row
pyautogui.write("Date & Time")
pyautogui.press("tab")

pyautogui.write("Fetched Data")
pyautogui.press("tab")

pyautogui.write("Comment")
pyautogui.press("enter")


# Data row
pyautogui.write(current_datetime)
pyautogui.press("tab")

pyautogui.write(weather_data)
pyautogui.press("tab")

pyautogui.write(comment)

time.sleep(2)


# ------------------------------------------------------------
# 9. Format columns
# ------------------------------------------------------------

print("7. Formatting Excel sheet...")

# Select the used range
pyautogui.hotkey("ctrl", "a")

# Auto-fit columns
pyautogui.hotkey("alt", "h")
time.sleep(1)

pyautogui.press("o")
time.sleep(1)

pyautogui.press("i")

time.sleep(2)


# ------------------------------------------------------------
# 10. Save Excel workbook
# ------------------------------------------------------------

print("8. Saving Excel workbook...")

pyautogui.hotkey("ctrl", "shift", "s")

time.sleep(3)


# ------------------------------------------------------------
# 11. Enter filename
# ------------------------------------------------------------

print("9. Entering filename...")

# Ctrl + A selects the filename field in the Save As dialog
pyautogui.hotkey("ctrl", "a")

pyautogui.write(excel_path, interval=0.01)

time.sleep(1)

pyautogui.press("enter")

time.sleep(5)


# ------------------------------------------------------------
# 12. Handle possible overwrite confirmation
# ------------------------------------------------------------

# If the file already exists, Excel may display
# a confirmation dialog.

pyautogui.press("left")
pyautogui.press("enter")

time.sleep(4)


# ------------------------------------------------------------
# 13. Handle possible Excel format confirmation
# ------------------------------------------------------------

# Excel may ask whether to keep the current format.

pyautogui.press("enter")

time.sleep(3)


# ------------------------------------------------------------
# 14. Take screenshot
# ------------------------------------------------------------

print("10. Taking screenshot of final Excel sheet...")

# Make sure Excel is active
pyautogui.hotkey("alt", "tab")

time.sleep(2)

# Move to top-left of worksheet
pyautogui.hotkey("ctrl", "home")

time.sleep(2)

# Take screenshot of entire screen
screenshot = pyautogui.screenshot()

screenshot.save(screenshot_path)


# ------------------------------------------------------------
# 15. Final result
# ------------------------------------------------------------

print()
print("=" * 60)
print("DAILY REPORT BOT COMPLETED")
print("=" * 60)

print("Date & Time :", current_datetime)
print("Weather     :", weather_data)
print("Excel file  :", excel_path)
print("Screenshot  :", screenshot_path)

print()
print("Files created successfully.")

