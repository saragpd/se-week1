import pyautogui
import pyperclip
import time
import os 
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
# ============================================================ 
# DAILY STATUS REPORT BOT 
# PyAutoGUI + Chrome + Microsoft Excel 
# =========================================================

# ------------------------------------------------------------ 
# 1. Basic configuration 
# ------------------------------------------------------------

# Give applications enough time to open/load 
pyautogui.PAUSE = 0.8
# Get current date and time automatically
now = datetime.now()

current_date = now.strftime("%Y-%m-%d")
current_datetime = now.strftime("%Y-%m-%d %H:%M:%S")

# File names 
excel_filename = f"daily_report_{current_date}.xlsx"
screenshot_filename = f"daily_report_{current_date}.png"

# Save files in the same folder as this Python script
script_folder = os.path.dirname(os.path.abspath(__file__))

excel_path = os.path.join(script_folder, excel_filename)
screenshot_path = os.path.join(script_folder, screenshot_filename)

# ------------------------------------------------------------ 
# 2. Helper function 
# ------------------------------------------------------------

def wait(seconds=2): 
	"""Wait for applications/web pages to respond.""" 
	time.sleep(seconds)
	
# ------------------------------------------------------------ 
# 3. Open Chrome 
# ------------------------------------------------------------

print("Step 1: Opening Chrome...")

pyautogui.hotkey("win", "r") 
wait(1)

pyautogui.write("chrome")
pyautogui.press("enter") 
wait(5)

# ------------------------------------------------------------ 
# 4. Open a public weather website 
# ------------------------------------------------------------

print("Step 2: Opening weather website...")
pyautogui.hotkey("ctrl", "l")
# wttr.in provides simple text-based weather information 
pyautogui.write("https://wttr.in/Gummidipoondi?format=3")

pyautogui.press("enter")
wait(6)

# ------------------------------------------------------------ 
# 5. Copy weather information from the browser 
# ------------------------------------------------------------
print("Step 3: Copying weather information...")

# Select everything visible on the webpage 
pyautogui.hotkey("ctrl", "a")

# Copy selected information 
pyautogui.hotkey("ctrl", "c")

wait(1)

# Read clipboard 
weather_data = pyperclip.paste().strip()

# If clipboard is empty, use a fallback message 
if not weather_data:
	weather_data = "Weather data could not be copied"
	
print("Fetched data:", weather_data)

# ------------------------------------------------------------ 
# 6. Generate a short comment automatically 
# ------------------------------------------------------------ 
print("Step 4: Creating report comment...") 
comment = "Daily weather status recorded successfully."

# ------------------------------------------------------------ 
# 7. Create Excel workbook 
# ------------------------------------------------------------

print("Step 5: Creating Excel workbook...")

# Create a new workbook 
workbook = Workbook()

# Select active worksheet 
worksheet = workbook.active

# Rename worksheet 
worksheet.title = "Daily Status"

# ------------------------------------------------------------ 
# 8. Add headers 
# ------------------------------------------------------------

worksheet["A1"] = "Date & Time" 
worksheet["B1"] = "Fetched Data" 
worksheet["C1"] = "Comment"

# Make headers bold 
for cell in worksheet[1]: 
	cell.font = Font(bold=True) 
	cell.alignment = Alignment(horizontal="center")

# ------------------------------------------------------------ 
# 9. Add today's data 
# ------------------------------------------------------------

worksheet["A2"] = current_datetime 
worksheet["B2"] = weather_data 
worksheet["C2"] = comment

# ------------------------------------------------------------ 
# 10. Format the worksheet 
# ------------------------------------------------------------

worksheet.column_dimensions["A"].width = 22 
worksheet.column_dimensions["B"].width = 45 
worksheet.column_dimensions["C"].width = 45

worksheet["A2"].alignment = Alignment(horizontal="center")
worksheet["B2"].alignment = Alignment(wrap_text=True)
worksheet["C2"].alignment = Alignment(wrap_text=True)

# ------------------------------------------------------------ 
# 11. Save Excel file 
# ------------------------------------------------------------

print("Step 6: Saving Excel file...")
workbook.save(excel_path)
print("Excel file saved:")

print(excel_path)

# ------------------------------------------------------------ 
#12. Open the generated Excel file 
# ------------------------------------------------------------ 

print("Step 7: Opening Excel file...")
# Use Windows Run dialog to open the Excel file 
pyautogui.hotkey("win", "r") 
wait(1)

pyautogui.write(f'"{excel_path}"')
pyautogui.press("enter")
# Give Excel enough time to open
wait(8)

# ------------------------------------------------------------ 
# 9. Add today's data 
# ------------------------------------------------------------

worksheet["A2"] = current_datetime 
worksheet["B2"] = weather_data 
worksheet["C2"] = comment

# ------------------------------------------------------------ 
# 10. Format the worksheet 
# ------------------------------------------------------------

worksheet.column_dimensions["A"].width = 22 
worksheet.column_dimensions["B"].width = 45 
worksheet.column_dimensions["C"].width = 45

worksheet["A2"].alignment = Alignment(horizontal="center")
worksheet["B2"].alignment = Alignment(wrap_text=True)
worksheet["C2"].alignment = Alignment(wrap_text=True)

# ------------------------------------------------------------ 
# 11. Save Excel file 
# ------------------------------------------------------------

print("Step 6: Saving Excel file...")
workbook.save(excel_path)
print("Excel file saved:")

print(excel_path)

# ------------------------------------------------------------ 
#12. Open the generated Excel file 
# ------------------------------------------------------------ 

print("Step 7: Opening Excel file...")
# Use Windows Run dialog to open the Excel file 
pyautogui.hotkey("win", "r") 
wait(1)

pyautogui.write(f'"{excel_path}"')
pyautogui.press("enter")
# Give Excel enough time to open
wait(8)

# ------------------------------------------------------------
# 13. Open the generated Excel file
# ------------------------------------------------------------

print("Step 8: Opening generated Excel workbook...")

pyautogui.hotkey("ctrl", "o")

wait(3)

# Type the complete file path
pyautogui.write(excel_path, interval=0.01)

wait(1)

pyautogui.press("enter")

# Wait for workbook to load
wait(8)

# ------------------------------------------------------------
# 14. Make sure the Excel sheet is visible
# ------------------------------------------------------------

print("Step 9: Preparing final Excel sheet...")

pyautogui.hotkey("ctrl", "home")

wait(2)

# ------------------------------------------------------------
# 15. Take screenshot
# ------------------------------------------------------------

print("Step 10: Taking screenshot...")

screenshot = pyautogui.screenshot()

screenshot.save(screenshot_path)

print("Screenshot saved:")
print(screenshot_path)