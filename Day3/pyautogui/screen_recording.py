import pyautogui

print("Taking screenshot...")

image = pyautogui.screenshot()

image.save("test_screenshot.png")

print("Screenshot successful!")