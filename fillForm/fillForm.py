import time
import pyautogui
import argparse

# Parse arguments
parser = argparse.ArgumentParser(description='Automatically fill in a form using input from a text file.')
parser.add_argument('-i', '--input', required=True, help='Path to input form text file')
args = parser.parse_args()

# Load form input
with open(args.input, 'r', encoding='utf-8') as file:
    lines = [line.strip() for line in file.readlines()]

print("You have 3 seconds to switch to your form window...")
time.sleep(3)

# Go through each line and perform actions
for line in lines:
    if line == "%SKIP":
        pyautogui.press('tab')
    elif line == "%FIRSTENTRY":
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('tab')
    elif line == "%ENTERDATE":
        pyautogui.press('enter')
        pyautogui.press('enter')
        time.sleep(0.3)
        pyautogui.press('tab')
    else:
        pyautogui.write(line)
        pyautogui.press('tab')
