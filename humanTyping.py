import pyautogui
import time
import random
import argparse

# -----------------------
# Configurable variables
# -----------------------

WAIT_TIME = 5          # seconds before typing starts
MIN_DELAY = 0.05       # minimum delay between keystrokes
MAX_DELAY = 0.15       # maximum delay between keystrokes
SPACE_DELAY = 0.3      # delay when space is pressed
ENTER_DELAY = 0.4      # delay when enter/newline is pressed


def human_type(text):
    for char in text:
        if char == " ":
            pyautogui.press("space")
            time.sleep(SPACE_DELAY)

        elif char == "\n":
            pyautogui.press("enter")
            time.sleep(ENTER_DELAY)

        else:
            pyautogui.write(char)
            time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))


def main(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        text = f.read()

    print(f"Typing will start in {WAIT_TIME} seconds...")
    time.sleep(WAIT_TIME)

    human_type(text)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Human-like typing script")
    parser.add_argument("file", help="Path to the txt file to type")

    args = parser.parse_args()

    main(args.file)
