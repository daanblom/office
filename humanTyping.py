import pyautogui
import time
import random
import argparse

# -----------------------
# Configurable variables
# -----------------------

WAIT_TIME = 5
MIN_DELAY = 0.001
MAX_DELAY = 0.005
SPACE_DELAY = 0.05
PERIOD_DELAY = 0.10      # delay after typing "."
ENTER_DELAY = 0.15


def normalize_text(text: str) -> str:
    # Replace common "special spaces" with a normal space
    return (
        text.replace("\u00A0", " ")   # non-breaking space
            .replace("\u202F", " ")   # narrow no-break space
            .replace("\u2007", " ")   # figure space
            .replace("\r\n", "\n")    # windows newlines
            .replace("\r", "\n")      # old mac newlines
    )


def human_type(text):
    for char in text:
        if char == " ":
            pyautogui.press("space")
            time.sleep(SPACE_DELAY)

        elif char == "\n":
            pyautogui.hotkey("shift", "enter")
            time.sleep(ENTER_DELAY)

        elif char == ".":
            pyautogui.write(".")
            time.sleep(PERIOD_DELAY)

        else:
            pyautogui.write(char)
            time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))


def main(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        text = f.read()

    text = normalize_text(text)

    print(f"Typing will start in {WAIT_TIME} seconds...")
    print("Click inside the text field now.")
    time.sleep(WAIT_TIME)

    human_type(text)
    print("Done.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Human-like typing script")
    parser.add_argument("file", help="Path to the txt file to type")
    args = parser.parse_args()
    main(args.file)