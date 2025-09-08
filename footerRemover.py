import time
import pyautogui
import keyboard

def main():
    print("Starting in 5 seconds, prepare footer... Press 'q' anytime to quit.")
    time.sleep(5)

    while True:
        if keyboard.is_pressed('q'):
            print("\nExiting.")
            break

        # Press Delete
        pyautogui.press('delete')


        # Press Page Down
        pyautogui.press('pagedown')

        # Left mouse click
        pyautogui.click(button='left')

        # Wait x seconds
        time.sleep(1)

if __name__ == "__main__":
    main()

