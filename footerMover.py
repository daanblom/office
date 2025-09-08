import time
import pyautogui
import keyboard

def main():
    print("Starting in 10 seconds, prepare footer... Press 'q' anytime to quit.")
    time.sleep(10)

    while True:
        if keyboard.is_pressed('q'):
            print("\nExiting.")
            break

        # Press Delete
        pyautogui.press('delete')
        
        time.sleep(0.12)

        # Press Ctrl+V
        pyautogui.hotkey('ctrl', 'v')

        time.sleep(0.12)

        # Press Page Down
        pyautogui.press('pagedown')

        time.sleep(0.12)

        # Left mouse click
        pyautogui.click(button='left')

        time.sleep(0.12)

if __name__ == "__main__":
    main()

