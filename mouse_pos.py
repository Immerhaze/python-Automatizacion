import pyautogui
import time
from pynput import mouse

# Define a callback function to be called when the mouse is clicked
def on_click(x, y, button, pressed):
    if pressed:
        print(f"\nMouse clicked at ({x}, {y})")

# Set up the listener to call the on_click function when a click is detected
listener = mouse.Listener(on_click=on_click)
listener.start()

try:
    while True:
        # Get the current mouse position
        x, y = pyautogui.position()
        # Print the mouse position
        print(f"Current mouse position: ({x}, {y})", end="\r")
        # Pause for a short time to avoid flooding the output
        time.sleep(0.1)
except KeyboardInterrupt:
    print("\nDone")
    listener.stop()
