from pynput import mouse, keyboard

def on_click(x, y, button, pressed):
    if pressed:
        print(f"MOUSE CLICK: {button}")

def on_press(key):
    print(f"KEY PRESS: {key}")

print("Listening for events...")
ml = mouse.Listener(on_click=on_click)
kl = keyboard.Listener(on_press=on_press)
ml.start()
kl.start()
ml.join()
