import pyautogui
import time
import sys
import winsound  # Built-in Windows sound library

# ENABLE FAIL-SAFE
pyautogui.FAILSAFE = True

def alert_sound(frequency=1000, duration=500):
    """Plays a beep sound."""
    winsound.Beep(frequency, duration)

def run_tally_export():
    print("--- BOT STARTING ---")
    alert_sound(800, 300) # Start sound
    time.sleep(5) #After the sound go to the Tally window manually

    try:
        # Steps 2-7: Period Setup
        print("Step 2-7: Setting Period...")
        pyautogui.press('k')
        time.sleep(5)
        pyautogui.hotkey('alt', 'f2')
        time.sleep(2)
        pyautogui.write('a23', interval=0.1) #Desired period can be taken
        pyautogui.press('enter')
        pyautogui.write('m24', interval=0.1) #Desired period can be taken
        pyautogui.press('enter')

        # Step 8: Loading
        print("Step 8: Waiting for Daybook to load (15s)...")
        time.sleep(15)

        # Step 9-11: Export Configuration
        print("Step 9-11: Navigating Configuration...")
        pyautogui.hotkey('ctrl', 'e')
        time.sleep(2)
        pyautogui.press('c')
        time.sleep(2)
        for i in range(18):
            pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(2) #File format is taken as html and for me, as the key is 18 ways down from the start, it is taken as such

        # Step 12-15: Filename and Export
        print("Step 12-15: Renaming and Exporting...")
        for _ in range(30): 
            pyautogui.press('backspace')
        pyautogui.write('Userdaybook5.export.html', interval=0.05) #Here required unique name can be taken
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(2)
        pyautogui.press('e')

        # Final Step
        print("Step 16: Export complete.")
        alert_sound(1500, 1000) # High pitched 'Success' sound
        
        # Optional: Pop-up notification
        pyautogui.alert("The Tally Export Bot has finished successfully!", "Status: Finished")

    except pyautogui.FailSafeException:
        print("\n!!! BOT STOPPED: Fail-safe triggered (Mouse in corner) !!!")
        alert_sound(400, 1000) # Low pitched 'Warning' sound
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    run_tally_export()
