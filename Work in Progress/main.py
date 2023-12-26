"""
By: Aaron Jauregui | Programs main file, handles first launch, and autonomous use.
"""
from resources import Device, open_onedrive, get_prefabs, save_prefabs
from time import sleep
import schedule
import threading
import json

TXT_DIV = "="  # This is the visual text divider you see for readability.
RUN_EVERY = 300  # This is the amount of seconds it will take between temperature reads. 900 = 15min, 1800 = 30min


def get_name_and_location(prefabs):
    print("A device name and location are needed...")
    name = input("Enter a Device Name: ")
    location = input(f"Enter the desired location of {name}: ")

    print("Confirm Your Info!")
    print("Name: ", name)
    print("Location: ", location)

    confirm = "n"
    try:
        confirm = input("Confirm the information entered (Y/N): ").lower()
    except AttributeError:
        print("Please enter a valid answer (Y/N)...")

    if confirm == "y":
        print("Information confirmed. Thank you!")
        prefabs["CONFIG"]["HAS_RUN"] = True
        prefabs["CONFIG"]["NAME"] = name
        prefabs["CONFIG"]["LOCATION"] = location
        save_prefabs(prefabs)
        return name, location
    else:
        print("Information not confirmed. Please try again.")

def main():
    print(TXT_DIV * 30)
    print("Initializing Device...")
    onedrive = threading.Thread(target=open_onedrive, args=(True,))
    prefabs = get_prefabs()
    if prefabs["CONFIG"]["HAS_RUN"] is not True:
        name, location = get_name_and_location(prefabs)
    print(TXT_DIV * 30)

    onedrive.start()
    device = Device(name, location)
    device.start_device()
    schedule.every(RUN_EVERY).seconds.do(device.start_device)
    while True:
        try:
            schedule.run_pending()
        except PermissionError:  # Possibly have a file open on a pc.
            print("Permission Denied. Close Any Open Excel Files.")
            sleep(3)
        except Exception as exc:  # Error in the main file.
            print(f"Error M: {exc}")
            sleep(3)
        sleep(1)


if __name__ == "__main__":
    main()
