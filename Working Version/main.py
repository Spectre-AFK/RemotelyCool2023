from resources import Device, save_file, find_file, confirm, open_onedrive
from time import sleep
import schedule
import threading

text_devider = "="  #  This is the visual text divider you see for readability.
run_every = 300  # This is the amount of seconds it will take between temperature reads. 900 = 15min, 1800 = 30min

def main():
    print(text_devider * 30)
    print("Initializing Device...")
    onedrive = threading.Thread(target=open_onedrive, args=(True,))
    # Getting Name
    name = find_file("name")
    if name == False:
        name = confirm("Name Device: ")
        save_file("name", name)
    else:
        print(f"Found Name: '{name}'")

    # Getting Location
    location = find_file("location")
    if location == False:
        location = confirm("Set Location: ")
        save_file("location", location)
    else:
        print(f"Found Location: '{location}'")
        
    onedrive.start()
    print(text_devider * 30)
    device_1 = Device(name, location)
    
    #Test Start
    print(f"Saving Starting Temperature...")
    device_1.start_device()
    
    schedule.every(run_every).seconds.do(device_1.start_device)
    while True:
        try:
            schedule.run_pending()
        except PermissionError:
            print("Permission Denied. Close Any Open Excel Files.")
            sleep(3)
        except Exception as exc:
            print(f"Error M: {exc}")
            sleep(3)
        sleep(1)


if __name__ == "__main__":
    main()
