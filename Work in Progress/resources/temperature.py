import os
from time import sleep

device_names = []

def read(device_name):
    location = f"/sys/bus/w1/devices/{device_name}/w1_slave"
    with open(location, "r") as file:
        text = file.read()
    secondline = text.split("\n")[1]
    temperature_data = secondline.split(" ")[9]
    temperature = float(temperature_data[2:])
    celsius = temperature / 1000
    fahrenheit = (celsius * 1.8) + 32
    return celsius, fahrenheit

def temperature_read(device):
    device_serial = device
    if read(device) is not None:
        temp_f = read(device)[1]
        temp_c = read(device)[0]
        return temp_f, temp_c
    else:
        print("Error: No Device Found.")

def sensor():
    for devices in os.listdir("/sys/bus/w1/devices"):
        if devices != "w1_bus_master1":
            device = devices
            device_names.append(devices)
    return device

def main():
    while True:
        try:
            serial_num = sensor()
            temp_f, temp_c = temperature_read(serial_num)
            temp_f = "{0:.2f}".format(temp_f)
            temp_c = "{0:.2f}".format(temp_c)
            return temp_f, temp_c
        except Exception as exc:
            print(f"Error S: {exc}")
        sleep(1)


class Temperature_Probe:
    def __init__(self, name, id):
        self.name = name
        self.id = id
    
    def action(self):
        temp_f, temp_c = main()
        return temp_f, temp_c


if __name__ == "__main__":
    main()