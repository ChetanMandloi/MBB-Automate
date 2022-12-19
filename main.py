import serial
import keyboard
import time
import numpy as np
import pandas as pd
# openpyxl

def GetLengthList():
    while True:
        try:
            start = float(input("Type starting Position in cm (0.0 - 30.0): "))
            if start < 0 or start > 30:
                print("Out of valid Range. Try again")
                continue
            end = float(input("Type ending Position in cm (0.0 - 30.0): "))
            if end < 0 or end > 30 or end < start:
                print("Out of valid Range. Try again")
                continue
            step_size = float(input("Type step size in cm (0.1 - 1.0): "))
            if step_size < 0.1 or step_size > 1.0:
                print("Out of valid Range. Try again")
                continue
            break
        except ValueError:
            print("Invalid input. Try again")
            continue
    return np.arange(start, end + step_size, step_size)


def main():
    # make sure the 'COM#' is set according the Windows Device Manager
    ser = serial.Serial('COM4', 9600, timeout=1)
    time.sleep(5)

    pos_list = GetLengthList()
    mag_list = []
    print("Press the Shift once to take each reading.")
    print("Position(cm)\tSensor Value(V)")
    for pos in pos_list:
        print("%.2f" % pos, end='\t\t\t')
        keyboard.wait("shift")
        ser.reset_input_buffer()
        line = ser.readline()
        string = line.decode()  # convert the byte string to a unicode string
        mag_list.append(float(string))  # convert the unicode string to an int
        print(string)
        time.sleep(0.2)
    ser.close()
    export_df = pd.DataFrame(list(zip(pos_list, mag_list)), columns=["Position(cm)", "Sensor Value(V)"])
    export_df.to_excel("MBB.xlsx", index=False)
    print("File exported as MBB.xlsx")


if __name__ == "__main__":
    main()
