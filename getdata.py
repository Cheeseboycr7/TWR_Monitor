#!/usr/bin/python

import logging
import serial
import time


def get_serial_device(device_file, baudrate=9600):
    """
    Get a serial device object.

    Args:
        device_file (str): The path to the serial device file.
        baudrate (int, optional): The baud rate to use. Defaults to 9600.

    Returns:
        serial.Serial: The serial device object.
    """
    return serial.Serial(port=device_file, baudrate=baudrate,
                         parity=serial.PARITY_NONE,
                         stopbits=serial.STOPBITS_ONE,
                         bytesize=serial.EIGHTBITS,
                         timeout=1
                         )


# This function contains the main logic
def main():
    """
    Main function.
    """
    # Path to serial device file for communication
    device = "/dev/ttyUSB0"
    # init the logging function
    logging.basicConfig(
        filename="./data.log",
        level=logging.INFO,
        format="%(asctime)s : %(message)s",
        datefmt="%Y-%m-%d %H:%M",
    )

    # call the code below with the created device file object ser
    with get_serial_device(device) as ser:
        # run script again and again
        while True:
            print("Start collectiong data ...")
            # run the write and read steps on each command specified in this list
            for command in ["DA\r", "DB\r", "DG\r", "DH\r", "DN\r"]:
                ser.write(command.encode())
                time.sleep(0.1)
                try:
                    data = ser.readline().decode()
                except serial.SerialException as e:
                    logging.error("Error reading from serial device: %s", e)
                    continue
                logging.info(";Data for %s :; %s ", command.replace("\r", ""), data)
                # pause for 100ms after reading data before write the next command
                time.sleep(0.1)

            print("collected data, going to sleep")
            #  logging.info("=================")
            # Waiting 30 seconds before run write and read for all commands again
            time.sleep(60)


if __name__ == "__main__":
    main()
