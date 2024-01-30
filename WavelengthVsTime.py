import time
import serial
import pyvisa as visa
import struct
import numpy as np
import pandas as pd
# Oriel Monochromator connection
mono_port = 'COM3'
mono_baud_rate = 9600  # Verify the baud rate in the monochromator manual.
mono = serial.Serial(mono_port, mono_baud_rate, timeout=2000)

# Keithley 2400 SourceMeter unit setup
rm = visa.ResourceManager()
smu = rm.open_resource('ASRL6::INSTR')
smu.timeout = 60000  # Set command timeout to 60 seconds
smu.read_termination = '\r'
smu.write_termination = '\r'

# Reset and configure Keithley 2400 for current measurement
smu.write('*RST')  # Reset the instrument
smu.write(':SENS:FUNC "CURR"')  # Set to measure current
smu.write(':SENS:CURR:RANG 1e-6')  # Set current range to 1 microampere
smu.write(':SENS:CURR:NPLC 1')  # Set Number of Power Line Cycles (NPLC) to 1 for faster measurements
smu.write(':OUTP ON')  # Turn on the output

# Function to set wavelength on the monochromator
def set_wavelength(mono, wavelength):
    command = f"GOWAVE {wavelength}\n"
    mono.write(command.encode('ASCII'))
    _ = mono.readline()

# Function to open monochromator shutter
def open_shutter(mono):
    mono.write(b'SHUTTER O\n')
    _ = mono.readline()

# Function to close monochromator shutter
def close_shutter(mono):
    mono.write(b'SHUTTER C\n')
    _ = mono.readline()

# Measurement parameters
measurement_interval = 0.5  # Interval in seconds between measurements
sleep_interval = measurement_interval - 0.139  # Adjust for command overhead time
measurement_duration = 2  # Duration in seconds for measurements at each wavelength

# Define wavelengths for measurements
wavelengths = np.arange(500, 200, -5)  # Corrected to decrement for typical scanning

# Initialize dictionary to store time and current data for each wavelength
data_by_wavelength = {}

# Main measurement loop
close_shutter(mono)  # Ensure shutter is closed before starting measurements
for wavelength in wavelengths:
    set_wavelength(mono, wavelength)  # Set monochromator to the current wavelength
    open_shutter(mono)  # Open shutter to start light exposure
    smu.write(':OUTP ON')  # Ensure the SourceMeter output is on

    current_readings = []  # List to store current readings and timestamps
    start_time = time.time()  # Record start time for the current wavelength
    
    while time.time() - start_time < measurement_duration:
        response = smu.query(':READ?')
        data = response.split(',')  # Split response into components
        current = float(data[1]) * -1e9  # Convert current to nA and correct polarity if needed
        current_time = time.time() - start_time  # Calculate elapsed time since start
        current_readings.append((current_time, current))  # Append time and current to list
        time.sleep(sleep_interval)  # Wait before next measurement

    data_by_wavelength[wavelength] = current_readings  # Store data for current wavelength
    close_shutter(mono)  # Close shutter after measurements
    time.sleep(1)  # Short pause before moving to next wavelength

mono.close()  # Close monochromator connection

# Function to restructure and save all data in the desired format
def save_to_excel_custom_format(data, filename="measurement_data_custom_format.xlsx"):
    # Initialize a list to hold all data points
    all_data = []

    # Iterate over each wavelength and its readings
    for wavelength, readings in data.items():
        for time, current in readings:
            # Append a tuple for each reading: (wavelength, time, current)
            all_data.append((wavelength, time, current))

    # Convert the list of tuples into a DataFrame
    df = pd.DataFrame(all_data, columns=['Wavelength (nm)', 'Time (s)', 'Current (nA)'])

    # Save the DataFrame to an Excel file
    df.to_excel(filename, index=False)

# Call the function to save the data
save_to_excel_custom_format(data_by_wavelength)
