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

# Constants for bias voltage setup
bias_voltage = 1.0  # Bias voltage in volts (V)

# Configure Keithley 2400 to source voltage
smu.write(':SOUR:FUNC VOLT')  # Set the instrument to source voltage
smu.write(f':SOUR:VOLT:LEV {bias_voltage}')  # Set the bias voltage level
smu.write(':SOUR:VOLT:RANG 20')  # Set the voltage range
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

def save_to_excel_with_interval(data, measurement_interval, filename="measurement_data_interval.xlsx"):
    # Determine the maximum number of measurements across all wavelengths to set up the time column
    max_measurements = max(len(readings) for readings in data.values())
    
    # Generate the time column based on the maximum number of measurements and the measurement interval
    time_column = np.arange(0, max_measurements * measurement_interval, measurement_interval)
    
    # Prepare an empty DataFrame with the time column initialized
    df = pd.DataFrame(time_column, columns=['Time (s)'])
    
    # Fill in the DataFrame with current values for each wavelength
    for wavelength, readings in data.items():
        # Extract current values, assuming the time interval is consistent across all measurements
        current_values = [current for _, current in readings]
        
        # If the number of readings for a wavelength is less than max_measurements, pad with NaNs
        if len(current_values) < max_measurements:
            current_values += [np.nan] * (max_measurements - len(current_values))
        
        # Add the current values as a new column to the DataFrame, naming the column after the wavelength
        df[f'Current at {wavelength} nm (nA)'] = current_values
    
    # Save the DataFrame to an Excel file
    df.to_excel(filename, index=False)

# Assuming data_by_wavelength is already defined and populated, and measurement_interval is known
save_to_excel_with_interval(data_by_wavelength, measurement_interval)
