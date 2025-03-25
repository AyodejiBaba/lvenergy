import pyvisa
import pandas as pd
import time
import datetime
import os

# Initialize the VISA resource manager
rm = pyvisa.ResourceManager()

# List all connected instruments
available_instruments = rm.list_resources()

# Print the available instruments
if available_instruments:
    print("Available instruments:")
    for i, instr in enumerate(available_instruments):
        print(f"{i + 1}. {instr}")
else:
    print("No instruments found. Please check your connections.")
    exit()

# Connect to the first available instrument (modify as needed)
instrument_address = available_instruments[0]  # Default to first found device
print(f"Connecting to: {instrument_address}")

scope = rm.open_resource(instrument_address)
scope.timeout = 5000  # 5 seconds timeout

# Define the shunt resistor value (in ohms) for current calculations
SHUNT_RESISTOR = 100  # Adjust based on your setup

def extract_numeric_value(response, is_frequency=False):
    """Extracts the numeric value from the oscilloscope response string."""
    try:
        # Handle frequency formatting separately
        if is_frequency:
            value = response.split(",")[1].strip().replace("Hz", "").strip()
        else:
            value = response.split(",")[1].strip().replace("V", "")

        return float(value)
    except (IndexError, ValueError):
        print(f"Invalid response format: {response}")
        return None

def get_instrument_values():
    """Fetches RMS voltage, Peak-to-Peak voltage, and frequency from the oscilloscope."""
    try:
        # Query values from the oscilloscope
        rms_response = scope.query("C1:PAVA? CRMS")  # Example: 'C1:PAVA CRMS,2.120000E+00V\n'
        pkpk_response = scope.query("C1:PAVA? PKPK")  # Example: 'C1:PAVA PKPK,4.500000E+00V\n'
        freq_response = scope.query("C1:PAVA? FREQ")  # Example: 'C1:PAVA FREQ,1.000000E+03Hz\n'

        # Extract numeric values
        rms_voltage = extract_numeric_value(rms_response)
        pkpk_voltage = extract_numeric_value(pkpk_response)
        frequency = extract_numeric_value(freq_response, is_frequency=True)

        # Validate extracted values
        if None in [rms_voltage, pkpk_voltage, frequency]:
            return None  # Skip this measurement if any value is invalid

        # Calculate Current (I = V/R)
        current_rms = rms_voltage / SHUNT_RESISTOR
        current_pkpk = pkpk_voltage / SHUNT_RESISTOR

        # Calculate Power (P = VI)
        power_rms = rms_voltage * current_rms
        power_pkpk = pkpk_voltage * current_pkpk

        return {
            "RMS Voltage (V)": rms_voltage,
            "Pk-Pk Voltage (V)": pkpk_voltage,
            "Frequency (Hz)": f"{frequency:.4g}",  # Format frequency to 4 significant figures
            "Current RMS (A)": current_rms,
            "Current Pk-Pk (A)": current_pkpk,
            "Power RMS (W)": power_rms,
            "Power Pk-Pk (W)": power_pkpk,
            "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    except pyvisa.VisaIOError as e:
        print(f"Error querying the instrument: {e}")
        return None

def main():
    # Get user input for number of measurements
    num_measurements = int(input("Enter the number of measurements to take: "))
    save_folder = input("Enter the folder path to save the results: ")

    # Ensure folder exists
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)

    data_list = []

    # Take measurements
    for i in range(num_measurements):
        print(f"Measurement {i+1}...")
        result = get_instrument_values()
        if result:
            print(result)
            data_list.append(result)
        time.sleep(1)  # Small delay between measurements

    # Convert to DataFrame and save
    if data_list:
        df = pd.DataFrame(data_list)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(save_folder, f"measurement_{timestamp}.xlsx")
        df.to_excel(filename, index=False)
        print(f"Results saved to {filename}")

if __name__ == "__main__":
    main()

# Reopen the instrument connection for future use
scope = rm.open_resource(instrument_address)