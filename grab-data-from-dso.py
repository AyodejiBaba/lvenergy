pwd         # check current working directory
# change to desired working directory

import os

# Change to your desired directory (update the path accordingly)
desired_directory = "c:\\Users\\ayode\\Downloads\\Azure Data Engineer\\audio_dsp\\lvenergy-results"  # Change this path
os.chdir(desired_directory)  # Change working directory

print(f"Current working directory: {os.getcwd()}")  # Verify the change

# check if measurement oscilloscope is available for remote connection via VISA

import pyvisa
import csv
import time

rm = pyvisa.ResourceManager()
devices = rm.list_resources()
print("Available VISA devices:", devices)
# Initialize VISA connection and confirm Connection

rm = pyvisa.ResourceManager()
scope = rm.open_resource("USB0::0xF4EC::0xEE38::SDSMMFCD6R2214::INSTR")
scope.timeout = 10000  # Set timeout to 10 seconds
print("Connected to:", scope)
# use this to clear the scope if miscellanous error shows again.

# scope.timeout = 20000  # 20 seconds
# scope.clear()
# simple request to test connnection to measuring oscilloscope

response = scope.query("*OPC?")
print(response)
# test to retrieve pk-pk from osc
pk_pk = scope.query("C1:PAVA? PKPK")
print(pk_pk)
# Measurement begins

import pyvisa
import csv
import time


# Track start time
start_time = time.time()

# Initialize VISA connection
rm = pyvisa.ResourceManager()
scope = rm.open_resource("USB0::0xF4EC::0xEE38::SDSMMFCD6R2214::INSTR")  # Update with your oscilloscope address
scope.timeout = 10000  # 10 seconds timeout

# Define measurement queries
measurements = {
    "PKPK": "C1:PAVA? PKPK",
    "RMS": "C1:PAVA? CRMS",
    "FREQ": "C1:PAVA? FREQ"
}

# Open CSV file in append mode
csv_filename = "tests"
# csv_filename = "piezo_only_on_speaker-vib-major.csv"
with open(csv_filename, "a", newline="") as file:
    writer = csv.writer(file)

    # Write header if file is empty
    file.seek(0, 2)  # Move to the end
    if file.tell() == 0:
        writer.writerow(["Time (s)", "PKPK (V)", "RMS (V)", "Frequency (Hz)"])

    # Collect multiple measurements
    for i in range(3):  # Adjust the number of samples
        timestamp = time.time()  # Get current time
        row_data = [f"{timestamp:.3f}"]  # Format timestamp to 3 decimal places

        for key, command in measurements.items():

            # time.sleep(0.15)

            try:
                response = scope.query(command)  # Query oscilloscope

                value = response.split(",")[1].strip()  # Extract value & clean it
                numeric_value = float(value.strip("VHz"))  # Convert to float
                formatted_value = f"{numeric_value:.5g}"  # Format to 3 significant figures
                row_data.append(formatted_value)
                print(f"{key}: {formatted_value}")  # Display result
            except Exception as e:
                print(f"Error retrieving {key}: {e}")
                row_data.append("ERROR")  # Mark errors in CSV

        # Write clean row to CSV
        writer.writerow(row_data)

        print(f"Data saved: {row_data}\n")
        time.sleep(1)  # Adjust sampling rate

# print(f"✅ Measurement completed. Data saved in {csv_filename}")
# Track end time
end_time = time.time()
execution_time = end_time - start_time  # Calculate duration

print(f"✅ Measurement completed. Data saved in {csv_filename}")
print(f"⏱️ Execution time: {execution_time:.2f} seconds")
