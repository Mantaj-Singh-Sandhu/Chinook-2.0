import os
import argparse

def update_input_file(parameter, devices_file="devices_list.txt", input_file="input_file.txt"):
    """
    Reads the devices_list.txt file, copies the line at the given parameter (integer index),
    and updates it in the input_file.txt with .parquet added to the line.

    :param parameter: The parameter (integer index) to select the line in devices_list.txt
    :param devices_file: The source file containing the list of devices
    :param input_file: The target file to update with the selected line
    """
    if not os.path.exists(devices_file):
        print(f"Error: {devices_file} does not exist.")
        return

    try:
        # Read devices_list.txt and get the line at the given index (parameter)
        with open(devices_file, "r") as df:
            lines = df.readlines()

        # Ensure the parameter is a valid line index
        if parameter < 0 or parameter >= len(lines):
            print(f"Error: Parameter {parameter} is out of range. File has {len(lines)} lines.")
            return

        # Get the specific line and add .parquet
        selected_line = lines[parameter].strip() + ".parquet"

        # Write the selected line to input_file.txt
        with open(input_file, "w") as inf:
            inf.write(selected_line + "\n")

        print(f"Successfully updated {input_file} with the line at index {parameter}, with .parquet added.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Update input_file.txt with the line at a specific index from devices_list.txt.")
    parser.add_argument("parameter", type=int, help="The integer index of the line in devices_list.txt")
    args = parser.parse_args()

    update_input_file(args.parameter)
