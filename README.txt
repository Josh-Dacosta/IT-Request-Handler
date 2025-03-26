CSCI 4307 AI Project
AI Optimization in IT Office
=============================

This script helps to manage IT device inventory by parsing emails, extracting required information through Natural Language Processing, and adding new devices to an existing inventory Excel file. It contains features for asset tag creation and integration with a device database for validation.

Prerequisites
-------------
Before running the script, make sure you have the following tools and libraries set up:

- Python 3.x
- The required Python libraries: `spaCy`, `openpyxl`, `pandas`
- spaCy model (`en_core_web_sm`)
- A device database text file (`devices_db.txt`)
- A device inventory Excel file (`inventory.xlsx`)

Step-by-Step Setup
------------------

1. Install Python Dependencies

First, ensure that you have Python 3.x installed. Then, install the necessary libraries via `pip`:

    pip install spacy openpyxl pandas

You will also need to download the `en_core_web_sm` spaCy model. Run the following command to download it:

    python -m spacy download en_core_web_sm

2. Create or Obtain the Device Database File

The device database (`devices_db.txt`) should contain information about the devices your company uses. Each device should be described in a comma-separated format like this:

    Dell, Latitude 5460, Intel i7, 16 GB, 512 GB, Windows 10 Enterprise

This file will be used by the script to validate whether the requested device matches the company's inventory.

3. Prepare the Device Inventory File

The device inventory (`inventory.xlsx`) should be an Excel file with the following columns:

- **Department**: The department that requested the device.
- **Room**: The room number where the device will be used.
- **Asset Tag**: A unique asset tag for each device.
- **Device Name**: The name of the device (e.g., laptop, desktop).
- **Advisor**: The name of the person using the device.
- **Make**: The make of the device (e.g., Dell).
- **Model**: The model of the device (e.g., Latitude, Optiplex).
- **Purchase Order**: The associated purchase order number.

If the `inventory.xlsx` file doesn't exist, the script will automatically create one with the correct structure.

4. Organize Your Files

Ensure your project directory contains the following files:

    - inventory.xlsx       # Excel file containing device inventory (can be empty if not existing)
    - devices_db.txt       # Text file containing device specifications
    - script.py            # The Python script to manage the inventory

You can adjust the file paths in the script if these files are located in different directories.

5. Running the Script

After setting up the files, you can run the script in your Python environment. The script will:

1. Parse incoming emails (ensure the email text is properly formatted).
2. Extract information such as department, room, advisor, device name, make, and model.
3. Generate an asset tag and purchase order.
4. Add the new device to the inventory Excel file.

Run the script with the following command:

    python script.py

The script will either add a new device to the inventory or suggest a similar device if the requested device is not valid according to the device database.

6. Example Email Format

For the script to correctly extract information from the email, ensure the email follows this format:

    Hello IT Support,

    I need a Dell Latitude 7420 with Intel Core i7, 16 GB RAM, and 512 GB SSD for the HR department. This device will be used by Ted Burns in Room 151.

    Please use Purchase Order PO12345 for this request.

The email should mention the device's make, model, department, room, and advisor.