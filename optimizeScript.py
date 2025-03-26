import spacy                        # For Natural Language Processing
from spacy.matcher import PhraseMatcher
import openpyxl                     # For Excel
import pandas as pd                 # For Excel
from email.parser import Parser     # Reading 'email' text files
import uuid                         # Create a new Asset tag

# Initialize spaCy model
nlp = spacy.load("en_core_web_sm")

# Variable for asset tag
asset_tag_counter = 1  # Base starting point for asset tag

# Load company's existing inventory (assuming Excel file)
def load_inventory(file_path):
    try:
        inventload = pd.read_excel(file_path)
        return inventload
    except FileNotFoundError:
        # For no existing file, return empty excel sheet with the formatted columns
        return pd.DataFrame(columns=["Department", "Room", "Asset Tag", "Device Name", "Advisor", "Make", "Model", "Purchase Order"])

# Find the last used asset tag to update the counter variable 
def get_last_asset_tag(file_path):
    inventload = load_inventory(file_path)
    if not inventload.empty:
        # Return the last asset tag from the 'Asset Tag' column ASSUMING DATA IS SORTED
        last_tag_number = inventload["Asset Tag"].iloc[-1]  # The last entry
        return last_tag_number
    else:
        return 1  # Example starting point for asset tags

# Create a new asset tag for the device
def update_asset_tag(file_path):
    global asset_tag_counter
    # call the last asset tag to retrieve current data
    asset_tag_counter = get_last_asset_tag(file_path)
    asset_tag_counter += 1  # Increment for the next asset tag# Ensure the asset tag is always 4 digits, padded with leading zeros
    return f"[{asset_tag_counter:04}]"

# Starting point for Purchase Order
purchase_order_counter = 1000  # Example starting point for Purchase Orders

def get_last_purchase_order(file_path):
    global purchase_order_counter
    df = load_inventory(file_path)
    if not df.empty:
        # Extract the last purchase order from the 'Purchase Order' column
        last_po = df["Purchase Order"].iloc[-1]  # Get the last entry
        # Extract the numeric part of the PO (removes 'PO' prefix)
        last_po_number = int(last_po.strip("PO"))
        purchase_order_counter = last_po_number  # Update the counter to the last PO number
    return purchase_order_counter

# Make a new Purchase Order number based on the last one from the Inventory file
def update_purchase_order(file_path):
    global purchase_order_counter
    purchase_order_counter = get_last_purchase_order(file_path)
    purchase_order_counter += 1  # Increment for the next purchase order
    # Return the new purchase order in the format PO0000
    return f"PO{purchase_order_counter:04}"

# Retrieve the acceptable devices dataset
def load_device_database(db_path):
    device_db = []  # Initialize an array to hold the data
    try:
        with open(db_path, 'r') as file:  # Open the text file with reading permission
            for line in file:
                parts = line.strip().split(", ")
                if len(parts) == 6:
                    make, model, processor, ram, storage, os = parts
                    device_db.append({
                        "Make": make,
                        "Model": model,
                        "Processor": processor,
                        "Ram": ram,
                        "Storage": storage,
                        "Operating System": os
                    })
    except FileNotFoundError:
        print(f"Device database file not found: {db_path}")
    return device_db

# Use spaCy to pull the required information from the email.
def extract_entities_from_email(email_text, device_db):
    """Extract entities from the email using spaCy and custom pattern matching."""
    doc = nlp(email_text)
    
    details = {
        "Department": None,
        "Room": None,
        "Advisor": None,
        "Device Name": None,
        "Make": None,
        "Model": None,
        "Processor": None,
        "Ram": None,
        "Storage": None,
        "Operating System": None,
        "Device Valid": False,
        "Suggested Device": None
    }

    # Convert each entity (using NER for Department, Advisor, and Room)
    for ent in doc.ents:
        if ent.label_ == "ORG":  # Department might be recognized as organization names
            details["Department"] = ent.text
        elif ent.label_ == "GPE":  # Location-related entities (could be Room or Location)
            details["Room"] = ent.text
        elif ent.label_ == "PERSON":  # Advisor names recognized as PERSON
            details["Advisor"] = ent.text

    # Set up a PhraseMatcher to match specific devices, make, and models to match to the dataset
    matcher = PhraseMatcher(nlp.vocab)
    
    # Device names and makes we're interested in (expandable list)
    devices = ["laptop", "desktop", "tablet"]
    makes = ["Dell", "Dell", "Dell"]
    models = ["Latitude", "Optiplex", "Precision"]

    # Identify patterns in the text
    device_patterns = [nlp.make_doc(device) for device in devices]
    make_patterns = [nlp.make_doc(make) for make in makes]
    model_patterns = [nlp.make_doc(model) for model in models]

    # Add the matched data to the pattern set
    matcher.add("Device", device_patterns)
    matcher.add("Make", make_patterns)
    matcher.add("Model", model_patterns)
    
    # Find matches in the document
    matches = matcher(doc)

    for match_id, start, end in matches:
        matched_span = doc[start:end]
        if "laptop" in matched_span.text.lower() or "desktop" in matched_span.text.lower():
            details["Device Name"] = matched_span.text
        elif matched_span.text in makes:
            details["Make"] = matched_span.text
        elif matched_span.text in models:
            details["Model"] = matched_span.text


    # Validate the device against the accepted database
    if details["Make"] and details["Model"]:
        for device in device_db:
            if (device["Make"] == details["Make"] and 
                device["Model"] == details["Model"] and 
                (not details["Processor"] or device["Processor"] == details["Processor"]) and
                (not details["Ram"] or device["Ram"] == details["Ram"]) and
                (not details["Storage"] or device["Storage"] == details["Storage"]) and
                (not details["Operating System"] or device["Operating System"] == details["Operating System"])):
                details["Device Valid"] = True
                break
        # If the device is not valid, suggest a similar one
        if not details["Device Valid"]:
            details["Suggested Device"] = suggest_similar_device(details, device_db)

    return details

def suggest_similar_device(details, device_db):
    """Suggest a similar device from the database based on the request."""
    best_match = None
    best_score = 0

    # Compare the requested device's specifications with each device in the database
    for device in device_db:
        score = 0
        if device["Make"] == details["Make"]:
            score += 1
        if device["Processor"] == details["Processor"]:
            score += 1
        if device["Ram"] == details["Ram"]:
            score += 1
        if device["Storage"] == details["Storage"]:
            score += 1
        if device["Operating System"] == details["Operating System"]:
            score += 1

        # Track the best match based on the score
        if score > best_score:
            best_score = score
            best_match = device

    # If we found a match, return a suggestion
    if best_match:
        return f"Suggested Device: {best_match['Make']} {best_match['Model']} with {best_match['Processor']}, {best_match['Ram']} RAM, {best_match['Storage']} Storage, {best_match['Operating System']}"

    return "No similar device found in the database."

def create_inventory_entry(details, file_path):
    """Create a new inventory entry."""
    new_entry = [
        details["Department"],
        details["Room"] if details["Room"] else "N/A",  # Default to "N/A" if no room is found
        update_asset_tag(file_path),
        details["Device Name"] if details["Device Name"] else "Unknown Device",  # Default if no device found
        details["Advisor"] if details["Advisor"] else "Unknown Advisor",  # Default if no advisor found
        details["Make"] if details["Make"] else "Unknown Make",  # Default if no make found
        details["Model"] if details["Model"] else "Unknown Model",  # Default if no model found
        update_purchase_order(file_path)  # Create a new PO
    ]
    return new_entry

def generate_work_note(details):
    """Generate a work note (email) requesting missing information or confirming device."""
    missing_fields = []

    # Identify missing fields and generate appropriate responses
    if not details["Department"]:
        missing_fields.append("Department")
    if not details["Advisor"]:
        missing_fields.append("Advisor")
    if not details["Make"]:
        missing_fields.append("Make")
    if not details["Model"]:
        missing_fields.append("Model")
    
    # Generate the response message
    if missing_fields:
        message = "Dear Client,\n\nWe received your request, but we noticed that some necessary information is missing or requires clarification. Could you please provide the following details?\n\n"
        for field in missing_fields:
            message += f"- {field}\n"
        message += "\nOnce we have this information, we can proceed with processing your request.\n\nThank you for your understanding.\n\nBest regards,\nYour IT Support Team"
    else:
        message = "Thank you for your request. We are processing your order and will get back to you with the next steps shortly.\n\nBest regards,\nYour IT Support Team"

    return message

def update_inventory(file_path, new_entry):
    """Update the inventory Excel file."""
    inventload = load_inventory(file_path)
    # Add the new entry to the dataframe
    inventload.loc[len(inventload)] = new_entry
    # Save back to Excel
    inventload.to_excel(file_path, index=False)

def parse_email_and_update_inventory(email_text, file_path, db_path):
    """Main function to parse the email and update the inventory."""
    # Load the device database
    device_db = load_device_database(db_path)

    # Process the email to extract the required information
    details = extract_entities_from_email(email_text, device_db)
    
    # Check if there is any missing information
    if details["Device Valid"] and details["Device Name"]:  # Ensure Device Name is valid
        # Create the new inventory entry
        new_entry = create_inventory_entry(details)
        # Update the inventory file
        update_inventory(file_path, new_entry)
        print(f"Inventory updated with asset tag {new_entry[2]}")
    else:
        # If missing data or invalid device, generate a work note and print it (or send it to the client)
        work_note = generate_work_note(details)
        print(work_note)

# Example usage:

email_text = """ 
Hello IT Support,

I need a Dell Latitude 7420 with Intel Core i7, 16 GB RAM, and 512 GB SSD for the HR department. This device will be used by John Doe in Room 101.

Please use Purchase Order PO12345 for this request.Thank you,
[Manager's Name]
Product Development Department
"""
file_path = "./datasets/inventory.xlsx"
db_path = "./datasets/devices_db.txt"
"""
with open(inv_path, 'r') as file_path:
    content = file_path.read()

with open(db_path, 'r') as device_db_file:
    contentdata = device_db_file()
"""
parse_email_and_update_inventory(email_text, file_path, db_path)