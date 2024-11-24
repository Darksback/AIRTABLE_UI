import customtkinter as ctk
import requests
import threading
import pandas as pd
import os
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from datetime import datetime
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed

# Airtable Configuration
API_KEY = "patJNieyOmCdDcA51.635e3b384caca979ac0666c13cd2516185d719ecb988abfd939309f793753d56"
AIRTABLE_URL = f"https://api.airtable.com/v0/appoGD1yWNzgOPH0O/TPD"

headers = {"Authorization": f"Bearer {API_KEY}"}

# UPS XML Request Template
UPS_XML_TEMPLATE = """<?xml version="1.0"?>
<AccessRequest>
  <AccessLicenseNumber>4D67F229A484D4F5</AccessLicenseNumber>
  <UserId>suhailups</UserId>
  <Password>password</Password>
</AccessRequest>
<?xml version="1.0"?>
<TrackRequest>
  <Request>
    <TransactionReference>
      <CustomerContext>Example</CustomerContext>
    </TransactionReference>
    <RequestAction>Track</RequestAction>
    <RequestOption>1</RequestOption>
  </Request>
  <TrackingNumber>{tracking_number}</TrackingNumber>
</TrackRequest>
"""

UPS_TRACK_URL = "https://onlinetools.ups.com/ups.app/xml/Track"

# Function to fetch records from Airtable
def fetch_records(search_value):
    all_records = []
    offset = None
    search_value_upper = search_value.upper()

    try:
        while True:
            params = {
                'filterByFormula': f"SEARCH('{search_value_upper}', UPPER({{TRACKING}}))",
                # Make sure TRACKING is a valid field
                'fields[]': ['TRACKING', 'SCAN', 'Created']  # Limit fields to reduce payload
            }

            # Print parameters for debugging
            print(f"Request Parameters: {params}")

            if offset:
                params['offset'] = offset

            # Send the request
            response = requests.get(AIRTABLE_URL, headers=headers, params=params)

            if response.status_code != 200:
                print(f"Failed to fetch data: {response.status_code}, {response.text}")
                raise ValueError(f"Airtable API Error: {response.status_code} - {response.text}")

            response.raise_for_status()  # Raise an error for bad responses

            data = response.json()

            if "records" not in data:
                raise ValueError("Response does not contain 'records'. Check your API key, base ID, or table name.")

            all_records.extend(data["records"])

            # Check if pagination is present in the response (pagination with offset)
            offset = data.get("offset")
            if not offset:
                break

        return all_records
    except requests.exceptions.RequestException as req_error:
        print(f"Network error: {req_error}")
        messagebox.showerror("Network Error", f"Failed to connect to Airtable: {req_error}")
        return []
    except ValueError as e:
        print(f"Value Error: {e}")
        messagebox.showerror("Error", f"Failed to retrieve records: {e}")
        return []
    except Exception as e:
        print(f"Unexpected error: {e}")
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        return []

# Function to get UPS tracking information
def fetch_ups_data(tracking_number):
    try:
        # Prepare the XML request
        xml_request = UPS_XML_TEMPLATE.format(tracking_number=tracking_number)

        # Send the request
        response = requests.post(UPS_TRACK_URL, data=xml_request, headers={"Content-Type": "application/xml"})
        response.raise_for_status()

        # Parse the XML response
        root = ET.fromstring(response.text)
        activity = root.find(".//Activity")
        description = activity.find("Status/StatusType/Description").text if activity is not None else "N/A"
        activity_date = activity.find("Date").text if activity is not None else "N/A"
        return activity_date, description
    except Exception as e:
        return "N/A", "N/A"

# Function to handle multiple search values asynchronously using ThreadPoolExecutor
def fetch_multiple_records_async(search_values):
    status_label.configure(text="Searching...")

    def wrapper():
        values = [value.strip() for value in search_values.split("\n") if value.strip()]
        all_records = []
        futures = []

        with ThreadPoolExecutor(max_workers=10) as executor:  # Adjust max_workers as needed
            for value in values:
                # Asynchronously fetch Airtable and UPS data
                futures.append(executor.submit(process_search_value, value))

            for future in as_completed(futures):
                result = future.result()
                if result:
                    all_records.extend(result)

        display_multiple_records(all_records)
        status_label.configure(text="Search completed.")

    thread = threading.Thread(target=wrapper)
    thread.start()

# Process a single search value: fetch both Airtable and UPS data
def process_search_value(value):
    airtable_records = fetch_records(value)
    ups_date, ups_scan = fetch_ups_data(value)

    for record in airtable_records:
        record["ups_date"] = ups_date
        record["ups_scan"] = ups_scan
        record["search_value"] = value

    if not airtable_records:
        # Append record even if Airtable is empty
        airtable_records.append({"search_value": value, "fields": {}, "createdTime": "N/A", "ups_date": ups_date, "ups_scan": ups_scan})

    return airtable_records

# Function to display records
def display_multiple_records(records):
    tree.delete(*tree.get_children())

    if not records:
        messagebox.showinfo("Search Results", "No matching records found in Airtable.")
        return

    displayed = set()
    global displayed_data
    displayed_data = []

    for record in records:
        search_value = record.get("search_value", "N/A")
        if search_value in displayed:
            continue

        fields = record.get("fields", {})
        scan = fields.get("SCAN", "N/A")
        created_time = record.get("createdTime", "N/A")
        ups_date = record.get("ups_date", "N/A")
        ups_scan = record.get("ups_scan", "N/A")

        # Handle the case where Airtable's created_time is in YYYYMMDD format
        if created_time != "N/A" and len(created_time) == 8:  # YYYYMMDD format
            created_time = f"{created_time[:4]}-{created_time[4:6]}-{created_time[6:]}"  # Convert to YYYY-MM-DD

        # Ensure the ups_date is in YYYY-MM-DD format
        if ups_date != "N/A" and len(ups_date) == 8:  # YYYYMMDD format
            ups_date = f"{ups_date[:4]}-{ups_date[4:6]}-{ups_date[6:]}"  # Convert to YYYY-MM-DD

        # Compare dates (created_time and ups_date) if they are not "N/A"
        most_recent_date = max(
            created_time[:10], ups_date, key=lambda d: datetime.strptime(d, "%Y-%m-%d") if d != "N/A" else datetime.min
        )

        # Insert the record into the Treeview
        tree.insert("", "end", values=(search_value, scan, created_time[:10], ups_date, ups_scan, most_recent_date))
        displayed_data.append({"TRACKING": search_value, "SCAN": scan, "Created": created_time[:10], "UPS Date": ups_date, "UPS_SCAN": ups_scan, "Most Recent": most_recent_date})
        displayed.add(search_value)

# Function to export data to Excel
def export_to_excel():
    if not displayed_data:
        messagebox.showinfo("Export Error", "No data to export.")
        return

    try:
        df = pd.DataFrame(displayed_data)
        file_path = "SearchResults.xlsx"
        df.to_excel(file_path, index=False)
        os.startfile(file_path)
    except Exception as e:
        messagebox.showerror("Export Error", f"Failed to export data to Excel: {e}")

# GUI setup (add a new column for UPS_SCAN)
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Airtable & UPS Search Interface")
root.geometry("800x600")

icon_image = Image.open(r"C:\Users\PC CABA DZ\Downloads\UPS-log.png")
icon_tk_image = ImageTk.PhotoImage(icon_image)
root.iconphoto(True, icon_tk_image)

input_frame = ctk.CTkFrame(root)
input_frame.pack(pady=10, padx=10, fill="x")

ctk.CTkLabel(input_frame, text="Search Values (one per line):").pack(pady=5)
search_text = ctk.CTkTextbox(input_frame, height=120)
search_text.pack(pady=10, padx=20)

buttons_frame = ctk.CTkFrame(root)
buttons_frame.pack(pady=5)

search_button = ctk.CTkButton(buttons_frame, text="Search", command=lambda: fetch_multiple_records_async(search_text.get("1.0", "end").strip()))
search_button.grid(row=0, column=0, padx=10)

export_button = ctk.CTkButton(buttons_frame, text="Export to Excel", command=export_to_excel)
export_button.grid(row=0, column=1, padx=10)

tree_frame = ctk.CTkFrame(root)
tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

tree = ttk.Treeview(tree_frame, columns=("TRACKING", "SCAN", "Created", "UPS Date", "UPS_SCAN"), show="headings", height=15)
tree.heading("TRACKING", text="TRACKING")
tree.heading("SCAN", text="AIRTABLE_SCAN")
tree.heading("Created", text="AIRTABLE_DATE")
tree.heading("UPS Date", text="UPS_DATE")
tree.heading("UPS_SCAN", text="UPS_SCAN")
tree.pack(fill="both", expand=True)

status_label = ctk.CTkLabel(root, text="", text_color="lightblue")
status_label.pack(pady=5)

displayed_data = []

root.mainloop()
