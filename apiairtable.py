import customtkinter as ctk
import requests
import threading
import pandas as pd
import os
from tkinter import ttk
from PIL import Image, ImageTk  # Import ImageTk to create Tkinter-compatible images
from datetime import datetime

# Airtable Configuration
API_KEY = "patJNieyOmCdDcA51.635e3b384caca979ac0666c13cd2516185d719ecb988abfd939309f793753d56"
AIRTABLE_URL = f"https://api.airtable.com/v0/appoGD1yWNzgOPH0O/TPD"

headers = {"Authorization": f"Bearer {API_KEY}"}

# Function to fetch records based on filter
def fetch_records(search_value):
    all_records = []
    offset = None

    # Convert the search value to uppercase
    search_value_upper = search_value.upper()

    try:
        while True:
            # Use Airtable's UPPER formula for case-insensitive search
            params = {
                'filterByFormula': f"SEARCH('{search_value_upper}', UPPER({{TRACKING}}))",
            }
            if offset:
                params['offset'] = offset

            response = requests.get(AIRTABLE_URL, headers=headers, params=params)
            response.raise_for_status()

            data = response.json()
            if "records" not in data:
                raise ValueError("Response does not contain 'records'. Check your API key, base ID, or table name.")

            all_records.extend(data["records"])

            offset = data.get("offset")
            if not offset:
                break

        return all_records
    except requests.exceptions.RequestException as req_error:
        messagebox.showerror("Network Error", f"Failed to connect to Airtable: {req_error}")
        return []
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        return []


# Function to handle multiple search values asynchronously
def fetch_multiple_records_async(search_values):
    # Update status to indicate searching
    status_label.configure(text="Searching...")

    def wrapper():
        # Split search values into a list, ignoring empty lines
        values = [value.strip() for value in search_values.split("\n") if value.strip()]
        all_records = []

        for value in values:
            records = fetch_records(value)
            if records:
                # Add the search value to each record for display
                for record in records:
                    record["search_value"] = value
                all_records.extend(records)
            else:
                # Add a placeholder if no records are found for this value
                all_records.append({"search_value": value, "fields": {}, "createdTime": "N/A"})

        display_multiple_records(all_records)
        # Update status to indicate completion
        status_label.configure(text="Search completed.")

    # Create a new thread to fetch records and display them
    thread = threading.Thread(target=wrapper)
    thread.start()


# Function to display multiple records in the Treeview
def display_multiple_records(records):
    # Clear the Treeview
    tree.delete(*tree.get_children())

    if not records:
        ctk.CTkMessagebox(title="Search Results", message="No matching records found in Airtable.")
        return

    # Display the most recent record for each searched value
    displayed = set()  # Keep track of displayed search values to avoid duplicates
    global displayed_data
    displayed_data = []  # Prepare data for exporting to Excel

    for record in records:
        search_value = record.get("search_value", "N/A")
        if search_value in displayed:
            continue  # Skip duplicates

        fields = record.get("fields", {})
        scan = fields.get("SCAN", "N/A")
        created_time = record.get("createdTime", "N/A")

        # Format the Created date (just the date part)
        created_date = created_time[:10] if created_time != "N/A" else "N/A"

        # Insert the record into the Treeview
        tree.insert("", "end", values=(search_value, scan, created_date))
        displayed_data.append({"TRACKING": search_value, "SCAN": scan, "Created": created_date})
        displayed.add(search_value)


# Function to export the displayed data to Excel
def export_to_excel():
    if not displayed_data:
        ctk.CTkMessagebox(title="Export Error", message="No data to export.")
        return

    try:
        # Create a DataFrame and export to Excel
        df = pd.DataFrame(displayed_data)
        file_path = "SearchResults.xlsx"
        df.to_excel(file_path, index=False)

        # Open the Excel file
        os.startfile(file_path)
    except Exception as e:
        ctk.CTkMessagebox(title="Export Error", message=f"Failed to export data to Excel: {e}")


# CustomTkinter GUI
ctk.set_appearance_mode("dark")  # Set the theme to dark
ctk.set_default_color_theme("blue")  # Set the color theme to blue

root = ctk.CTk()
root.title("Airtable Search Interface")
root.geometry("600x550")

# Load the .png image using Pillow
icon_image = Image.open(r"C:\Users\PC CABA DZ\Downloads\UPS-log.png")

# Convert the image to a format that Tkinter can use (PhotoImage)
icon_tk_image = ImageTk.PhotoImage(icon_image)

# Set the window icon using the PhotoImage object
root.iconphoto(True, icon_tk_image)

# Input Box Frame (Centered on Top)
input_frame = ctk.CTkFrame(root)
input_frame.pack(pady=10, padx=10, fill="x")

ctk.CTkLabel(input_frame, text="Search Values (one per line):").pack(pady=5)
search_text = ctk.CTkTextbox(input_frame, height=120)
search_text.pack(pady=10, padx=20)

# Buttons Frame (Below the Input Box)
buttons_frame = ctk.CTkFrame(root)
buttons_frame.pack(pady=5)

search_button = ctk.CTkButton(buttons_frame, text="Search", command=lambda: fetch_multiple_records_async(search_text.get("1.0", "end").strip()))
search_button.grid(row=0, column=0, padx=10)

export_button = ctk.CTkButton(buttons_frame, text="Export to Excel", command=export_to_excel)
export_button.grid(row=0, column=1, padx=10)

# Treeview for displaying TRACKING, SCAN, and Created fields
tree_frame = ctk.CTkFrame(root)
tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

tree = ttk.Treeview(tree_frame, columns=("TRACKING", "SCAN", "Created"), show="headings", height=15)
tree.heading("TRACKING", text="TRACKING")
tree.heading("SCAN", text="SCAN")
tree.heading("Created", text="Created")
tree.pack(fill="both", expand=True)

# Status Label
status_label = ctk.CTkLabel(root, text="", text_color="lightblue")
status_label.pack(pady=5)

# Initialize global variable for data
displayed_data = []

root.mainloop()