import tkinter as tk
from tkinter import filedialog, OptionMenu, ttk
import pdfplumber
import openpyxl
import os
import re
import json
import subprocess
from datetime import datetime

def load_positions(filename):
    # Load the positions from the selected file
    positions = {}
    with open(filename, 'r') as f:
        positions = json.load(f)
    return positions

def extract_text():
    positions = load_positions(root.filename.get())
    filenames = filedialog.askopenfilenames(
        title="Select PDF Files",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if filenames:
        total_files = len(filenames)
        progress['maximum'] = total_files
        progress['value'] = 0
        progress.pack(side=tk.BOTTOM, fill=tk.X)
        root.update_idletasks()

        # Get the selected template name
        template_name = os.path.splitext(os.path.basename(root.filename.get()))[0]

        # Get current date and time
        current_datetime = datetime.now().strftime("%H.%M-%d.%m.%Y")

        # Create a new workbook and select the active sheet
        wb = openpyxl.Workbook()
        sheet = wb.active

        # Define column headers for the output Excel file
        headers = [
            "Taxpayer Name", "Email Address", "Certificate Date", "PIN", "L.R. Number", "Building",
            "Street/Road", "City/Town", "County", "District", "Tax Area", "Station", "P. O. Box",
            "Postal Code", "Income Tax - PAYE", "Income Tax - PAYE Effective frm Date",
            "Income Tax - PAYE Effective Till Date", "Income Tax - PAYE Status",
            "Value Added Tax (VAT)", "Value Added Tax (VAT) Effective from Date",
            "Value Added Tax (VAT) Effective till Date", "Value Added Tax (VAT) Status",
            "Income Tax - Company", "Income Tax - Company Effective from Date",
            "Income Tax - Company Effective till Date", "Income Tax - Company Status"
        ]

        # Write the headers to the first row
        sheet.append(headers)

        for index, file in enumerate(filenames, start=1):
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    row = []
                    for label, position in positions.items():
                        # Crop the page to the position
                        cropped_page = page.crop((position['x1'], position['y1'], position['x2'], position['y2']))
                        # Extract the text from the cropped page
                        text = cropped_page.extract_text().strip()  # Remove leading/trailing spaces
                        # Remove commas from the text
                        text = text.replace(",", "")
                        # Check if the text is numerical
                        if re.match(r'^-?\d+(?:\.\d+)?$', text):
                            text = float(text)
                        # Add the extracted text to the row
                        row.append(text)

                    # Organize the data based on the first column
                    organized_row = [
                        row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9],
                        row[10], row[11], row[12], row[13], "", "", "", "", "", "", "", "", "", "", "", ""
                    ]

                    # Distribute data into respective columns based on the first group
                    if len(row) > 14:
                        if row[14] == "Income Tax - PAYE":
                            organized_row[14:18] = row[14:18]
                        elif row[14] == "Value Added Tax (VAT)":
                            organized_row[18:22] = row[14:18]
                        elif row[14] == "Income Tax - Company":
                            organized_row[22:26] = row[14:18]

                    if len(row) > 18:
                        if row[18] == "Income Tax - PAYE":
                            organized_row[14:18] = row[18:22]
                        elif row[18] == "Value Added Tax (VAT)":
                            organized_row[18:22] = row[18:22]
                        elif row[18] == "Income Tax - Company":
                            organized_row[22:26] = row[18:22]

                    if len(row) > 22:
                        if row[22] == "Income Tax - PAYE":
                            organized_row[14:18] = row[22:26]
                        elif row[22] == "Value Added Tax (VAT)":
                            organized_row[18:22] = row[22:26]
                        elif row[22] == "Income Tax - Company":
                            organized_row[22:26] = row[22:26]

                    # Write the organized row to the sheet
                    sheet.append(organized_row)
            progress['value'] = index
            root.update_idletasks()

        # Save the workbook to an Excel file with the template name and current date and time
        file_path = f"{template_name} Extraction {current_datetime}.xlsx"
        wb.save(file_path)
        progress.pack_forget()
        complete_label.pack()
        file_label['text'] = "Open Excel file: " + os.path.abspath(file_path)
        file_label['command'] = lambda: open_file_explorer(template_name, current_datetime)
        file_label.pack()

def open_file_explorer(template_name, current_datetime):
    file_path = f"{template_name} Extraction {current_datetime}.xlsx"
    subprocess.Popen('explorer /select,"{}"'.format(os.path.abspath(file_path)))

root = tk.Tk()
root.title("PDF Text Extractor")
root.geometry("600x400")

# Get a list of all text files in the current directory
txt_files = [f for f in os.listdir() if f.endswith('.txt')]

if txt_files:
    # Create a StringVar to hold the selected filename
    root.filename = tk.StringVar(root)
    root.filename.set(txt_files[0])  # Set the default value to the first file

    # Create the dropdown menu
    dropdown = OptionMenu(root, root.filename, *txt_files)
    dropdown.pack()

    extract_button = tk.Button(root, text="Select Files and Extract Text", command=extract_text)
    extract_button.pack()
else:
    # Display a message if no text files are found
    label = tk.Label(root, text="No TXT files found in the directory.")
    label.pack()

progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=200, mode='determinate')
complete_label = tk.Label(root, text="Extraction complete", fg="green")
file_label = tk.Button(root, text="Open Excel file: ", fg="blue", cursor="hand2", command=lambda: open_file_explorer(None, None))

root.mainloop()
