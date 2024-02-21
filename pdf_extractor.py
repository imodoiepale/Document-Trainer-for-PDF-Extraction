import tkinter as tk
from tkinter import filedialog, OptionMenu
import pdfplumber
import json
import openpyxl
import os
import re

def load_positions(filename):
    # Load the positions from the selected file
    with open(filename, 'r') as f:
        positions = json.load(f)

    # Convert the coordinates to integers
    for label, position in positions.items():
        for key in position:
            position[key] = int(position[key])

    return positions



def extract_text():
    positions = load_positions(root.filename.get())

    filenames = filedialog.askopenfilenames(
        title="Select PDF Files",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if filenames: 
        # Create a new workbook and select the active sheet
        wb = openpyxl.Workbook()
        sheet = wb.active

        # Write the labels to the first row
        sheet.append(list(positions.keys()))

        for file in filenames:
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    row = []
                    for label, position in positions.items():
                        # Crop the page to the position
                        cropped_page = page.crop((position['x1'], position['y1'], position['x2'], position['y2']))
                        # Extract the text from the cropped page
                        text = cropped_page.extract_text()
                        # Remove commas from the text
                        text = text.replace(",", "")
                        # Check if the text is numerical
                        if re.match(r'^-?\d+(?:\.\d+)?$', text):
                            text = float(text)
                        # Add the extracted text to the row
                        row.append(text)
                    # Write the row to the sheet
                    sheet.append(row)

        # Save the workbook to an Excel file
        wb.save("extracted_text.xlsx")

root = tk.Tk()
root.title("PDF Text Extractor")
root.geometry("400x400")

# Get a list of all text files in the current directory
txt_files = [f for f in os.listdir() if f.endswith('.txt')]

# Create a StringVar to hold the selected filename
root.filename = tk.StringVar(root)
root.filename.set(txt_files[0])  # Set the default value to the first file

# Create the dropdown menu
dropdown = OptionMenu(root, root.filename, *txt_files)
dropdown.pack()

extract_button = tk.Button(root, text="Select Files and Extract Text", command=extract_text)
extract_button.pack()

root.mainloop()
