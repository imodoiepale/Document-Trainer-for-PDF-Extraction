import tkinter as tk
from tkinter import filedialog, Text, Entry, Label, Button, Canvas, simpledialog
import json
import fitz
from PIL import Image, ImageTk

root = tk.Tk()
root.geometry("800x1080")  # Set the size of the window

sidebar = tk.Frame(root, width=200, bg='white', height=500, relief='sunken', borderwidth=2)
sidebar.pack(expand=False, fill='both', side='left', anchor='nw')

# Create a canvas to draw on
canvas = Canvas(root, width=800, height=1080)
canvas.pack()

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        doc = fitz.open(file_path)
        page = doc.load_page(0)  # load first page
        pix = page.get_pixmap()  # render page to an image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        photo = ImageTk.PhotoImage(image=img)
        canvas.create_image(0, 0, image=photo, anchor='nw')
        canvas.image = photo  # keep a reference to the image

        # Prompt for a filename
        filename = simpledialog.askstring("Filename", "Enter a filename for the template")
        # Initialize saved_positions for this file
        root.saved_positions = {}
        # Save the filename
        root.filename = filename

def start_position(event):
    # Start drawing the rectangle
    canvas.start_x = event.x
    canvas.start_y = event.y
    canvas.curr_rectangle = canvas.create_rectangle(canvas.start_x, canvas.start_y, canvas.start_x, canvas.start_y, outline='red')

def save_position(event):
    # Update the rectangle as the mouse moves
    x, y = event.x, event.y
    canvas.coords(canvas.curr_rectangle, canvas.start_x, canvas.start_y, x, y)

def end_position(event):
    # Finalize the rectangle when the mouse button is released
    x, y = event.x, event.y
    
    # Prompt for a label
    name = simpledialog.askstring("Label", "Enter a label for this region")
    
    # Define the region as a rectangle around the clicked point
    region = {'x1': canvas.start_x, 'y1': canvas.start_y, 'x2': x, 'y2': y}
    
    # Save the region and its label
    root.saved_positions[name] = region
    
    # Add a label for the selected position
    canvas.create_text(x, y, text=name, fill='blue')
    
    # Save all positions with each label on a new line
    with open(f'{root.filename}.txt', 'w') as f:
        json.dump(root.saved_positions, f, indent=4)

# Add the "Open PDF File" button
open_button = Button(sidebar, text="Open PDF File", command=open_file)
open_button.pack()

canvas.bind("<Button-1>", start_position)
canvas.bind("<B1-Motion>", save_position)
canvas.bind("<ButtonRelease-1>", end_position)

root.mainloop()
