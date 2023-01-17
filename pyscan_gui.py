import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from PyPDF4 import PdfFileReader, PdfFileWriter
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import pytesseract
import threading

def open_file():
    global file_name
    file_name = filedialog.askopenfilename(initialdir = '/', title = "Select file", filetypes = (("PDF files", "*.pdf"), ("all files", "*.*")))
    file_path.set(file_name)
    if len(file_name) > 35:
        file_path_label.configure(text = "..." + file_name[-32:])
    else:
        file_path_label.configure(text = file_name)
    
def save_to():
    global save_directory
    save_directory = filedialog.askdirectory(initialdir = '/', title = "Save To")
    if len(save_directory) > 35:
        save_to_label.configure(text = "..." + save_directory[-32:])
    else:
        save_to_label.configure(text = save_directory)

def extract_scan_save():
    global file_name
    global save_directory
    # Change the text of the button to show the current status
    if save_directory == '':
        status_label.configure(text = "Please specify the directory to save the files.")
        return
    extract_scan_save_button.configure(text = "Processing...")
    # Open the scanned PDF file
    with open(file_path.get(), 'rb') as file:
        pdf = PdfFileReader(file)

        # Iterate through each page
        for i in range(pdf.numPages):
            page = pdf.getPage(i)
            status_label.configure(text = "Processing page {} of {}".format(i+1, pdf.numPages))
            # Extract the image of the page
            images = convert_from_path(file_name, first_page=i+1, last_page=i+1)
            if images:
                # Use OCR to extract the text from the image
                text = pytesseract.image_to_string(images[0], lang="deu", config='--psm 12')
                # Split the text by whitespace
                words = text.split()

                # Find the index of the word "Seriennummer"
                try:
                    index = words.index("(Seriennummer):")
                except ValueError:
                    index = -1

                if index != -1:
                    # Use the next word
                    new_file_name = words[index + 1]
                    # Add additional number to the file name
                    new_file_name =  new_file_name + "_" + ticket_nummer_entry.get()
                    # Create a new PDF file for each page
                    output = PdfFileWriter()
                    output.addPage(page)
                    with open("{}/{}.pdf".format(save_directory, new_file_name), "wb") as outputStream:
                        output.write(outputStream)
                else:
                    print("No image data found on page", i+1)
    # Change the text of the button back to the original text
    extract_scan_save_button.configure(text = "Extract, Scan and Save")
    #Update the status label to show that the process has completed
    status_label.configure(text = "Done!")
    print("Extraction, Scanning and Saving Done!")

def on_button_click():
    thread = threading.Thread(target=extract_scan_save)
    thread.start()
 
root = tk.Tk()
root.resizable(False, False)
root.geometry("420x240")
root.title("DGUV Werkstatt Extract Tool")

# Bechtle icon in the upper left corner
root.wm_iconbitmap("C:/Users/PhilippPavelic/Documents/Code/pyscan/bechtle.ico")
# Adding Bechtle Logo
icon = Image.open("C:/Users/PhilippPavelic/Documents/Code/pyscan/bechtle.ico")
icon = icon.resize((64, 64))
icon = ImageTk.PhotoImage(icon)

icon_label = tk.Label(root, image=icon)
#icon_label.grid(row=0, column=3)
icon_label.place(x=340, y=0)

file_path = tk.StringVar()

open_file_button = tk.Button(root, text = "Open File", command = open_file)
#open_file_button.grid(row = 0, column = 0, padx = 10, pady = 10)
open_file_button.place(x=25, y=25, width=65, height=25)

file_path_label = tk.Label(root, text = "")
#file_path_label.grid(row = 0, column = 1, padx = 10, pady = 10)
file_path_label.place(x=110, y=27, width=220, height=25)

save_to_button = tk.Button(root, text = "Save To", command = save_to)
#save_to_button.grid(row = 1, column = 0, padx = 10, pady = 10)
save_to_button.place(x=25, y=75, width=65, height=25)

save_to_label = tk.Label(root, text = "")
#save_to_label.grid(row = 1, column = 1, padx = 10, pady = 10)
save_to_label.place(x=110, y=77, width=220, height=25)

ticket_nummer_label = tk.Label(root, text = "Ticketnummer: ")
#ticket_nummer_label.grid(row = 2, column = 0, padx = 10, pady = 10)
ticket_nummer_label.place(x=25, y=125, width=80, height=25)

ticket_nummer_entry = tk.Entry(root)
#ticket_nummer_entry.grid(row = 2, column = 1, columnspan = 2, padx = 10, pady = 10)
ticket_nummer_entry.place(x=185, y=125, width=150, height=25)

extract_scan_save_button = tk.Button(root, text = "Extract, Scan and Save", command = on_button_click)
#extract_scan_save_button.grid(row = 3, column = 0, columnspan = 3, padx = 10, pady = 10)
#extract_scan_save_button.place(x=145, y=180, width=130, height=25)
extract_scan_save_button.place(x=145, y=180, width=130, height=25)

status_label = tk.Label(root, text = "")
status_label.grid(row = 4, column = 0, columnspan = 2, padx = 10, pady = 10)
#status_label.place(x=110, y=75, width=230, height=25)

root.mainloop()