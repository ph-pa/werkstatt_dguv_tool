import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from PyPDF4 import PdfFileReader, PdfFileWriter
from pdf2image import convert_from_path
import pytesseract

def open_file():
    global file_name
    file_name = filedialog.askopenfilename(initialdir = '/', title = "Select file", filetypes = (("PDF files", "*.pdf"), ("all files", "*.*")))
    file_path.set(file_name)
    
def save_to():
    global save_directory
    save_directory = filedialog.askdirectory(initialdir = '/', title = "Save To")
    save_to_label.configure(text = save_directory)

def extract_scan_save():
    global file_name
    global ticket_nummer
    
    # Open the scanned PDF file
    with open(file_path.get(), 'rb') as file:
        pdf = PdfFileReader(file)

        # Iterate through each page
        for i in range(pdf.numPages):
            page = pdf.getPage(i)
            # Extract the image of the page
            images = convert_from_path(file_path.get(), first_page=i+1, last_page=i+1)
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
                    # Use the next word as the new file name
                    new_file_name = words[index + 1]
                    # Add additional number to the file name
                    new_file_name =  new_file_name + "_" + ticket_nummer_entry.get()
                    # Create a new PDF file for each page
                    output = PdfFileWriter()
                    output.addPage(page)
                    with open("{}.pdf".format(save_directory, new_file_name), "wb") as outputStream:
                        output.write(outputStream)
            else:
                print("No image data found on page", i+1)
        print("Extraction, Scanning and Saving Done!")

root = tk.Tk()
root.title("Scan and Extract")

file_path = tk.StringVar()
ticket_nummer = tk.StringVar()

open_file_button = tk.Button(root, text = "Open File", command = open_file)
open_file_button.grid(row = 0, column = 0, padx = 10, pady = 10)
open_file_button.place(x=110, y=75, width=230, height=25)

file_path_label = tk.Label(root, textvariable = file_path)
file_path_label.grid(row = 0, column = 1, padx = 10, pady = 10)
#file_path_label.place(x=110, y=75, width=230, height=25)

save_to_button = tk.Button(root, text = "Save To", command = save_to)
save_to_button.grid(row = 3, column = 0, padx = 10, pady = 10)
#save_to_button.place(x=110, y=75, width=230, height=25)

save_to_label = tk.Label(root, text = "")
save_to_label.grid(row = 3, column = 1, padx = 10, pady = 10)
#save_to_label.place(x=110, y=75, width=230, height=25)

ticket_nummer_label = tk.Label(root, text = "Ticketnummer: ")
ticket_nummer_label.grid(row = 1, column = 0, padx = 10, pady = 10)
#ticket_nummer_label.place(x=110, y=75, width=230, height=25)

ticket_nummer_entry = tk.Entry(root)
ticket_nummer_entry.grid(row = 1, column = 1, padx = 10, pady = 10)
#ticket_nummer_entry.place(x=110, y=75, width=230, height=25)

extract_scan_save_button = tk.Button(root, text = "Extract, Scan and Save", command = extract_scan_save)
extract_scan_save_button.grid(row = 2, column = 0, columnspan = 2, padx = 10, pady = 10)
#extract_scan_save_button.place(x=110, y=75, width=230, height=25)

root.mainloop()