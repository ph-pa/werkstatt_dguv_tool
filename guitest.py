import os
from tkinter import *
from tkinter import filedialog
from PyPDF4 import PdfFileReader, PdfFileWriter
from pdf2image import convert_from_path
import pytesseract

def browse_file():
    global file_path
    file_path = filedialog.askopenfilename()
    file_label.config(text=file_path)

def extract_and_save():
    # Open the scanned PDF file
    with open(file_path, 'rb') as file:
        pdf = PdfFileReader(file)

        # Iterate through each page
        for i in range(pdf.numPages):
            page = pdf.getPage(i)
            # Extract the image of the page
            images = convert_from_path(file_path, first_page=i+1, last_page=i+1)
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
                    new_file_name =  new_file_name + "_" + ticket_nummer.get()
                    # Create a new PDF file for each page
                    output = PdfFileWriter()
                    output.addPage(page)
                    with open("{}.pdf".format(new_file_name), "wb") as outputStream:
                        output.write(outputStream)
            else:
                print("No image data found on page", i+1)
        print("Extraction, Scanning and Saving Done!")

root = Tk()
root.title("PDF Extractor")

file_label = Label(root, text="No file selected")
file_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

browse_button = Button(root, text="Browse", command=browse_file)
browse_button.grid(row=1, column=0, padx=10, pady=10)

ticket_nummer = Entry(root)
ticket_nummer.grid(row=1, column=1, padx=10, pady=10)

extract_button = Button(root, text="Extract and Save", command=extract_and_save)
extract_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

root.mainloop