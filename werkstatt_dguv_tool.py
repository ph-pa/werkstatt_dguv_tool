import os
from PyPDF4 import PdfFileReader, PdfFileWriter
from pdf2image import convert_from_path
import pytesseract

# Prompt the user for scanned file
file_name = input("Enter the scan file name: ")

# Prompt the user for Ticketnummer
ticket_nummer = input("Please enter the Ticketnummer: ")

# Open the scanned PDF file
with open(file_name, 'rb') as file:
    pdf = PdfFileReader(file)

    # Iterate through each page
    for i in range(pdf.numPages):
        page = pdf.getPage(i)
        # Extract the image of the page
        images = convert_from_path(file_name, first_page=i+1, last_page=i+1)
        if images:
            # Use OCR to extract the text from the image
            text = pytesseract.image_to_string(images[0], lang="deu", config='--psm 12') #PSM 12 (Sparse text with OSD) seems to get the best results regarding serial number recognition.
            
            # Write the extracted text to a .txt file
            #with open('scanned-page-{}.txt'.format(i+1), 'w') as f:
                #f.write(text)
            
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
                new_file_name =  new_file_name + "_" + ticket_nummer
                # Create a new PDF file for each page
                output = PdfFileWriter()
                output.addPage(page)
                with open("{}.pdf".format(new_file_name), "wb") as outputStream:
                    output.write(outputStream)
                # Delete the text file
                #os.remove("{}.txt".format(new_file_name))
        else:
            print("No image data found on page", i+1)

print("Extraction, Scanning and Saving Done!")
