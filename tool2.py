import pandas as pd
from fpdf import FPDF
import csv 

file_name = input("Enter the excel file name: ")

""" # Open the txt file using the csv module
with open('names.txt', newline='') as file:
    # Use the csv reader to read the file
    reader = csv.reader(file, delimiter=',')
    # Create an empty list to store the names
    names = []
    # Iterate through each row in the file
    for row in reader:
        # Append the first element in the row (assuming the name is the first element) to the names list
        names.append(row[0]) """
        
with open('names.txt', newline='') as file:
    # Use the csv reader to read the file
    reader = csv.reader(file, delimiter=',')
    # Create an empty list to store the names
    names = []
    # Iterate through each row in the file
    for row in reader:
        if len(row) > 1:
            # Append all elements in the row to the names list
            names.extend(row)
        else:
            # Append the first element in the row (assuming the name is the first element) to the names list
            names.append(row[0])

# You can now use the names list for further processing
print(names)

# Function to extract serial numbers from excel file
def extract_serial_numbers(file_name):
    # Read the excel file
    df = pd.read_excel(file_name)
    # Extract the serial numbers column
    #serial_numbers = df['Seriennummer'] # This extracts using column name
    serial_numbers = df.iloc[:,0]
    #serial_numbers = df.loc[:, 'Column Name'] # Specify exact row and column you want data from
    return serial_numbers

# Function to assign serial numbers to names
def assign_serial_numbers(serial_numbers, names):
    # Create a dictionary to store the name-serial number pairs
    name_serial_numbers = {}
    # Assign one serial number to each name
    for i in range(len(names)):
        name_serial_numbers[names[i]] = serial_numbers[i]
    return name_serial_numbers

# Function to insert data into pdf template
def insert_data_into_pdf(name_serial_numbers, template_path, output_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    # Insert the name-serial number pairs into the pdf template
    for name, serial_number in name_serial_numbers.items():
        pdf.cell(200, 10, txt="Name: {}".format(name), ln=1, align="L")
        pdf.cell(200, 10, txt="Serial Number: {}".format(serial_number), ln=1, align="L")
    pdf.output(output_path)

# Example usage
#file_path = "serial_numbers.xlsx"
#names = ["John Doe", "Jane Doe", "Bob Smith"]

serial_numbers = extract_serial_numbers(file_name)
name_serial_numbers = assign_serial_numbers(serial_numbers, names)
insert_data_into_pdf(name_serial_numbers, "template.pdf", "output.pdf")
