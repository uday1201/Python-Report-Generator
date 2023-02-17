###
# Command format : python(3) [scriptname].py [OutputfolderName] [ImageFolderPath]
###
import sys
import os
from PIL import Image
from io import BytesIO
from PyPDF2 import PdfWriter, PdfReader, PdfMerger
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch


# importing the command line arguments in the format
args = sys.argv
# Error if more than 3 args are passed
if len(args) > 3:
    raise Exception("Arguments should not be more than 3")

# printing tha args for debugging
print("The output folder name is : " + args[1])
print("Image Folder path : " + args[2])

# creating a dir for report
dirPath = os.getcwd()+'/ImageReports/'
if not os.path.exists(dirPath):
    # Create the directory
    os.mkdir(dirPath)
    print(f"Directory '{dirPath}' created successfully.")
else:
    print(f"Directory '{dirPath}' already exists.")

# Set the directory path for the images
image_folder = args[2]

# Create a new PDF file for the title page
title = 'Thermal Sensor Output Report'
date = datetime.now().strftime("%m/%d/%Y %H:%M:%S")
title_pdf = BytesIO()
c = canvas.Canvas(title_pdf)
c.setFont('Helvetica-Bold', 32)
c.drawCentredString(4.25*inch, 7*inch, title)
c.setFont('Helvetica', 12)
c.drawCentredString(4.25*inch, 6.5*inch, f'Date: {date}')
c.showPage()
c.save()

# Create a new PDF file to store the images
output_pdf = PdfMerger()

# Add the title page to the output PDF file
output_pdf.append(PdfReader(title_pdf))

# Loop through all the images in the folder
for image_file in os.listdir(image_folder):
    # Check if the file is an image file
    if image_file.startswith('Inspected') and (image_file.endswith('.png')):
        print(image_file)
        # Open the image file
        with Image.open(os.path.join(image_folder, image_file)) as image:
            # Convert the image to PDF format and add it to the output PDF file
            with BytesIO() as f:
                image.save(f, format='PDF')
                pdf_data = f.getvalue()
                output_pdf.append(PdfReader(BytesIO(pdf_data)))

# Write the output PDF file to the specified path
with open(dirPath+'thermal_data_report_'+ datetime.now().strftime("%m-%d-%Y_%H-%M-%S") +'.pdf', 'wb') as f:
    output_pdf.write(f)