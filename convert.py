import os
import comtypes.client

# Set the directory to the current working directory
pptx_directory = os.getcwd()

# Start PowerPoint application
powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
powerpoint.Visible = 1

# Iterate through all files in the directory
for filename in os.listdir(pptx_directory):
    if filename.endswith(".pptx"):
        pptx_path = os.path.join(pptx_directory, filename)
        pdf_path = os.path.join(pptx_directory, f"{os.path.splitext(filename)[0]}.pdf")
        
        try:
            # Open the presentation in read-only mode
            presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False, ReadOnly=True)
            presentation.SaveAs(pdf_path, 32)  # 32 is the format code for PDF
            presentation.Close()
            print(f"Converted: {filename} to PDF successfully.")

        except comtypes.COMError as e:
            print(f"Failed to convert {filename} to PDF. Error: {e}")

# Quit PowerPoint application
powerpoint.Quit()
