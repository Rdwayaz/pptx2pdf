import os
import comtypes.client

pptx_directory = os.getcwd()

powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
powerpoint.Visible = 1

for filename in os.listdir(pptx_directory):
    if filename.endswith(".pptx"):
        pptx_path = os.path.join(pptx_directory, filename)
        pdf_path = os.path.join(pptx_directory, f"{os.path.splitext(filename)[0]}.pdf")
        
        try:
            presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False, ReadOnly=True)
            presentation.SaveAs(pdf_path, 32)  # 32 is the format code for PDF
            presentation.Close()
            print(f"Converted: {filename} to PDF successfully.")

        except comtypes.COMError as e:
            print(f"Failed to convert {filename} to PDF. Error: {e}")

powerpoint.Quit()
