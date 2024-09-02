import os
import docx2pdf
import pythoncom
import win32com.client as win32

# Input and output directories
input_dir = "C:/Users/iamyo/Desktop/attachments"
output_dir = "C:/Users/iamyo/Desktop/attachments/done"

# Make sure the output directory exists
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Iterate through files
for filename in os.listdir(input_dir):
    input_path = os.path.join(input_dir, filename)
    output_path = os.path.join(output_dir, filename.replace(os.path.splitext(filename)[1], ".pdf"))

    # Normalize paths
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    if filename.endswith(".docx"):  # Handle .docx files
        try:
            docx2pdf.convert(input_path, output_path)
            print(f"{filename} converted to PDF.")
        except Exception as e:
            print(f"Failed to convert {filename}: {e}")

    elif filename.endswith(".doc"):  # Handle .doc files
        pythoncom.CoInitialize()  # Initialize the COM library
        try:
            word = win32.DispatchEx('Word.Application')  # Use DispatchEx to get a new Word instance each time
            word.Visible = False

            # Open the .doc file and save as PDF
            doc = word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17)  # Save as PDF
            doc.Close()
            print(f"{filename} converted to PDF.")
        except Exception as e:
            print(f"Failed to convert {filename}: {e}")
        finally:
            try:
                word.Quit()  # Ensure Word is closed even if an error occurs
            except Exception as quit_error:
                print(f"Failed to close Word for {filename}: {quit_error}")
