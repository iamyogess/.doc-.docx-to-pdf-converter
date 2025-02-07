import os
import docx2pdf
import pythoncom
import win32com.client as win32

# Input and output directories
input_dir = "C:/Users/iamyo/Desktop/attachments"
output_dir = "C:/Users/iamyo/Desktop/attachments/done"

# Ensure the output directory exists
os.makedirs(output_dir, exist_ok=True)

# Iterate through files
for filename in os.listdir(input_dir):
    input_path = os.path.join(input_dir, filename)
    output_file = filename.replace(os.path.splitext(filename)[1], ".pdf")
    output_path = os.path.join(output_dir, output_file)

    # Normalize paths
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    try:
        if filename.endswith(".docx"):
            # Convert .docx files (docx2pdf only accepts input & output directories)
            docx2pdf.convert(input_path, output_dir)
            print(f"✔ {filename} converted to PDF.")

        elif filename.endswith(".doc"):
            pythoncom.CoInitialize()
            word = win32.DispatchEx('Word.Application')
            word.Visible = False

            doc = word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17)  # Save as PDF
            doc.Close()
            print(f"✔ {filename} converted to PDF.")

    except Exception as e:
        print(f"❌ Failed to convert {filename}: {e}")

    finally:
        try:
            if 'word' in locals():  # Ensure Word is closed if initialized
                word.Quit()
        except Exception as quit_error:
            print(f"⚠ Failed to close Word for {filename}: {quit_error}")
