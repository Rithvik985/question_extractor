import os
import subprocess

# Set these paths
libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
input_dir = r"C:\Users\chitr\Desktop\qpextractor"
output_dir = r"C:\Users\chitr\Desktop\qpextractor\test_files"

def convert_docx_to_pdf(input_dir, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for file in os.listdir(input_dir):
        if file.endswith(".doc"):
            input_path = os.path.join(input_dir, file)


            # Call LibreOffice to convert
            result = subprocess.run([
                libreoffice_path,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                input_path
            ], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

            print(f"Converting: {file}")
            if result.returncode == 0:
                print(f"✅ Success: {file}")
            else:
                print(f"❌ Failed: {file}")
                print(result.stderr)

convert_docx_to_pdf(input_dir, output_dir)
