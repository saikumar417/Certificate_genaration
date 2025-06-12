Certificate Generation Script
Author: N Sai Kumar
Registration No: 23751F0020
Branch: MCA 2023-25 Batch
Guidance: Mr. J. Sheik Mohamed sir CO-ORDINATER, SITAMS.
Note: Taken help from OpenAI's GPT-4 architecture.

Description:
This script generates certificates using data from an Excel file and adds various elements like logos, participant details, and signatures to the certificates. The generated certificates are saved as PNG files, and all certificates are compiled into a single PDF (in landscape orientation) and a PowerPoint file.

Requirements:
Installations:

Pillow - For image manipulation (e.g., adding text to certificates, resizing logos).
bash
Copy code
pip install Pillow

pandas - For reading Excel files.
bash
Copy code
pip install pandas

tkinter - For file dialog operations (comes pre-installed with Python). If missing, you can install using the system's package manager.
python-pptx - For generating PowerPoint presentations.

bash
Copy code
pip install python-pptx
reportlab - For generating PDFs.

bash
Copy code
pip install reportlab
openpyxl - For reading Excel files (xlsx format).

bash
Copy code
pip install openpyxl

Additional Dependencies:
arial.ttf font (ensure you have this font installed or replace it with any font you prefer in the script).

How to Run:
Clone or download this project to your local machine.

Place your certificate template (certificate.png), logos, and other required input files in the input_files folder.

Run the script by executing the following command:

bash
Copy code
python <your_script_name>.py
Select your Excel file when prompted by the file dialog.

The script will generate individual certificate PNGs, a single PDF file (in landscape), and a PowerPoint presentation.

The output will be saved in the certificates_output folder.

Output:
Each certificate is saved as an individual PNG.
All certificates are combined into a single landscape PDF (All_Certificates_Landscape.pdf).
All certificates are saved into a PowerPoint file (All_Certificates.pptx).