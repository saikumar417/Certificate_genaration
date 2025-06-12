import fitz  # PyMuPDF
import shutil
import os

def protect_pdf(input_pdf_path, output_pdf_path):
    # Open the existing PDF
    doc = fitz.open(input_pdf_path)

    # Define a temporary output path
    temp_output_path = output_pdf_path + ".temp.pdf"

    # Save to a temporary file to avoid incremental save issues
    doc.save(
        temp_output_path,
        encryption=fitz.PDF_ENCRYPT_AES_256,  # Apply AES-256 encryption
        owner_pw="",                          # Empty owner password to limit edit access
        incremental=False                     # Ensure the save is not incremental
    )
    doc.close()  # Close the document after saving

    # Move the temporary file to the final output path
    shutil.move(temp_output_path, output_pdf_path)

    print(f"Protected PDF saved to {output_pdf_path}")

# Example usage:
# input_pdf = "your_certificate.pdf"
# output_pdf = "protected_certificate.pdf"
# protect_pdf(input_pdf, output_pdf)
