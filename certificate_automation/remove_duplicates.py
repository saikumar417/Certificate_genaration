import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import traceback
import datetime

def log_error(error_message):
    """Log errors to a file."""
    with open("error_log.txt", "a") as log_file:
        log_file.write(f"{datetime.datetime.now()} - {error_message}\n")


def main():
    try:
        Tk().withdraw()

        input_file = askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx")])
        if not input_file:
            print("No file selected.")
            return

        df = pd.read_excel(input_file)

        if 'Name' not in df.columns:
            print("The specified column 'Name' does not exist in the Excel file.")
            return

        df['Normalized_Name'] = df['Name'].str.strip().str.lower()

        df_duplicates = df[df.duplicated(subset='Normalized_Name', keep=False)]
        df_unique = df.drop_duplicates(subset='Normalized_Name', keep='first')

        with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_unique.to_excel(writer, sheet_name='Unique Rows', index=False)
            df_duplicates.to_excel(writer, sheet_name='Removed Rows', index=False)

        print(f"Processed file saved at: {input_file}")
        print("Removed rows have been saved in the 'Removed Rows' sheet, and unique rows in 'Unique Rows' sheet.")
    except Exception as e:
        error_message = traceback.format_exc()
        log_error(error_message)
        print("An error occurred. Check error_log.txt for more details.")


if __name__ == "__main__":
    main()
