import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory
import os

# Function to standardize date format
def standardize_date_format(df, date_column):
    df[date_column] = pd.to_datetime(df[date_column], errors='coerce')  # Convert to datetime
    df[date_column] = df[date_column].dt.strftime('%d-%m-%Y')  # Format as dd-mm-yyyy
    return df

# Function to convert names to UPPERCASE
def capitalize_names(df, name_column):
    df[name_column] = df[name_column].str.upper()  # Convert to UPPERCASE
    return df

# Function to remove duplicate rows based on "Name" column
def remove_duplicates(df, name_column):
    return df.drop_duplicates(subset=[name_column])

# Main script
def process_excel_file():
    Tk().withdraw()  # Hides the root Tkinter window
    print("Please select the Excel file...")

    input_file = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not input_file:
        print("No file selected.")
        return

    print(f"\nSelected file: {input_file}")

    # Prompt user for output directory
    print("Select the folder where the files should be saved...")
    output_dir = askdirectory()
    if not output_dir:
        print("No output directory selected. Exiting.")
        return
    
    print(f"Files will be saved in: {output_dir}\n")

    # Load the Excel file
    df = pd.read_excel(input_file)

    # Check for required columns
    required_columns = ['Department', 'Name', 'Date']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        print(f"Error: Missing columns: {', '.join(missing_columns)}")
        return

    # Standardize date format in 'Date' column
    df = standardize_date_format(df, 'Date')

    # Convert names to UPPERCASE
    df = capitalize_names(df, 'Name')

    # Remove duplicate rows based on 'Name' column
    df = remove_duplicates(df, 'Name')

    # Get unique department names
    departments = df['Department'].dropna().unique()

    # Separate data for each department
    for dept in departments:
        dept_data = df[df['Department'] == dept]
        output_file = os.path.join(output_dir, f'{dept}_data.xlsx')
        dept_data.to_excel(output_file, index=False)
        print(f"✅ Data for '{dept}' saved to: {output_file}")

    print("\n🎉 Processing complete!")

def main():  # New wrapper function for the process
    process_excel_file()

if __name__ == "__main__":
    main()  # This still allows direct running

