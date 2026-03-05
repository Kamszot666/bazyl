# diagnostic_script.py
import pandas as pd

def check_excel_structure(file_path):
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Display the columns and their data types
        print("Columns in the Excel file:")
        print(df.columns.tolist())
        print("\nData types of each column:")
        print(df.dtypes)

        # Check for missing values
        print("\nMissing values in each column:")
        print(df.isnull().sum())

    except Exception as e:
        print("An error occurred:", e)

if __name__ == "__main__":
    # Example file path (adjust as necessary)
    check_excel_structure('example.xlsx')
