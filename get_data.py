import sys
from datetime import datetime, timedelta

import pandas as pd


# Function to process the Excel file
def process_excel(file_path):
    
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Ensure the 'arrivalDateTime' column is present
        if 'arrivalDateTime' not in df.columns:
            print("The 'arrivalDateTime' column is missing from the file.")
            return

        # Convert 'arrivalDateTime' column from text to datetime
        df['arrivalDateTime'] = pd.to_datetime(df['arrivalDateTime'], format='%B %d, %Y, %I:%M %p', errors='coerce')
        if df['arrivalDateTime'].isnull().any():
            print("Some 'arrivalDateTime' values could not be parsed. Check your data.")

        # Extract Month and Time
        df['Month'] = df['arrivalDateTime'].dt.strftime('%B')
        df['Time'] = df['arrivalDateTime'].dt.time

        # Sort by 'arrivalDateTime'
        df = df.sort_values(by='arrivalDateTime')

        # Calculate time difference in minutes between consecutive rows
        df['Time Difference (mins)'] = df['arrivalDateTime'].diff().dt.total_seconds().div(60).fillna(0)

        # Mark rows with a time difference <= 5 minutes
        df['Within 5 Min'] = df['Time Difference (mins)'] <= 5

        # Save the processed data to a new Excel file
        output_file = "processed_" + file_path
        df.to_excel(output_file, index=False)

        print(f"Processing complete. Data saved to {output_file}")

    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Entry point of the script
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 getdata.py <excelfile>")
    else:
        excel_file = sys.argv[1]
        process_excel(excel_file)
