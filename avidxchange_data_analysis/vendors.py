import pandas as pd

# Load the CSV file into a DataFrame
file_path = '/Users/harrisonweiss/Downloads/Corporate Facilities.xlsx'  # Replace with the path to your CSV file
data = pd.read_csv(file_path)

# Get the unique vendor names from the 'Vendor Name' column
unique_vendors = data['Vendor Name'].unique()

# Print the list of unique vendor names
for vendor in unique_vendors:
    print(vendor)
