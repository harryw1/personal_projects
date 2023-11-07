import pandas as pd
from prophet import Prophet
import tkinter as tk
from tkinter import filedialog

# Function to prompt for file selection
def ask_for_file():
    # Create a root window, but don't display it
    root = tk.Tk()
    root.withdraw()  # We don't want a full GUI, so keep the root window from appearing
    
    # Show an "Open" dialog box and return the path to the selected file
    file_path = filedialog.askopenfilename(
        title="Select the Excel file",
        # filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))  # Optional: specify file types
    )
    
    return file_path

# Use the function to get the file path
excel_file_path = ask_for_file()

# Data Preprocessing
# Read the entire Excel sheet
df_full = pd.read_excel(excel_file_path)

# Find the first row index where all columns have non-NaN values
start_row = None
for i in range(len(df_full)):
    # Assuming the actual data has fewer NaNs than the header
    if df_full.iloc[i].notnull().sum() > 10:  # 'threshold' is how many non-NaNs you expect
        start_row = i
        break

# Now read the Excel file starting from the detected row and set header row to the first row
df = pd.read_excel(excel_file_path, skiprows=start_row, header=1)
df.columns = df.columns.str.strip()

# Convert 'Revenue Month' to datetime
df['Revenue Month'] = pd.to_datetime(df['Revenue Month'])

# Select relevant columns and rename them to 'ds' and 'y'
df_prophet = df[['Revenue Month', 'Total Cost']].rename(columns={'Revenue Month': 'ds', 'Total Cost': 'y'})

# Drop rows with missing values or fill them
df_prophet = df_prophet.dropna()

# Modeling with Prophet
model = Prophet(daily_seasonality=False, weekly_seasonality=False, yearly_seasonality=True)
model.fit(df_prophet)

# Make future predictions
future = model.make_future_dataframe(periods=365, freq='D')
forecast = model.predict(future)

# Save file features
def ask_save_as_file():
    # Create a root window, but don't display it
    root = tk.Tk()
    root.withdraw()  # We don't want a full GUI, so keep the root window from appearing

    # Show a "Save As" dialog box and return the path to the selected file
    file_path = filedialog.asksaveasfilename(
        title="Save the Excel file",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))  # Optional: specify file types
    )
    
    # Destroy the root window after use
    root.destroy()

    return file_path

# Use the function to get the file path
excel_save_path = ask_save_as_file()

# If the user cancels, the function will return '', so check if a path was provided
if excel_save_path:
    # Continue with saving the file using the provided path
    # ... your code to save the file goes here
    print(f"File will be saved to: {excel_save_path}")
else:
    print("File save operation was canceled.")

forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].to_excel(excel_save_path, index=False)

print(f"The forecast has been saved to {excel_save_path}")
