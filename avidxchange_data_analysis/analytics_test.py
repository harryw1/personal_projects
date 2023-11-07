import pandas as pd
from prophet import Prophet
import tkinter as tk
from tkinter import filedialog

# Function to ask for the file to open
def ask_for_file():
    # Set up the GUI to ask for the Excel file
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path

# Function to save the forecast to an Excel file
def save_forecast(forecast_df, category):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        forecast_df.to_excel(file_path, index=False)
        print(f"Forecast for {category} saved to {file_path}")

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

print(df.head())

# Loop through each category to create separate forecasts
for category in df['Heirarchy Name'].unique():
    
    df_category = df[df['Category'] == category]

    # Preprocess the data for Prophet (rename columns to 'ds' and 'y')
    df_prophet = df_category[['Revenue Month', 'Total Cost']].rename(columns={'Revenue Month': 'ds', 'Total Cost': 'y'})

    # Drop rows with missing values or fill them
    df_prophet = df_prophet.dropna()

    # Initialize and fit the Prophet model
    model = Prophet()
    model.fit(df_prophet)

    # Create a DataFrame to hold future dates for the forecast
    future_dates = model.make_future_dataframe(periods=365, freq='D')

    # Make the forecast
    forecast = model.predict(future_dates)

    # Save the forecast to an Excel file
    save_forecast(forecast, category)
