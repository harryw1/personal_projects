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
        filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))  # Optional: specify file types
    )
    
    return file_path

# Use the function to get the file path
excel_file_path = ask_for_file()

# Data Preprocessing
df = pd.read_excel(excel_file_path)

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

# Save the forecast to an Excel file
forecast_output_path = filedialog.asksaveasfilename # filedialog.asksaveasfilename to get path from user
forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].to_excel(forecast_output_path, index=False)

print(f"The forecast has been saved to {forecast_output_path}")
