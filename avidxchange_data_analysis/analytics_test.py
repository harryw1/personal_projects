from prophet import Prophet
import pandas as pd

# Assuming 'time_series_data' is your preprocessed DataFrame with the 'Revenue Month' and 'Total Cost'
# for a given category. You need to rename the columns to 'ds' and 'y' as Prophet expects.

# Preparing the data
df_prophet = time_series_data[['Revenue Month', 'Total Cost']].rename(columns={'Revenue Month': 'ds', 'Total Cost': 'y'})

# Create and fit the Prophet model
# The daily_seasonality, weekly_seasonality, and yearly_seasonality parameters
# can be adjusted according to your data's specific seasonality
model = Prophet(daily_seasonality=False, weekly_seasonality=False, yearly_seasonality=True)
model.fit(df_prophet)

# Make future predictions
# Replace 'periods' with the number of periods you want to forecast
future = model.make_future_dataframe(periods=365, freq='D')
forecast = model.predict(future)

# Plot the forecast
fig1 = model.plot(forecast)
fig2 = model.plot_components(forecast)

# Show the plot if you're running this in an interactive environment like Jupyter Notebook
plt.show()
