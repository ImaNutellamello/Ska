import pandas as pd
from datetime import datetime

# Read the Excel file
file_path = 'data.xlsx'
df = pd.read_excel(file_path)

# Convert date strings to datetime objects
df['x'] = pd.to_datetime(df['x_column'], format='%d/%m/%Y')
df['y'] = pd.to_datetime(df['y_column'], format='%d/%m/%Y')
df['z'] = pd.to_datetime(df['z_column'], format='%d/%m/%Y')

# Convert datetime objects to the number of days
df['x'] = (df['x'] - df['x'].min()).dt.days
df['y'] = (df['y'] - df['y'].min()).dt.days
df['z'] = (df['z'] - df['z'].min()).dt.days

# Calculate the coordinates using the section formula
df['y_coord'] = (df['z'] * df['x_coord'] - df['x'] * df['z_coord']) / (df['z'] - df['x'])
df['x_coord'] = df['x']
df['z_coord'] = df['z']

# Calculate the ratio
df['ratio'] = (df['y_coord'] - df['x_coord']) / (df['z_coord'] - df['x_coord'])

# Print the result
print(df[['x_coord', 'y_coord', 'z_coord', 'ratio']])