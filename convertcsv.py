import pandas as pd

# Step 1: Read the Excel file
df = pd.read_excel('Susa_BergerUndCo.xlsx', engine='openpyxl')  # or remove engine if unsure

# Step 2: Export to CSV
df.to_csv('summary.csv', index=False)
