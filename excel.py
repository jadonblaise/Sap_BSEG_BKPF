import pandas as pd

# Try ISO-8859-1 or cp1252 if UTF-8 fails
df = pd.read_csv(
    'BSEG.csv',
    sep='|', 
    encoding='latin1',
    on_bad_lines='skip'
)

# Export to Excel
df.to_excel('BSEG1.xlsx', index=False, engine='openpyxl')

