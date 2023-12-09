import pandas as pd
import numpy as np

# all occurrence = Press Ctrl + Alt + Shift + J on Windows/Linux

df = pd.read_excel(r'I:\Openpyxl_tutorial\Parts\4 Dec regional ( test ) - Copy.xlsx')

lat = 'GOOGLE_EARTH_LAT'
long = 'GOOGLE_EARTH_LONG'
latitude = 'LATITUDE'
longitude = 'LONGTITUDE'

# Step 1: Drop rows with both GOOGLE_EARTH_LAT and GOOGLE_EARTH_LONG as 0/Null/na
df = df.dropna(subset=[lat, long], how='all')
df = df[(df[lat] != 0) & (df[long] != 0)]


# Step 2: Replace or delete rows with GOOGLE_EARTH_LAT/GOOGLE_EARTH_LONG as 0/Null/na
for index, row in df.iterrows():
    if pd.notna(row[lat]):
        # Check and replace GOOGLE_EARTH_LAT if valid latitude exists
        if pd.notna(row[latitude]):
            df.at[index, lat] = row[latitude]
    if pd.notna(row[long]):
        # Check and replace GOOGLE_EARTH_LONG if valid longitude exists
        if pd.notna(row[longitude]):
            df.at[index, long] = row[longitude]

# Step 3: Swap GOOGLE_EARTH_LAT and GOOGLE_EARTH_LONG if GOOGLE_EARTH_LAT >= 75 and GOOGLE_EARTH_LONG <= 30
for index, row in df.iterrows():
    df[lat], df[long] = np.where((pd.isna(row[lat]) & pd.isna(row[long])) and ((df[lat] >= 75) & (df[long] <= 30)),
                                     (df[long], df[lat]),
                                     (df[lat], df[long]))

# Step 4: Remove leading zeros
df[lat] = df[lat].astype(str).str.lstrip('0')
df[long] = df[long].astype(str).str.lstrip('0')

# Step 5: Remove '?' from GOOGLE_EARTH_LAT and GOOGLE_EARTH_LONG
df[lat] = df[lat].replace('\?', '', regex=True)
df[long] = df[long].replace('\?', '', regex=True)

# Step 6: Remove parts after/before comma based on column
df[lat] = df[lat].apply(lambda x: x.split(',')[0] if isinstance(x, str) and ',' in x else x)
df[long] = df[long].apply(lambda x: x.split(',')[1] if isinstance(x, str) and ',' in x else x)

# Step 7: Replace ',_' / '_,' / ',' with space
df[lat] = df[lat].replace({', ': ' ', ' ': ' ', ',': ' '}, regex=True)
df[long] = df[long].replace({', ': ' ', ' ': ' ', ',': ' '}, regex=True)

# Step 4: Remove leading spaces and unwanted characters from GOOGLE_EARTH_LAT and GOOGLE_EARTH_LONG
df[lat] = df[lat].astype(str).str.lstrip(' ').replace('[^0-9.-]', '', regex=True)
df[long] = df[long].astype(str).str.lstrip(' ').replace('[^0-9.-]', '', regex=True)

# Save the result to a new Excel file
df.to_excel('I:\Openpyxl_tutorial\Parts\Regional_Final_Result.xlsx', index=False)