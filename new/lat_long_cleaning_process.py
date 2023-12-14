import pandas as pd
import numpy as np

# all occurrence = Press Ctrl + Alt + Shift + J on Windows/Linux

# df = pd.read_excel(r'13 Dec regional.xlsx')

input_file_name = r'13 Dec regional.xlsx'
output_file_name = r'13 Dec regional_1.xlsx'

lat = 'GOOGLE_EARTH_LAT'
long = 'GOOGLE_EARTH_LONG'
latitude = 'LATITUDE'
longitude = 'LONGTITUDE'


def drop_blank_na_rows(df):
    # Drop rows with both GOOGLE_EARTH_LAT and GOOGLE_EARTH_LONG as 0/Null/na
    df = df.dropna(subset=[lat, long], how='all')
    df = df.dropna(subset=[lat, long], how='any')
    df = df[(df[lat] != 0) & (df[long] != 0)]
    return df


def replace_google_lat_with_latitude(df):
    # Replace or delete rows with GOOGLE_EARTH_LAT/GOOGLE_EARTH_LONG as 0/Null/na
    for index, row in df.iterrows():
        if pd.notna(row[lat]):
            # Check and replace GOOGLE_EARTH_LAT if valid latitude exists
            if pd.notna(row[latitude]):
                df.at[index, lat] = row[latitude]
        if pd.notna(row[long]):
            # Check and replace GOOGLE_EARTH_LONG if valid longitude exists
            if pd.notna(row[longitude]):
                df.at[index, long] = row[longitude]
    return df


def swap_google_lat_long(df):
    # Swap GOOGLE_EARTH_LAT and GOOGLE_EARTH_LONG if GOOGLE_EARTH_LAT >= 75 and GOOGLE_EARTH_LONG <= 30
    for index, row in df.iterrows():
        df[lat], df[long] = np.where((pd.isna(row[lat]) & pd.isna(row[long])) and ((df[lat] >= 75) & (df[long] <= 30)),
                                         (df[long], df[lat]),
                                         (df[lat], df[long]))
    return df


def remove_leading_zeros_from_google_lat_long(df):
    # Remove leading zeros
    df[lat] = df[lat].astype(str).str.lstrip('0')
    df[long] = df[long].astype(str).str.lstrip('0')
    return df


def remove_question_sign_from_google_lat_long(df):
    # Remove '?' from GOOGLE_EARTH_LAT and GOOGLE_EARTH_LONG
    df[lat] = df[lat].replace('\?', '', regex=True)
    df[long] = df[long].replace('\?', '', regex=True)
    return df


def choose_google_lat_long_where_comma_occur(df):
    # Remove parts after/before comma based on column
    df[lat] = df[lat].apply(lambda x: x.split(',')[0] if isinstance(x, str) and ',' in x else x)
    df[long] = df[long].apply(lambda x: x.split(',')[1] if isinstance(x, str) and ',' in x else x)
    return df


def replace_comma_with_space(df):
    # Replace ',_' / '_,' / ',' with space
    df[lat] = df[lat].replace({', ': ' ', ' ': ' ', ',': ' '}, regex=True)
    df[long] = df[long].replace({', ': ' ', ' ': ' ', ',': ' '}, regex=True)
    return df


def remove_leading_space_and_character_from_google_lat_long(df):
    # Remove leading spaces and unwanted characters from GOOGLE_EARTH_LAT and GOOGLE_EARTH_LONG
    df[lat] = df[lat].astype(str).str.lstrip(' ').replace('[^0-9.-]', '', regex=True)
    df[long] = df[long].astype(str).str.lstrip(' ').replace('[^0-9.-]', '', regex=True)
    return df


def remove_rows_with_one_blank_or_space(df):
    columns_to_check = [lat, long]
    mask = df[columns_to_check].apply(lambda x: x.map(lambda x: isinstance(x, (str, bytes)) and x.isspace() or pd.isna(x) or x == ''))
    df = df[~mask.any(axis=1)]
    return df


def remove_negative_value_contained_rows(df):
    df[lat] = pd.to_numeric(df[lat], errors='coerce')
    df[long] = pd.to_numeric(df[long], errors='coerce')
    df = df[df[lat].ge(0) & df[lat].le(30)]
    df = df[df[long].ge(0) & df[long].ge(75)]
    return df


def save_file(df):
    df.to_excel(output_file_name, index=False)


if __name__ == '__main__':
    df = pd.read_excel(input_file_name)
    df = drop_blank_na_rows(df)
    df = replace_google_lat_with_latitude(df)
    df = swap_google_lat_long(df)
    df = remove_leading_zeros_from_google_lat_long(df)
    df = remove_question_sign_from_google_lat_long(df)
    df = choose_google_lat_long_where_comma_occur(df)
    df = replace_comma_with_space(df)
    df = remove_leading_space_and_character_from_google_lat_long(df)
    df = remove_rows_with_one_blank_or_space(df)
    df = remove_negative_value_contained_rows(df)
    save_file(df)