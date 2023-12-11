# import module
import pandas as pd
import re
import numpy as np

# Paste path for the file
file_path = r'C:\Users\mfmohammad\OneDrive - UNICEF\Desktop\Clean Address\Address List To Clean.xlsx'

# read the file
address_df = pd.read_excel(file_path)

column_to_clean = ['Mailing City', 'Mailing State/Province', 'Mailing Zip/Postal Code', 'Mailing Country']

def clean_data(column):
    pattern = r'[^A-Za-z0-9 ]'
    cleaned_column = column.astype(str).str.replace(pattern, '', regex=True).str.title().str.strip()
    cleaned_column = cleaned_column.replace('Nan','', regex=False)
    return cleaned_column

def extract_postal_code(address):
    pattern = re.compile(r'\b\d{5}\b')

    if pd.isna(address):
        return np.nan
    address = address.strip()
    match = pattern.search(address)

    if match:
        return match.group()
    else:
        pass

def populate_country(row):
    postal_code = row['Mailing Zip/Postal Code']

    brunei_postcode_pattern = re.compile(r'^[A-Za-z]{2}\d{4}$')
    
    if isinstance(postal_code, str):
        if len(postal_code) == 6 and brunei_postcode_pattern.match(postal_code):
            return 'Brunei'
        elif len(postal_code) == 6:
            return 'Singapore'
        elif len(postal_code) == 5:
            return 'Malaysia'
    
    return row['Mailing Country']



def main():
    address_df[column_to_clean] = address_df[column_to_clean].apply(clean_data)
    address_df['Mailing Zip/Postal Code'] = address_df['Mailing Zip/Postal Code'].astype(str).str.upper()
    address_df['Test'] = address_df['Mailing Street'].apply(extract_postal_code)
    address_df['Mailing Country'] = address_df.apply(populate_country, axis=1)
    address_df['Mailing Country'] = address_df['Mailing Country'].replace('Brunei Darussalam', 'Brunei')

if __name__ == "__main__":
    main()

# rename output file
file_name = "test.xlsx"

# save the file
address_df.to_excel(file_name, index=False)

print(f'Excel file cleaned and save as {file_name} successfully')






