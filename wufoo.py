import requests
import json
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font

# Wufoo API configuration
API_KEY = 'apikey from wufoo'
WUFOO_SUBDOMAIN = 'domain for wufoo' 

# List of form hashes and brand names
forms_and_brands = [
    {'hash': 'hash-form', 'brand': 'can remove this or for naming'},
]

def get_all_entries(form_hash):
    all_entries = []
    page = 1
    entries_url = f'https://{WUFOO_SUBDOMAIN}.wufoo.com/api/v3/forms/{form_hash}/entries.json'
    
    while True:
        params = {
            'pageStart': (page - 1) * 100,
            'pageSize': 100
        }
        response = requests.get(entries_url, params=params, auth=(API_KEY, 'pass'))
        
        if response.status_code != 200:
            print(f"Error fetching entries for form {form_hash}: {response.status_code}")
            return None
        
        data = response.json()
        entries = data['Entries']
        
        if not entries:
            break
        
        all_entries.extend(entries)
        page += 1
    
    return all_entries

def get_field_titles(form_hash):
    fields_url = f'https://{WUFOO_SUBDOMAIN}.wufoo.com/api/v3/forms/{form_hash}/fields.json'
    response = requests.get(fields_url, auth=(API_KEY, 'pass'))
    
    if response.status_code != 200:
        print(f"Error fetching field titles for form {form_hash}: {response.status_code}")
        return None
    
    fields = response.json()['Fields']
    return {field['ID']: field['Title'] for field in fields}

def clean_and_format_data(df, field_titles):
    # Rename columns using field titles
    df.rename(columns=field_titles, inplace=True)
    
    # Clean up date fields
    date_columns = ['Date Created', 'Date Updated']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d %H:%M:%S')
    
    # Remove any empty columns
    df = df.dropna(axis=1, how='all')
    
    return df

def save_to_excel(df, filename):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Form Submissions')
        
        # Access the workbook and the active sheet
        workbook = writer.book
        worksheet = writer.sheets['Form Submissions']
        
        # Format header row
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
        
        # Auto-adjust columns' width
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

# Main execution for multiple forms
for form in forms_and_brands:
    form_hash = form['hash']
    brand_name = form['brand']

    # Fetch entries and field titles
    entries = get_all_entries(form_hash)
    field_titles = get_field_titles(form_hash)

    if entries and field_titles:
        # Create DataFrame
        df = pd.DataFrame(entries)
  
        
        # Clean and format DataFrame
        df = clean_and_format_data(df, field_titles)
        
        # Generate timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save to CSV
        # csv_filename = f'{brand_name}_submissions_{timestamp}.csv'
        # df.to_csv(csv_filename, index=False, date_format='%Y-%m-%d %H:%M:%S')
        # print(f"Saved {len(entries)} entries to {csv_filename}")
        
        # Save to Excel
        excel_filename = f'{brand_name}_submissions_{timestamp}.xlsx'
        save_to_excel(df, excel_filename)
        print(f"Saved {len(entries)} entries to {excel_filename}")

    else:
        print(f"No entries found or there was an error for form {form_hash}.")
