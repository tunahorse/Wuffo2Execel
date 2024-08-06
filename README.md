## Overview
------------

This script exports data from Wufoo forms using the Wufoo API and saves it to Excel files.

## Configuration
----------------

### Wufoo API Configuration

* **API_KEY**: Your Wufoo API key
* **WUFOO_SUBDOMAIN**: Your Wufoo subdomain

### Form and Brand Configuration

* **forms_and_brands**: A list of dictionaries containing the form hash and brand name for each form to export.

## Functions
-------------

### get_all_entries(form_hash)

* Fetches all entries for a given form hash using the Wufoo API.
* Returns a list of entries or `None` if an error occurs.

### get_field_titles(form_hash)

* Fetches the field titles for a given form hash using the Wufoo API.
* Returns a dictionary of field IDs to titles or `None` if an error occurs.

### clean_and_format_data(df, field_titles)

* Renames columns in the DataFrame using the field titles.
* Cleans up date fields by converting them to a standard format.
* Removes any empty columns.
* Returns the cleaned and formatted DataFrame.

### save_to_excel(df, filename)

* Saves the DataFrame to an Excel file using the `openpyxl` library.
* Formats the header row and auto-adjusts column widths.

## Main Execution
------------------

The script iterates over the `forms_and_brands` list and performs the following steps for each form:

1. Fetches entries and field titles using the `get_all_entries` and `get_field_titles` functions.
2. Creates a DataFrame from the entries.
3. Cleans and formats the DataFrame using the `clean_and_format_data` function.
4. Generates a timestamp for the file name.
5. Saves the DataFrame to an Excel file using the `save_to_excel` function.

## Example Use Case
--------------------

To use this script, simply update the `API_KEY`, `WUFOO_SUBDOMAIN`, and `forms_and_brands` variables with your own values. Then, run the script to export your Wufoo form data to Excel files.

