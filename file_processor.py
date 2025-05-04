"""
File processing functions for TRAK file generation.
"""
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from price_calculator import get_cd_price, get_lp_price

def process_input_data(input_file_path, output_dir):
    """
    Process the input data and create a complete Excel file.
    
    Args:
        input_file_path (str): Path to the input Excel file
        output_dir (str): Directory to save the output file
        
    Returns:
        tuple: (DataFrame with processed data, path to the Excel output file)
    """
    # Read the Excel file into a DataFrame, ensuring the UPC column is read as a string
    df = pd.read_excel(input_file_path, dtype={'UPC': str})

    # Drop rows where all of the required columns are NaN (blank)
    df = df.dropna(subset=['UPC', 'TITLE', 'ARTIST', 'MANUF', 'GENRE', 'CONFIG', 'DEPT', 'MISC', 'PRICE', 'VENDOR', 'COST'], how='all')

    # Replace NaN values with an empty string
    df = df.fillna('')

    # Set MISC and VENDOR to MANUF if they're not provided
    df['MISC'] = df.apply(lambda row: row['MANUF'] if not row['MISC'] else row['MISC'], axis=1)
    df['VENDOR'] = df.apply(lambda row: row['MANUF'] if not row['VENDOR'] else row['VENDOR'], axis=1)

    # Remove dollar signs from PRICE, COST, and LIST columns
    df['PRICE'] = df['PRICE'].replace({r'\$': ''}, regex=True)
    df['COST'] = df['COST'].replace({r'\$': ''}, regex=True)
    df['LIST'] = df['LIST'].replace({r'\$': ''}, regex=True)

    # Process each row to fill in missing values
    processed_rows = []
    
    for index, row in df.iterrows():
        processed_row = row.copy()
        
        # Set department based on CONFIG if DEPT is empty
        if not processed_row['DEPT']:
            if processed_row['CONFIG'] == 'CD':
                processed_row['DEPT'] = '02'
            elif processed_row['CONFIG'] == 'LP':
                processed_row['DEPT'] = '01'
        else:
            # Ensure DEPT is formatted as two digits
            dept = str(int(processed_row['DEPT']))
            if len(dept) == 1:
                processed_row['DEPT'] = '0' + dept
            else:
                processed_row['DEPT'] = dept
                
        # Calculate LIST and PRICE based on CONFIG if missing
        cost_value = float(processed_row['COST']) if processed_row['COST'] else 0
        
        if processed_row['CONFIG'] == 'CD':
            # If price is missing, determine it from the cost
            if not processed_row['PRICE']:
                calculated_price = get_cd_price(cost_value)
                if calculated_price:
                    processed_row['PRICE'] = format(calculated_price, '.2f')
            
            # If list price is missing, use the same calculation as price
            if not processed_row['LIST']:
                calculated_list = get_cd_price(cost_value)
                if calculated_list:
                    processed_row['LIST'] = format(calculated_list, '.2f')
                    
        elif processed_row['CONFIG'] == 'LP':
            # If price is missing, determine it from the cost using the LP pricing table
            if not processed_row['PRICE']:
                calculated_price = get_lp_price(cost_value)
                if calculated_price:
                    processed_row['PRICE'] = format(calculated_price, '.2f')
            
            # If list price is missing, use the same calculation as price
            if not processed_row['LIST']:
                calculated_list = get_lp_price(cost_value)
                if calculated_list:
                    processed_row['LIST'] = format(calculated_list, '.2f')
        
        # Format cost value
        if processed_row['COST']:
            processed_row['COST'] = format(float(processed_row['COST']), '.2f')
            
        processed_rows.append(processed_row)
    
    # Create a new DataFrame with processed data
    processed_df = pd.DataFrame(processed_rows)
    
    # Reorder columns to match the desired output format
    ordered_columns = ['UPC', 'TITLE', 'ARTIST', 'MANUF', 'GENRE', 'CONFIG', 'DEPT', 'MISC', 'LIST', 'PRICE', 'VENDOR', 'COST']
    processed_df = processed_df[ordered_columns]
    
    # Create output file paths
    excel_output_path = os.path.join(output_dir, 'trakdelim.xlsx')
    
    # Save to a new Excel file
    processed_df.to_excel(excel_output_path, index=False)
    
    return processed_df, excel_output_path

def generate_delimited_file(df, output_dir):
    """
    Generate the delimited text file from the processed data.
    
    Args:
        df (DataFrame): Processed data
        output_dir (str): Directory to save the output file
        
    Returns:
        str: Path to the generated text file
    """
    # Create output file path
    text_output_path = os.path.join(output_dir, 'trakdelim.txt')
    
    # Open the output file for writing in the specified directory
    with open(text_output_path, 'w') as file:
        
        # Iterate over each row in the DataFrame
        for index, row in df.iterrows():
            # Format the numeric values for the delimited file (no decimal points)
            # Ensure values are strings before calling replace
            list_price = str(row['LIST']).replace('.', '') if pd.notna(row['LIST']) else ''
            price = str(row['PRICE']).replace('.', '') if pd.notna(row['PRICE']) else ''
            cost = str(row['COST']).replace('.', '') if pd.notna(row['COST']) else ''
            
            # Format the row data according to the specified layout
            formatted_row = (
                f"C|{row['UPC']}|{row['TITLE']}|{row['ARTIST']}|{row['MANUF']}|||{row['GENRE']}|||{row['MISC']}|{row['CONFIG']}|||{row['DEPT']}|{list_price}||||||{row['VENDOR']}|{cost}|||||||||{price}"
            )

            # Write the formatted data to the output file
            file.write(formatted_row + '\n')
            
    return text_output_path

def create_new_spreadsheet(file_path):
    """
    Create a new spreadsheet with required and optional columns.
    
    Args:
        file_path (str): Path where the new spreadsheet should be saved
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Create a new DataFrame with the required columns
        required_columns = ['UPC', 'TITLE', 'ARTIST', 'MANUF', 'GENRE', 'CONFIG', 'COST']
        optional_columns = ['DEPT', 'MISC', 'LIST', 'PRICE', 'VENDOR']
        all_columns = required_columns + optional_columns
        
        # Create an empty DataFrame with the columns
        df = pd.DataFrame(columns=all_columns)
        
        # Save to Excel
        df.to_excel(file_path, index=False)
        
        # Format the Excel file with styles
        workbook = load_workbook(file_path)
        sheet = workbook.active
        
        # Format the header row - required columns in red, optional in black
        for col_idx, column_name in enumerate(all_columns, start=1):
            cell = sheet.cell(row=1, column=col_idx)
            if column_name in required_columns:
                cell.font = Font(color="FF0000", bold=True)  # Red
            else:
                cell.font = Font(bold=True)  # Black, bold
        
        # Save the formatted workbook
        workbook.save(file_path)
        return True
    except Exception:
        return False