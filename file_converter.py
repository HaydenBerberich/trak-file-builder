import pandas as pd

# Read the Excel file into a DataFrame, ensuring the UPC column is read as a string
df = pd.read_excel('files/red-test.xlsx', dtype={'UPC': str})

# Drop rows where all of the required columns are NaN (blank)
df = df.dropna(subset=['UPC', 'TITLE', 'ARTIST', 'MANUF', 'GENRE', 'CONFIG', 'DEPT', 'MISC', 'PRICE', 'VENDOR', 'COST'], how='all')

# Replace NaN values with an empty string
df = df.fillna('')

# Remove dollar signs from PRICE, COST, and LIST columns
df['PRICE'] = df['PRICE'].replace({r'\$': ''}, regex=True)
df['COST'] = df['COST'].replace({r'\$': ''}, regex=True)
df['LIST'] = df['LIST'].replace({r'\$': ''}, regex=True)

# Function to determine the selling price for CDs based on cost
def get_cd_price(cost):
    if cost <= 0:
        return None
    elif cost <= 0.99:
        return 1.99
    elif cost <= 1.99:
        return 3.99
    elif cost <= 2.99:
        return 5.99
    elif cost <= 3.99:
        return 7.99
    elif cost <= 4.99:
        return 9.99
    elif cost <= 5.99:
        return 11.99
    elif cost <= 6.99:
        return 13.99
    elif cost <= 7.99:
        return 14.99
    elif cost <= 9.99:
        return 15.99
    elif cost <= 11.99:
        return 16.99
    elif cost <= 12.99:
        return 17.99
    elif cost <= 13.99:
        return 21.99
    elif cost <= 14.99:
        return 22.99
    elif cost <= 15.99:
        return 24.99
    elif cost <= 16.99:
        return 25.99
    elif cost <= 17.99:
        return 26.99
    elif cost <= 18.99:
        return 27.99
    elif cost <= 19.99:
        return 29.99
    elif cost <= 20.99:
        return 31.99
    else:
        return cost * 1.4  # If cost > 20.99, price = cost * 1.4

# Function to determine the selling price for LPs based on cost
def get_lp_price(cost):
    if cost <= 0:
        return None
    elif cost <= 0.99:
        return 1.99
    elif cost <= 1.99:
        return 3.99
    elif cost <= 2.99:
        return 5.99
    elif cost <= 3.99:
        return 7.99
    elif cost <= 4.99:
        return 9.99
    elif cost <= 5.99:
        return 11.99
    elif cost <= 6.99:
        return 13.99
    elif cost <= 7.99:
        return 14.99
    elif cost <= 9.99:
        return 15.99
    elif cost <= 10.99:
        return 19.99
    elif cost <= 11.99:
        return 22.99
    elif cost <= 12.99:
        return 22.99
    elif cost <= 13.99:
        return 23.99
    elif cost <= 14.99:
        return 24.99
    elif cost <= 15.99:
        return 25.99
    elif cost <= 16.99:
        return 27.99
    elif cost <= 17.99:
        return 29.99
    elif cost <= 18.99:
        return 30.99
    elif cost <= 19.99:
        return 31.99
    elif cost <= 20.99:
        return 33.99
    elif cost <= 21.99:
        return 34.99
    elif cost <= 22.99:
        return 35.99
    elif cost <= 23.99:
        return 36.99
    elif cost <= 24.99:
        return 38.99
    elif cost <= 25.99:
        return 39.99
    elif cost <= 26.99:
        return 41.99
    elif cost <= 27.99:
        return 44.99
    elif cost <= 28.99:
        return 45.99
    elif cost <= 29.99:
        return 46.99
    elif cost <= 30.99:
        return 47.99
    elif cost <= 31.99:
        return 48.99
    elif cost <= 32.99:
        return 49.99
    elif cost <= 33.99:
        return 50.99
    elif cost <= 34.99:
        return 52.99
    elif cost <= 35.99:
        return 54.99
    elif cost <= 36.99:
        return 55.99
    elif cost <= 37.99:
        return 58.99
    elif cost <= 38.99:
        return 59.99
    elif cost <= 39.99:
        return 61.99
    elif cost <= 40.99:
        return 62.99
    elif cost <= 41.99:
        return 64.99
    elif cost <= 42.99:
        return 65.99
    elif cost <= 43.99:
        return 66.99
    elif cost <= 44.99:
        return 68.99
    elif cost <= 45.99:
        return 69.99
    elif cost <= 46.99:
        return 71.99
    elif cost <= 47.99:
        return 73.99
    elif cost <= 48.99:
        return 74.99
    elif cost <= 49.99:
        return 76.99
    else:
        return cost * 1.4  # If cost > 49.99, price = cost * 1.4

# Open the output file for writing in the current directory
with open('trakdelim.txt', 'w') as file:
    
    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        # Convert fields to strings
        upc = str(row['UPC'])  # Keep leading zeros
        
        # Set department based on CONFIG if DEPT is empty
        if not row['DEPT']:
            if row['CONFIG'] == 'CD':
                dept = '02'
            elif row['CONFIG'] == 'LP':
                dept = '01'
            else:
                dept = ''
        else:
            dept = str(int(row['DEPT']))
            if len(dept) == 1:
                dept = '0' + dept
                
        # Calculate LIST and PRICE based on CONFIG if missing
        cost_value = float(row['COST']) if row['COST'] else 0
        
        if row['CONFIG'] == 'CD':
            # If price is missing, determine it from the cost
            if not row['PRICE']:
                calculated_price = get_cd_price(cost_value)
                if calculated_price:
                    price = format(calculated_price, '.2f').replace('.', '')
                else:
                    price = ''
            else:
                price = format(float(row['PRICE']), '.2f').replace('.', '') if row['PRICE'] else ''
            
            # If list price is missing, use the same calculation as price
            if not row['LIST']:
                calculated_list = get_cd_price(cost_value)
                if calculated_list:
                    list_price = format(calculated_list, '.2f').replace('.', '')
                else:
                    list_price = ''
            else:
                list_price = format(float(row['LIST']), '.2f').replace('.', '') if row['LIST'] else ''
        elif row['CONFIG'] == 'LP':
            # If price is missing, determine it from the cost using the LP pricing table
            if not row['PRICE']:
                calculated_price = get_lp_price(cost_value)
                if calculated_price:
                    price = format(calculated_price, '.2f').replace('.', '')
                else:
                    price = ''
            else:
                price = format(float(row['PRICE']), '.2f').replace('.', '') if row['PRICE'] else ''
            
            # If list price is missing, use the same calculation as price
            if not row['LIST']:
                calculated_list = get_lp_price(cost_value)
                if calculated_list:
                    list_price = format(calculated_list, '.2f').replace('.', '')
                else:
                    list_price = ''
            else:
                list_price = format(float(row['LIST']), '.2f').replace('.', '') if row['LIST'] else ''
        else:
            # For other CONFIG types, just format the existing values
            price = format(float(row['PRICE']), '.2f').replace('.', '') if row['PRICE'] else ''
            list_price = format(float(row['LIST']), '.2f').replace('.', '') if row['LIST'] else ''
        
        cost = format(float(row['COST']), '.2f').replace('.', '') if row['COST'] else ''
        
        # Format the row data according to the specified layout
        formatted_row = (
            f"C|{upc}|{row['TITLE']}|{row['ARTIST']}|{row['MANUF']}|||{row['GENRE']}|||{row['MISC']}|{row['CONFIG']}|||{dept}|{list_price}||||||{row['VENDOR']}|{cost}|||||||||{price}"
        )

        # Write the formatted data to the output file
        file.write(formatted_row + '\n')