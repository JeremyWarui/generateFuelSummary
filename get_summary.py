'''
Module: get_summary
This module reads the data from the excel file and writes the
data to a new excel file.
'''

import os
from sys import argv
import pandas as pd


def generate_fuel_summary():
    '''
    Function that generates a summary of the fuel details from the input
    '''
    # get the name of the excel file
    input_file = argv[1]
    if not input_file.endswith('.XLSX'):
        print('Invalid file format. Please provide an excel file')
        return
    # use name of input_file to generate output file
    formatted_name = os.path.splitext(os.path.basename(input_file))[0]
    output_file = f"{formatted_name}_summary.xlsx"

    # read the excel file
    data_file = pd.read_excel(argv[1])
    # sort the data by registration_num and ticket
    data_file.sort_values(by=['Registration_num', 'Ticket'], inplace=True)

    # Replace comma with decimal point in Quantity column
    # data_file['Quantity'] = data_file['Quantity'].str.replace(
    #     ',', '.').astype(float)
    data_file['Quantity'] = data_file['Quantity'].str.replace(
        ',', '.').str.replace('\xa0', '').astype(float)

    # Select only the desired columns
    selected_cols = ['Registration_num', 'Ticket',
                     'Product_or_Article', 'Quantity', 'Amount_incl_Tax']
    data_selected = data_file[selected_cols]

    # Group data by Registration_num and Product_or_Article
    grouped_data = data_file.groupby(['Registration_num', 'Product_or_Article']).agg({
        'Quantity': 'sum',
        'Amount_incl_Tax': 'sum'
    }).reset_index()

    # Write to separate sheets in the same Excel workbook
    with pd.ExcelWriter(output_file) as writer:
        data_selected.to_excel(writer, sheet_name='Summary', index=False)
        grouped_data.to_excel(writer, sheet_name='Totals', index=False)

    print(f"Data saved to {output_file}")

if __name__ == "__main__":
    generate_fuel_summary()
