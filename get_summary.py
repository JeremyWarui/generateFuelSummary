'''
Module: get_summary
This module reads the data from the excel file and writes the
data to a new excel file.
'''

from sys import argv
import pandas as pd


def generate_fuel_summary():
    '''
    Function that generates a summary of the fuel details from the input
    '''
    # read the excel file
    data_file = pd.read_excel(argv[1])

    # sort the data by registration_num and ticket
    data_file.sort_values(by=['Reg_num', 'Ticket'], inplace=True)

    # Replace comma with decimal point in Quantity column
    # data_file['Quantity'] = data_file['Quantity'].str.replace(
    #     ',', '.').astype(float)
    data_file['Quantity'] = data_file['Quantity'].str.replace(
        ',', '.').str.replace('\xa0', '').astype(float)

    # Select only the desired columns
    selected_cols = ['Reg_num', 'Ticket',
                     'Product_or_Article', 'Quantity', 'Amount_incl_Tax']
    data_selected = data_file[selected_cols]

    # Group data by Registration_num and Product_or_Article
    grouped_data = data_file.groupby(['Reg_num', 'Product_or_Article']).agg({
        'Quantity': 'sum',
        'Amount_incl_Tax': 'sum'
    }).reset_index()

    # Write to separate sheets in the same Excel workbook
    with pd.ExcelWriter('combined_results.xlsx') as writer:
        data_selected.to_excel(writer, sheet_name='Summary', index=False)
        grouped_data.to_excel(writer, sheet_name='Totals', index=False)

    print("Data saved to 'combined_results.xlsx'")


if __name__ == "__main__":
    generate_fuel_summary()
