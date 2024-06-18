import pandas as pd
import os
import sys
from datetime import datetime

def main():
    # Command line arguments
    if len(sys.argv) != 2:
        print(r"Usage: python script.py <C:\Users\Aadhil\Documents\GitHub\COMP593-Lab3\salesdatas.py>")
        sys.exit(1)

    csv_file_path = sys.argv[1]
    if not os.path.isfile(csv_file_path):
        print(f"The file {csv_file_path} does not exist.")
        sys.exit(1)

    # Read the CSV file
    sales_data = pd.read_csv(csv_file_path)

    # Create Orders directory
    today_date = datetime.now().strftime("%Y-%m-%d")
    orders_dir = os.path.join(os.path.dirname(csv_file_path), f"Orders_{today_date}")
    if not os.path.exists(orders_dir):
        os.makedirs(orders_dir)

    # Process each order
    for order_id, order_data in sales_data.groupby('ORDER ID'):
        order_data = order_data.sort_values(by='ITEM NUMBER')
        order_data['TOTAL PRICE'] = order_data['ITEM QUANTITY'] * order_data['ITEM PRICE']

        # Calculate grand total
        grand_total = order_data['TOTAL PRICE'].sum()

        # Save to Excel file
        excel_file_path = os.path.join(orders_dir, f"Order_{order_id}.xlsx")
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            order_data.to_excel(writer, index=False, sheet_name='Order')

            # Get workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Order']

            # Format the prices
            money_fmt = workbook.add_format({'num_format': '$#,##0.00'})

            # Set column widths
            worksheet.set_column('A:A', 10)
            worksheet.set_column('B:B', 12)
            worksheet.set_column('C:C', 15)
            worksheet.set_column('D:D', 12)
            worksheet.set_column('E:E', 10)
            worksheet.set_column('F:F', 15, money_fmt)

            # Write the grand total
            worksheet.write(len(order_data) + 1, 4, 'Grand Total')
            worksheet.write(len(order_data) + 1, 5, grand_total, money_fmt)

if __name__ == "__main__":
    main()
