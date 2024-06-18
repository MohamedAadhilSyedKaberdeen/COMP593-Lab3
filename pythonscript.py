import sys
import os
import pandas as pd
import xlsxwriter
import datetime

def get_sales_data_csv():
    if len(sys.argv)!= 2:
        print("Error: No command line parameter provided. Please provide the path to the sales data CSV file.")
        sys.exit(1)

    csv_file_path = sys.argv[1]

    if not os.path.isfile(csv_file_path):
        print(f"Error: The file '{csv_file_path}' does not exist.")
        sys.exit(1)

    return csv_file_path

import datetime

def create_orders_dir(sales_csv):
    orders_dir = os.path.join(os.path.dirname(sales_csv), f"Orders_{datetime.date.today().isoformat()}")
    if not os.path.exists(orders_dir):
        os.makedirs(orders_dir)
    return orders_dir

def process_sales_data_csv(sales_csv, orders_dir):
    df = pd.read_csv(sales_csv)
    df['TOTAL PRICE'] = df['ITEM QUANTITY'] * df['ITEM PRICE']

    for order_id, group in df.groupby('ORDER ID'):
        order_df = group.drop('ORDER ID', axis=1).sort_values('ITEM NUMBER')
        order_df.loc['GRAND TOTAL'] = order_df.sum(numeric_only=True)
        order_df.loc['GRAND TOTAL', 'ITEM NUMBER'] = ''

        writer = pd.ExcelWriter(os.path.join(orders_dir, f"Order_{order_id}.xlsx"), engine='xlsxwriter')
        order_df.to_excel(writer, index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        money_format = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})

        for col in range(3, 6):
            worksheet.set_column(col, col, 15, money_format)

        writer.save()

if __name__ == '__main__':
    sales_csv = get_sales_data_csv()
    orders_dir = create_orders_dir_(sales_csv)
    process_sales_data_csv(sales_csv, orders_dir)
