import csv
from datetime import datetime
import pyodbc

def retrieve_data_from_database():
    # Connect to the Access database
    conn = pyodbc.connect(
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\andrew\Desktop\PycharmProjects\pythonProject\Bosda-BizLibrary-20230825.accdb')

    # Create a cursor to execute SQL queries
    cursor = conn.cursor()

    # Retrieve data from Invn_ProdLocation table
    cursor.execute("""
        SELECT item_id, warehouse_id, qty, subclass_id
        FROM Invn_ProdLocation
        WHERE warehouse_id <> 'LGRET'
            AND loc_rack NOT IN ('V', 'X', 'Y', 'Z', 'RCV')
        """)
    stock_dict = {}
    stock_warehouses = []

    subclasses = {}

    for row in cursor:
        item_id = row.item_id
        warehouse_id = row.warehouse_id
        qty = row.qty
        subclass = row.subclass_id

        # Skip rows with any None value
        if None in (item_id, warehouse_id, qty):
            continue

        stock_key = (item_id)

        if stock_key in stock_dict:
            stock_dict[stock_key].append(qty)  # Group duplicate quantities together
        else:
            stock_dict[stock_key] = [qty]

        stock_warehouses.append(warehouse_id)
        subclasses[item_id]=subclass

    # Retrieve data from Invn_Reserved table
    cursor.execute("SELECT item_id, qty_reserved, warehouse_id FROM Invn_Reserved WHERE NOT isInTransit")
    reserved_dict = {}
    reserved_warehouses = []

    for row in cursor:
        item_id = row.item_id
        qty_reserved = row.qty_reserved
        warehouse_id = row.warehouse_id

        # Skip rows with any None value
        if None in (item_id, qty_reserved, warehouse_id):
            continue

        reserved_key = (item_id)

        if reserved_key in reserved_dict:
            reserved_dict[reserved_key].append(qty_reserved)  # Group duplicate quantities together
        else:
            reserved_dict[reserved_key] = [qty_reserved]

        reserved_warehouses.append(warehouse_id)

    # Retrieve data from AltPartNumber table
    cursor.execute("SELECT item_id, alt_item_id FROM AltPartNumber WHERE alt_item_id IS NOT NULL")
    alt_item_dict = {}
    for row in cursor:
        item_id = row.item_id
        alt_item_id = row.alt_item_id
        if item_id in alt_item_dict:
            alt_item_dict[item_id].append(alt_item_id)
        else:
            alt_item_dict[item_id] = [alt_item_id]

    # Close the cursor and the database connection
    cursor.close()
    conn.close()

    # Ensure stock_dict and reserved_dict have the same set of item_id keys
    for item_id in list(stock_dict.keys()):
        if item_id not in reserved_dict:
            reserved_dict[item_id] = [0]

    for item_id in list(reserved_dict.keys()):
        if item_id not in stock_dict:
            stock_dict[item_id] = [0]

    for item_id in stock_dict:
        sum = 0
        for stock in stock_dict[item_id]:
            sum += stock
        stock_dict[item_id] = sum

    for item_id in reserved_dict:
        sum = 0
        for reserved in reserved_dict[item_id]:
            sum += reserved
        reserved_dict[item_id] = sum

    alt_item_qtys = {}
    for item_id in alt_item_dict:
        if item_id in stock_dict:
            sum = 0
            for ind in alt_item_dict[item_id]:
                if ind in stock_dict:
                    sum += stock_dict[ind]
            alt_item_qtys[item_id] = sum

    alt_item_reserved = {}
    for item_id in alt_item_dict:
        if item_id in reserved_dict:
            reserved = 0
            for ind in alt_item_dict[item_id]:
                if ind in reserved_dict:
                    reserved += reserved_dict[ind]
            alt_item_reserved[item_id] = reserved

    stock_with_alts = {}
    for item_id in stock_dict:
        if item_id in alt_item_qtys:
            stock_with_alts[item_id] = stock_dict[item_id] + alt_item_qtys[item_id]
        else:
            stock_with_alts[item_id] = stock_dict[item_id]

    reserved_with_alts = {}
    for item_id in reserved_dict:
        if item_id in alt_item_reserved:
            reserved_with_alts[item_id] = reserved_dict[item_id] + alt_item_reserved[item_id]
        else:
            reserved_with_alts[item_id] = reserved_dict[item_id]

    final = {}
    for item_id in stock_with_alts:
        final[item_id] = stock_with_alts[item_id] - reserved_with_alts[item_id]

    return final, stock_dict, reserved_dict, alt_item_qtys, alt_item_reserved, subclasses


final, stock_dict, reserved_dict, alt_item_qtys, alt_item_reserved, subclasses= retrieve_data_from_database()

final_subclasses = {}
for item_id in subclasses:
    if item_id in stock_dict:
        final_subclasses[item_id]=subclasses[item_id]

rows = []
for item_id in final:
    row = {
        "Item ID": item_id,
        "Subclass": final_subclasses.get(item_id,0),
        "Stock": stock_dict.get(item_id, 0),
        "Reserved": reserved_dict.get(item_id, 0),
        "Alternate Qtys": alt_item_qtys.get(item_id, 0),
        "Alternate Reserves": alt_item_reserved.get(item_id, 0),
        "Available": final[item_id]
    }
    rows.append(row)

# Get the current date
current_date = datetime.now().strftime("%Y-%m-%d")

# Construct the CSV file name
csv_filename = f"BizStock_report_{current_date}.csv"

# Write the data to the CSV file
with open(csv_filename, mode="w", newline="") as csvfile:
    fieldnames = ["Item ID","Subclass", "Stock", "Reserved", "Alternate Qtys", "Alternate Reserves", "Available"]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

    # Write the header row
    writer.writeheader()

    # Write the data rows
    writer.writerows(rows)
