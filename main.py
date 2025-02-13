import xlwings as xw
import tkinter as tk
from tkinter import ttk
import pandas as pd
from fuzzywuzzy import fuzz, process

# Global variable to store the selected row
selected_row = None

def open_matcher():
    # Import data from Excel
    wb = xw.Book('demo.xlsm')
    sht = wb.sheets['Sheet1']
    # Get the data from the csv
    df = pd.read_csv('Store Parts Inventory.csv')
    manufacturers = sorted(df['Provider'].dropna().astype(str).unique().tolist())

    root = tk.Tk()
    root.title("Excel Matcher")
    root.geometry("600x600")
    
    # Create a frame for the top widgets
    top_frame = ttk.Frame(root)
    top_frame.pack(fill='x', padx=10, pady=10)
    
    # Add an entry field to the GUI
    entry_label = ttk.Label(top_frame, text="Enter Value:")
    entry_label.pack(side='left')
    
    entry = ttk.Entry(top_frame, width=30)
    entry.pack(side='left', padx=10)
    
    # Add a dropdown to select manufacturers
    manufacturer_label = ttk.Label(top_frame, text="Select Manufacturer:")
    manufacturer_label.pack(side='left', padx=(20, 0))
    
    manufacturer_var = tk.StringVar()
    manufacturer_dropdown = ttk.Combobox(top_frame, textvariable=manufacturer_var, width=30)
    manufacturer_dropdown['values'] = manufacturers
    manufacturer_dropdown.pack(side='left', padx=10)

    # Create a frame to contain tree1 and its scrollbar
    tree1_frame = ttk.Frame(root)
    tree1_frame.pack(expand=True, fill='both', padx=10, pady=10)

    # Create a treeview for the first table
    columns = ("Part Number", "Description", "Manufacturer", "Cost")
    tree1 = ttk.Treeview(tree1_frame, columns=columns, show='headings')
    
    # Define headings for the first table
    for col in columns:
        tree1.heading(col, text=col)

    # Create a vertical scrollbar for tree1
    tree1_scrollbar = ttk.Scrollbar(tree1_frame, orient="vertical", command=tree1.yview)
    tree1.configure(yscrollcommand=tree1_scrollbar.set)

    # Pack the treeview and scrollbar
    tree1.pack(side='left', expand=True, fill='both')
    tree1_scrollbar.pack(side='right', fill='y')
    
    # Bind the selection event to a function
    tree1.bind("<<TreeviewSelect>>", lambda event: on_tree_select(event, tree1))
    
    # Create a frame for the buttons
    button_frame = ttk.Frame(root)
    button_frame.pack(fill='x', padx=10, pady=10)

    # Add a Search button to the button frame
    search_button = ttk.Button(button_frame, text="Search", command=lambda: fuzzymatch(entry.get(), manufacturer_dropdown.get(), tree1))
    search_button.pack(side='left', padx=5)

    # Add a "Search Recommendations" button to the button frame
    search_recommendations_button = ttk.Button(button_frame, text="Search Recommendations", command=search_recommendations)
    search_recommendations_button.pack(side='left', padx=5)
    
    # Create a frame to contain tree2 and its scrollbar
    tree2_frame = ttk.Frame(root)
    tree2_frame.pack(expand=True, fill='both', padx=10, pady=10)

    # Create a treeview for the second table
    tree2 = ttk.Treeview(tree2_frame, columns=columns, show='headings')
    
    # Define headings for the second table
    for col in columns:
        tree2.heading(col, text=col)

    # Create a vertical scrollbar for tree2
    tree2_scrollbar = ttk.Scrollbar(tree2_frame, orient="vertical", command=tree2.yview)
    tree2.configure(yscrollcommand=tree2_scrollbar.set)

    # Pack the treeview and scrollbar
    tree2.pack(side='left', expand=True, fill='both')
    tree2_scrollbar.pack(side='right', fill='y')
    
    bottom_frame = ttk.Frame(root)
    bottom_frame.pack(fill='x', padx=10, pady=10)
    
    # Add a quantity entry field
    quantity_label = ttk.Label(bottom_frame, text="Quantity:")
    quantity_label.pack(side='left', padx=(0, 5))
    quantity_entry = ttk.Entry(bottom_frame, width=10)
    quantity_entry.pack(side='left', padx=(0, 20))

    # Add a margin entry field
    margin_label = ttk.Label(bottom_frame, text="Margin:")
    margin_label.pack(side='left', padx=(0, 5))
    margin_entry = ttk.Entry(bottom_frame, width=10)
    margin_entry.pack(side='left', padx=(0, 20))

    # Add an "Add to Quote" button at the bottom middle
    add_to_quote_button = ttk.Button(bottom_frame, text="Add to Quote", command=lambda: add_to_quote(sht, entry.get(), quantity_entry.get(), margin_entry.get()))
    add_to_quote_button.pack(side='left', padx=5)
    
    root.mainloop()

def fuzzymatch(customer_desc, selected_manufacturer, tree1, min_score=70):
    df_csv = pd.read_csv('Store Parts Inventory.csv')
    
    if selected_manufacturer:
        df_csv = df_csv[df_csv['Provider'] == selected_manufacturer]
    
    matches = process.extract(customer_desc, df_csv['Description'], scorer=fuzz.token_set_ratio, limit=None)
    
    # Filter matches based on the minimum score
    filtered_matches = [match for match in matches if match[1] >= min_score]
    
    # Clear existing items in tree1
    for item in tree1.get_children():
        tree1.delete(item)
    
    # Insert matched items into tree1
    for match in filtered_matches:
        matched_row = df_csv[df_csv['Description'] == match[0]].iloc[0]
        tree1.insert("", "end", values=(matched_row['Part Number'], match[0], matched_row['Provider'], matched_row['Weighted Average Cost']))

def on_tree_select(event, tree):
    global selected_row
    selected_item = tree.selection()[0]
    selected_row = tree.item(selected_item, 'values')

def search_recommendations():
    print("Search Recommendations")

def add_to_quote(sheet, entry_text, quantity, margin):
    global selected_row
    if selected_row and quantity and margin:
        part_number = selected_row[0]
        description = selected_row[1]
        manufacturer = selected_row[2]
        cost = selected_row[3]
        # Convert margin to a percentage
        margin_percentage = float(margin) / 100
        # add to next blank row in the quote sheet
        next_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
        sheet.range('A' + str(next_row)).value = entry_text
        sheet.range('B' + str(next_row)).value = next_row - 12
        sheet.range('C' + str(next_row)).value = manufacturer
        sheet.range('D' + str(next_row)).value = part_number
        sheet.range('E' + str(next_row)).value = description
        sheet.range('F' + str(next_row)).value = quantity
        sheet.range('G' + str(next_row)).value = cost
        sheet.range('H' + str(next_row)).value = margin_percentage
        sheet.range('I' + str(next_row)).value = float(cost) * (1 + margin_percentage)
        sheet.range('J' + str(next_row)).value = float(cost) * (1 + margin_percentage) * float(quantity)
    else:
        print("Missing data")

if __name__ == "__main__":
    open_matcher()