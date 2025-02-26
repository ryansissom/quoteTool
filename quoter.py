import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
from fuzzywuzzy import fuzz, process
import tkinter.font as tkFont
import csv

# Global dictionary to store selected rows for each treeview
selected_rows = {
    "tree1": None,
    "tree2": None
}

def open_matcher():
    # Get the directory of the current script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_file = os.path.join(script_dir, 'Store Parts Inventory.csv')

    # Get the data from the csv
    df = pd.read_csv(csv_file)
    manufacturers = sorted(df['Provider'].dropna().astype(str).unique().tolist())
    store_names = sorted(df['Store Name'].dropna().astype(str).unique().tolist())

    root = tk.Tk()
    root.title("Excel Matcher")
    root.geometry("600x600")
    
    # Create a frame for the top widgets
    top_frame = ttk.Frame(root)
    top_frame.pack(fill='x', padx=10, pady=10)
    
    # Add a search entry field to the GUI
    search_label = ttk.Label(top_frame, text="Search:")
    search_label.pack(side='left')
    
    search_entry = ttk.Entry(top_frame, width=30)
    search_entry.pack(side='left', padx=10)
    
    # Add a customer description entry field to the GUI
    customer_description_label = ttk.Label(top_frame, text="Customer Description:")
    customer_description_label.pack(side='left', padx=(20, 0))
    
    customer_description_entry = ttk.Entry(top_frame, width=30)
    customer_description_entry.pack(side='left', padx=10)
    
    # Add a dropdown to select manufacturers
    manufacturer_label = ttk.Label(top_frame, text="Select Manufacturer:")
    manufacturer_label.pack(side='left', padx=(20, 0))
    
    manufacturer_var = tk.StringVar()
    manufacturer_dropdown = ttk.Combobox(top_frame, textvariable=manufacturer_var, width=30)
    manufacturer_dropdown['values'] = manufacturers
    manufacturer_dropdown.pack(side='left', padx=10)

    # Add a dropdown to select store names
    store_name_label = ttk.Label(top_frame, text="Select Store Name:")
    store_name_label.pack(side='left', padx=(20, 0))
    
    store_name_var = tk.StringVar()
    store_name_dropdown = ttk.Combobox(top_frame, textvariable=store_name_var, width=30)
    store_name_dropdown['values'] = store_names
    store_name_dropdown.pack(side='left', padx=10)

    # Create a frame to contain tree1 and its scrollbar
    tree1_frame = ttk.Frame(root)
    tree1_frame.pack(expand=True, fill='both', padx=10, pady=10)

    # Create a treeview for the first table
    columns_tree1 = ("Part Number", "Description", "Manufacturer", "Store Name", "Cost", "Quantity Available", "Lead Time")
    tree1 = ttk.Treeview(tree1_frame, columns=columns_tree1, show='headings')
    
    # Define headings for the first table
    for col in columns_tree1:
        tree1.heading(col, text=col)

    # Set column widths to fit the text size for tree1
    font = tkFont.Font()
    for col in columns_tree1:
        max_width = max(font.measure(col), 100)  # Set a minimum width of 100
        tree1.column(col, width=max_width, stretch=False)

    # Create a vertical scrollbar for tree1
    tree1_scrollbar = ttk.Scrollbar(tree1_frame, orient="vertical", command=tree1.yview)
    tree1.configure(yscrollcommand=tree1_scrollbar.set)

    # Pack the treeview and scrollbar
    tree1.pack(side='left', expand=True, fill='both')
    tree1_scrollbar.pack(side='right', fill='y')
    
    # Bind the selection event to a function
    tree1.bind("<<TreeviewSelect>>", lambda event: on_tree_select(event, tree1, "tree1"))
    
    # Bind right-click event to tree1 for copying cell text
    tree1.bind("<Button-3>", lambda event: show_context_menu(event, tree1))
    
    # Bind double-click event to tree1 for adding item to quote
    tree1.bind("<Double-1>", lambda event: add_selected_item_to_quote(event, tree1, customer_description_entry.get(), tree2))

    # Create a frame for the buttons
    button_frame = ttk.Frame(root)
    button_frame.pack(fill='x', padx=10, pady=10)

    # Add a Search button to the button frame
    search_button = ttk.Button(button_frame, text="Search", command=lambda: fuzzymatch(search_entry.get(), manufacturer_dropdown.get(), store_name_dropdown.get(), tree1))
    search_button.pack(side='left', padx=5)

    # Add a "Custom Entry" button to the button frame
    custom_entry_button = ttk.Button(button_frame, text="Custom Entry", command=lambda: open_custom_entry_window(tree2, customer_description_entry.get()))
    custom_entry_button.pack(side='left', padx=5)

    # Add a quantity entry field
    quantity_label = ttk.Label(button_frame, text="Quantity:")
    quantity_label.pack(side='left', padx=(0, 5))
    quantity_entry = ttk.Entry(button_frame, width=10)
    quantity_entry.pack(side='left', padx=(0, 20))

    # Add a margin entry field
    margin_label = ttk.Label(button_frame, text="Margin:")
    margin_label.pack(side='left', padx=(0, 5))
    margin_entry = ttk.Entry(button_frame, width=10)
    margin_entry.pack(side='left', padx=(0, 20))

    # Add an "Add to Quote" button to the button frame
    add_to_quote_button = ttk.Button(button_frame, text="Add to Quote", command=lambda: add_to_quote(customer_description_entry.get(), quantity_entry.get(), margin_entry.get(), tree2))
    add_to_quote_button.pack(side='left', padx=5)

    # Create a frame to contain tree2 and its scrollbar
    tree2_frame = ttk.Frame(root)
    tree2_frame.pack(expand=True, fill='both', padx=10, pady=10)

    # Create a treeview for the second table
    columns_tree2 = ("Customer Description", "Part Number", "Description", "Manufacturer", "Store Name", "Cost", "Quantity", "Margin", "Total")
    tree2 = ttk.Treeview(tree2_frame, columns=columns_tree2, show='headings')
    
    # Define headings for the second table
    for col in columns_tree2:
        tree2.heading(col, text=col)

    # Set column widths to fit the text size for tree2
    font = tkFont.Font()
    for col in columns_tree2:
        max_width = max(font.measure(col), 100)  # Set a minimum width of 100
        tree2.column(col, width=max_width, stretch=False)

    # Create a vertical scrollbar for tree2
    tree2_scrollbar = ttk.Scrollbar(tree2_frame, orient="vertical", command=tree2.yview)
    tree2.configure(yscrollcommand=tree2_scrollbar.set)

    # Pack the treeview and scrollbar
    tree2.pack(side='left', expand=True, fill='both')
    tree2_scrollbar.pack(side='right', fill='y')
    
    # Bind the selection event to a function
    tree2.bind("<<TreeviewSelect>>", lambda event: on_tree_select(event, tree2, "tree2"))
    
    # Bind right-click event to tree2 for copying cell text
    tree2.bind("<Button-3>", lambda event: show_context_menu(event, tree2))
    
    # Bind double-click event to tree2 for editing Customer Description and Description
    tree2.bind("<Double-1>", lambda event: edit_cell(event, tree2))

    bottom_frame = ttk.Frame(root)
    bottom_frame.pack(fill='x', padx=10, pady=10)

    # Add an "Export to CSV" button at the bottom middle
    export_button = ttk.Button(bottom_frame, text="Export to CSV", command=lambda: export_to_csv(tree2))
    export_button.pack(side='left', padx=5)

    # Add a "Delete Line" button at the bottom middle
    delete_button = ttk.Button(bottom_frame, text="Delete Line", command=lambda: delete_selected_item(tree2))
    delete_button.pack(side='left', padx=5)
    
    # Bind Enter key to search_entry to trigger search
    search_entry.bind("<Return>", lambda event: fuzzymatch(search_entry.get(), manufacturer_dropdown.get(), store_name_dropdown.get(), tree1))
    
    root.mainloop()

def show_context_menu(event, tree):
    # Create a context menu
    context_menu = tk.Menu(tree, tearoff=0)
    context_menu.add_command(label="Copy", command=lambda: copy_cell_text(tree, event))
    context_menu.post(event.x_root, event.y_root)

def copy_cell_text(tree, event):
    # Get the selected item and column
    region = tree.identify("region", event.x, event.y)
    if region == "cell":
        row_id = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        cell_value = tree.item(row_id, "values")[int(column_id[1:]) - 1]
        # Copy the cell value to the clipboard
        tree.clipboard_clear()
        tree.clipboard_append(cell_value)
        print(f"Copied to clipboard: {cell_value}")

def edit_cell(event, tree):
    # Get the selected item and column
    region = tree.identify("region", event.x, event.y)
    if region == "cell":
        row_id = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        column_index = int(column_id[1:]) - 1
        if column_index in [0, 2]:  # Allow editing of the first column (Customer Description) and third column (Description)
            cell_value = tree.item(row_id, "values")[column_index]
            entry = tk.Entry(tree)
            entry.insert(0, cell_value)
            entry.place(x=event.x, y=event.y, anchor="w")

            def save_edit(event):
                new_value = entry.get()
                values = list(tree.item(row_id, "values"))
                values[column_index] = new_value
                tree.item(row_id, values=values)
                entry.destroy()

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", lambda event: entry.destroy())
            entry.focus()

def add_selected_item_to_quote(event, tree1, customer_description, tree2):
    global selected_rows
    if selected_rows["tree1"]:
        part_number, description, manufacturer, store_name, cost, quantity_available, lead_time = selected_rows["tree1"]
        try:
            quantity = 1  # Default quantity
            margin = 0  # Default margin
            total = float(cost) * int(quantity) * (1 + float(margin) / 100)
            total_rounded = round(total, 2)
            tree2.insert("", "end", values=(customer_description, part_number, description, manufacturer, store_name, cost, quantity, margin, total_rounded))
            print(f"Added to quote: {selected_rows['tree1']} with quantity {quantity} and margin {margin}")
        except ValueError as e:
            print(f"Error calculating total: {e}")
    else:
        print("No item selected in tree1")

def open_custom_entry_window(tree2, customer_description):
    custom_entry_window = tk.Toplevel()
    custom_entry_window.title("Custom Entry")
    custom_entry_window.geometry("600x400")  # Increased size

    # Add entry fields for custom input
    customer_description_label = ttk.Label(custom_entry_window, text="Customer Description:")
    customer_description_label.pack(side='top', padx=(0, 5), pady=(5, 0))
    customer_description_entry = ttk.Entry(custom_entry_window, width=50)  # Increased width
    customer_description_entry.pack(side='top', padx=(0, 20), pady=(0, 5))
    customer_description_entry.insert(0, customer_description)  # Pre-fill with the customer description

    part_number_label = ttk.Label(custom_entry_window, text="Part Number:")
    part_number_label.pack(side='top', padx=(0, 5), pady=(5, 0))
    part_number_entry = ttk.Entry(custom_entry_window, width=50)  # Increased width
    part_number_entry.pack(side='top', padx=(0, 20), pady=(0, 5))

    description_label = ttk.Label(custom_entry_window, text="Description:")
    description_label.pack(side='top', padx=(0, 5), pady=(5, 0))
    description_entry = ttk.Entry(custom_entry_window, width=50)  # Increased width
    description_entry.pack(side='top', padx=(0, 20), pady=(0, 5))

    manufacturer_label = ttk.Label(custom_entry_window, text="Manufacturer:")
    manufacturer_label.pack(side='top', padx=(0, 5), pady=(5, 0))
    manufacturer_entry = ttk.Entry(custom_entry_window, width=50)  # Increased width
    manufacturer_entry.pack(side='top', padx=(0, 20), pady=(0, 5))

    cost_label = ttk.Label(custom_entry_window, text="Cost:")
    cost_label.pack(side='top', padx=(0, 5), pady=(5, 0))
    cost_entry = ttk.Entry(custom_entry_window, width=50)  # Increased width
    cost_entry.pack(side='top', padx=(0, 20), pady=(0, 5))

    quantity_label = ttk.Label(custom_entry_window, text="Quantity:")
    quantity_label.pack(side='top', padx=(0, 5), pady=(5, 0))
    quantity_entry = ttk.Entry(custom_entry_window, width=50)  # Increased width
    quantity_entry.pack(side='top', padx=(0, 20), pady=(0, 5))

    margin_label = ttk.Label(custom_entry_window, text="Margin:")
    margin_label.pack(side='top', padx=(0, 5), pady=(5, 0))
    margin_entry = ttk.Entry(custom_entry_window, width=50)  # Increased width
    margin_entry.pack(side='top', padx=(0, 20), pady=(0, 5))

    # Add a button to add the custom entry to tree2 and close the window
    add_button = ttk.Button(custom_entry_window, text="Add", command=lambda: add_custom_entry_and_close(customer_description_entry.get(), part_number_entry.get(), description_entry.get(), manufacturer_entry.get(), cost_entry.get(), quantity_entry.get(), margin_entry.get(), tree2, custom_entry_window))
    add_button.pack(side='top', padx=5, pady=10)

def add_custom_entry_and_close(customer_description, part_number, description, manufacturer, cost, quantity, margin, tree2, custom_entry_window):
    try:
        total = float(cost) * int(quantity) * (1 + float(margin) / 100)
        total_rounded = round(total, 2)
        tree2.insert("", "end", values=(customer_description, part_number, description, manufacturer, "", cost, quantity, margin, total_rounded))
        print(f"Custom entry added: {part_number, description, manufacturer, cost, quantity, margin, total_rounded}")
        custom_entry_window.destroy()  # Close the custom entry window
    except ValueError as e:
        print(f"Error calculating total: {e}")

def fuzzymatch(customer_desc, selected_manufacturer, selected_store_name, tree1, min_score=70):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_file = os.path.join(script_dir, 'Store Parts Inventory.csv')
    df_csv = pd.read_csv(csv_file)
    
    if selected_manufacturer:
        df_csv = df_csv[df_csv['Provider'] == selected_manufacturer]
    
    if selected_store_name:
        df_csv = df_csv[df_csv['Store Name'] == selected_store_name]
    
    matches = process.extract(customer_desc, df_csv['Description'], scorer=fuzz.token_set_ratio, limit=None)
    
    # Filter matches based on the minimum score
    filtered_matches = [match for match in matches if match[1] >= min_score]
    
    # Clear existing items in tree1
    for item in tree1.get_children():
        tree1.delete(item)
    
    # Insert matched items into tree1
    for match in filtered_matches:
        matched_row = df_csv[df_csv['Description'] == match[0]].iloc[0]
        tree1.insert("", "end", values=(matched_row['Part Number'], match[0], matched_row['Provider'], matched_row['Store Name'], matched_row['Weighted Average Cost'], matched_row['Parts Quantity'], 'data required'))

def on_tree_select(event, tree, tree_name):
    global selected_rows
    selected_items = tree.selection()
    if selected_items:
        selected_item = selected_items[0]
        selected_rows[tree_name] = tree.item(selected_item, 'values')
    else:
        selected_rows[tree_name] = None
        print("No item selected")

def search_recommendations():
    print("Search Recommendations")

def add_to_quote(customer_description, quantity, margin, tree2):
    global selected_rows
    if selected_rows["tree1"]:
        part_number, description, manufacturer, store_name, cost, quantity_available, lead_time = selected_rows["tree1"]
        try:
            total = float(cost) * int(quantity) * (1 + float(margin) / 100)
            total_rounded = round(total, 2)
            tree2.insert("", "end", values=(customer_description, part_number, description, manufacturer, store_name, cost, quantity, margin, total_rounded))
            print(f"Added to quote: {selected_rows['tree1']} with quantity {quantity} and margin {margin}")
        except ValueError as e:
            print(f"Error calculating total: {e}")
    else:
        print("No item selected in tree1")

def add_custom_entry(customer_description, part_number, description, manufacturer, cost, quantity, margin, tree2):
    try:
        total = float(cost) * int(quantity) * (1 + float(margin) / 100)
        total_rounded = round(total, 2)
        tree2.insert("", "end", values=(customer_description, part_number, description, manufacturer, "", cost, quantity, margin, total_rounded))
        print(f"Custom entry added: {part_number, description, manufacturer, cost, quantity, margin, total_rounded}")
    except ValueError as e:
        print(f"Error calculating total: {e}")

def delete_selected_item(tree2):
    selected_item = tree2.selection()
    if selected_item:
        tree2.delete(selected_item)
        print("Selected item deleted")
    else:
        print("No item selected in tree2")

def export_to_csv(tree2, columns_to_include=["Customer Description", "Part Number", "Description", "Price per Unit", "Quantity", "Total"]):
    # Open a file dialog to choose where to save the file
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
    
    if not file_path:
        return  # User cancelled the save dialog
    
    # Get all items from tree2
    items = tree2.get_children()
    # Get column names
    all_columns = [tree2.heading(col)["text"] for col in tree2["columns"]]
    
    # Open a file for writing
    with open(file_path, "w", newline="") as file:
        writer = csv.writer(file)
        # Write the column names
        writer.writerow(columns_to_include)
        # Write the data
        for item in items:
            values = tree2.item(item, "values")
            customer_description = values[0]
            part_number = values[1]
            description = values[2]
            cost = float(values[5])
            quantity = int(values[6])
            total = float(values[8])
            price_per_unit = round(total / quantity, 2) if quantity != 0 else 0
            writer.writerow([customer_description, part_number, description, price_per_unit, quantity, total])
    
    print(f"Data exported to {file_path}")

# Example usage:
# export_to_csv(tree2, columns_to_include=["Customer Description", "Part Number", "Description", "Price per Unit", "Quantity", "Total"])

if __name__ == "__main__":
    open_matcher()