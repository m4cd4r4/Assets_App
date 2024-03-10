import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
import pandas as pd
import datetime
import openpyxl

# Application setup
class InventoryApp:
    def __init__(self, master):
        self.master = master
        master.title("Inventory Management")

        # Loading the workbook
        self.excel_file = "EUC_Perth_Assets.xlsx"
        
        # Create frame for buttons
        self.button_frame = ttk.Frame(master)
        self.button_frame.pack(fill='x', padx=5, pady=5)

        # Basement 4.2 button
        self.basement_btn = ttk.Button(self.button_frame, text="Basement 4.2", command=lambda: self.load_inventory_items('4.2_Items'))
        self.basement_btn.pack(side='left', padx=5)
        
        # Build Room button
        self.build_room_btn = ttk.Button(self.button_frame, text="Build Room", command=lambda: self.load_inventory_items('BR_Items'))
        self.build_room_btn.pack(side='left', padx=5)
        
        # '-' button
        self.subtract_btn = ttk.Button(self.button_frame, text="-", command=lambda: self.update_inventory("-"))
        self.subtract_btn.pack(side='left', padx=5)

        # Quantity input
        self.quantity_var = tk.StringVar()
        self.quantity_entry = ttk.Entry(self.button_frame, textvariable=self.quantity_var, width=5)
        self.quantity_entry.pack(side='left', padx=5)

        # '+' button
        self.add_btn = ttk.Button(self.button_frame, text="+", command=lambda: self.update_inventory("+"))
        self.add_btn.pack(side='left', padx=5)
        
        # Treeview for items
        self.tree_frame = ttk.Frame(master)
        self.tree_frame.pack(fill='both', expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(self.tree_frame, columns=("Item", "LastCount", "NewCount"), show='headings')
        self.tree.heading('Item', text='Item')
        self.tree.heading('LastCount', text='LastCount')
        self.tree.heading('NewCount', text='NewCount')
        self.tree.pack(fill='both', expand=True)

    def load_inventory_items(self, sheet_name):
        # Clear existing items from tree
        for item in self.tree.get_children():
            self.tree.delete(item)
       
        try:
            # Load items from specified sheet into the treeview
            df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
            
            # Print DataFrame (for debugging)
            print(df)
            
            for index, row in df.iterrows():
                # Debug: Print each row to be inserted
                print(row['Item'], row['LastCount'], row['NewCount'])
                
                # Ensure data types are correct, you might need to convert them explicitly if they are not
                self.tree.insert('', tk.END, values=(str(row['Item']), int(row['LastCount']), int(row['NewCount'])))
        except Exception as e:
            print(f"Error reading Excel file: {e}")

    def update_inventory(self, operation):
        # Placeholder for updating inventory based on operation
        pass

    # Method to log changes to Excel file will be implemented here

# Tkinter main loop
def main():
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()

# Running the app
main()