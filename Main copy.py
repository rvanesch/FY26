import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

class DataApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Loader")
        self.root.geometry("1200x700")
        
        # UI Constants and Memory
        self.SIDE_WIDTH = 200
        self.original_orders_df = pd.DataFrame() 
        self.file_loaded = False 

        # Define which columns should be treated as currency
        self.currency_cols = [
            "Invoice Value (no tax)", 
            "Local currency Item Price (no tax) USD", 
            "Order Value (no tax) USD"
        ]

        # --- Menu Bar ---
        self.menu_bar = tk.Menu(self.root)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.file_menu.add_command(label="Open Codes File", command=self.open_codes_file)
        self.file_menu.add_command(label="Open Orders File", command=self.open_orders_file)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Quit", command=root.quit)
        self.menu_bar.add_cascade(label="File Menu", menu=self.file_menu)
        self.root.config(menu=self.menu_bar)

        # --- Main Layout ---
        self.main_container = tk.Frame(self.root)
        self.main_container.pack(side="top", fill="both", expand=True)

        # Side Frame
        self.side_frame = tk.LabelFrame(self.main_container, text="Side Frame", width=self.SIDE_WIDTH)
        self.side_frame.pack(side="left", fill="y", padx=5, pady=5)
        self.side_frame.pack_propagate(False)
        
        self.code_listbox = tk.Listbox(self.side_frame, selectmode="extended", exportselection=False)
        self.code_listbox.pack(fill="both", expand=True, padx=5, pady=5)

        # Data Frame
        self.data_frame = tk.LabelFrame(self.main_container, text="Data Frame")
        self.data_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        # --- Status Bar ---
        self.status_bar = tk.Frame(self.root, height=40, bd=1, relief="flat")
        self.status_bar.pack(side="bottom", fill="x")

        self.status_left = tk.Frame(self.status_bar, borderwidth=1, relief="solid", width=self.SIDE_WIDTH)
        self.status_left.pack(side="left", fill="y", padx=(5, 0), pady=5)
        self.status_left.pack_propagate(False)

        self.status_center = tk.Label(self.status_bar, text="Status Center", borderwidth=1, relief="solid")
        self.status_center.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        self.status_right = tk.Frame(self.status_bar, borderwidth=1, relief="solid", width=250)
        self.status_right.pack(side="right", fill="y", padx=(0, 5), pady=5)
        self.status_right.pack_propagate(False)
        
        self.row_count_label = tk.Label(self.status_right, text="Rows: 0", font=("Arial", 9, "bold"))
        self.row_count_label.pack(side="right", padx=10)

        # Buttons
        self.btn_cancel = tk.Button(self.status_right, text="Cancel", command=self.clear_data_view)
        self.btn_load = tk.Button(self.status_right, text="Load Codes", state="disabled", command=self.load_selected_codes)
        self.btn_clear_side = tk.Button(self.status_left, text="Clear", command=self.reset_orders_view)
        self.btn_filter = tk.Button(self.status_left, text="Filter", command=self.filter_orders)

        # --- Treeview with Dual Scrollbars ---
        self.tree_container = tk.Frame(self.data_frame)
        self.tree_container.pack(fill="both", expand=True)

        self.v_scroll = ttk.Scrollbar(self.tree_container, orient="vertical")
        self.v_scroll.pack(side="right", fill="y")
        self.h_scroll = ttk.Scrollbar(self.data_frame, orient="horizontal")
        self.h_scroll.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(self.tree_container, selectmode="extended", 
                                 yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)

        self.v_scroll.config(command=self.tree.yview)
        self.h_scroll.config(command=self.tree.xview)
        self.tree.bind("<<TreeviewSelect>>", self.check_selection)

    def open_codes_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.hide_left_buttons()
                self.display_data(df)
                self.file_loaded = True 
                self.status_center.config(text=f"Codes File Loaded")
                self.btn_cancel.pack(side="left", padx=5)
                self.btn_load.pack(side="left", padx=5)
            except Exception as e:
                messagebox.showerror("Error", f"Could not read file: {e}")

    def open_orders_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.hide_right_buttons()
                self.file_loaded = False 
                df = pd.read_excel(file_path)
                
                required_cols = [
                    "HPE Order #", "Purchase Order No", "Opportunity ID", "HPE Quote Number", 
                    "Customer Name (Sold To Name)", "Order Entry Date", "Product Number", 
                    "Product Description", "Product Line Code", "Ordered Quantity", 
                    "OptionDescription", "Invoice Value (no tax)", 
                    "Local currency Item Price (no tax) USD", "Order Value (no tax) USD"
                ]
                
                available_cols = [col for col in required_cols if col in df.columns]
                if not available_cols:
                    messagebox.showwarning("Column Missing", "No required Order columns found.")
                    return
                
                self.original_orders_df = df[available_cols] 
                self.display_data(self.original_orders_df)
                self.status_center.config(text=f"Orders: {file_path.split('/')[-1]}")
                self.btn_clear_side.pack(side="left", expand=True, padx=2)
                self.btn_filter.pack(side="left", expand=True, padx=2)
                
            except Exception as e:
                messagebox.showerror("Error", f"Could not read file: {e}")

    def format_as_currency(self, value):
        """Helper to convert numbers to $ currency strings."""
        try:
            if pd.isna(value) or value == "":
                return ""
            num = float(value)
            return f"${num:,.2f}"
        except (ValueError, TypeError):
            return value # Return as-is if it's not a number

    def display_data(self, df):
        self.tree.delete(*self.tree.get_children())
        cols = list(df.columns)
        self.tree["columns"] = cols
        self.tree["show"] = "headings"
        
        for col in cols:
            self.tree.heading(col, text=col)
            # Currency columns get right-aligned
            alignment = "e" if col in self.currency_cols else "w"
            self.tree.column(col, width=180, stretch=False, anchor=alignment)
        
        for _, row in df.iterrows():
            formatted_vals = []
            for col_name, val in row.items():
                if col_name in self.currency_cols:
                    formatted_vals.append(self.format_as_currency(val))
                elif pd.isna(val):
                    formatted_vals.append("")
                else:
                    formatted_vals.append(val)
            self.tree.insert("", "end", values=formatted_vals)
        
        self.row_count_label.config(text=f"Rows: {len(df)}")

    def filter_orders(self):
        selected_indices = self.code_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("Filter", "Please select items in the Side Frame first.")
            return

        selected_codes = [str(self.code_listbox.get(i)) for i in selected_indices]
        if "Product Line Code" not in self.original_orders_df.columns:
            messagebox.showerror("Error", "Required column 'Product Line Code' not found.")
            return

        filtered_df = self.original_orders_df[self.original_orders_df["Product Line Code"].astype(str).isin(selected_codes)]
        self.display_data(filtered_df)
        self.status_center.config(text="View Filtered")

    def reset_orders_view(self):
        self.code_listbox.selection_clear(0, tk.END)
        if not self.original_orders_df.empty:
            self.display_data(self.original_orders_df)
            self.status_center.config(text="Filters reset")

    def check_selection(self, event=None):
        if self.file_loaded and len(self.tree.selection()) > 0:
            self.btn_load.config(state="normal")
        else:
            self.btn_load.config(state="disabled")

    def hide_right_buttons(self):
        self.btn_load.pack_forget()
        self.btn_cancel.pack_forget()

    def hide_left_buttons(self):
        self.btn_clear_side.pack_forget()
        self.btn_filter.pack_forget()

    def clear_data_view(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []
        self.file_loaded = False
        self.hide_right_buttons()
        self.row_count_label.config(text="Rows: 0")

    def load_selected_codes(self):
        selected_items = self.tree.selection()
        cols = self.tree["columns"]
        try:
            col_list_lower = [c.lower() for c in cols]
            code_idx = col_list_lower.index("code")
        except ValueError:
            messagebox.showerror("Error", "No 'code' column found.")
            return

        self.code_listbox.delete(0, tk.END)
        for item_id in selected_items:
            row_values = self.tree.item(item_id)["values"]
            code_val = row_values[code_idx]
            self.code_listbox.insert(tk.END, code_val)
        self.clear_data_view()

if __name__ == "__main__":
    root = tk.Tk()
    app = DataApp(root)
    root.mainloop()