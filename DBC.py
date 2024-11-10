import tkinter as tk
from tkinter import filedialog, messagebox, Menu, Toplevel, Checkbutton, IntVar, Frame, Label, Entry, Canvas
from tkinter import ttk
import pandas as pd
from pandas import ExcelWriter
from datetime import datetime


class GUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("File Combiner")
        self.geometry("620x250")
        self.resizable(False, False)
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.new_mandatory = tk.StringVar()
        self.mandatory_columns = self.load_mandatory_columns()
        self.reference_column = None
        # Main frame
        main_frame = ttk.Frame(self, padding="10 10 10 10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Menu bar
        my_menu = Menu(self)
        self.config(menu=my_menu)
        # Create a file menu item
        file_menu = Menu(my_menu, tearoff=False)
        my_menu.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Exit", command=self.quit)
        # Create an option menu item
        option_menu = Menu(my_menu, tearoff=False)
        my_menu.add_cascade(label="Option", menu=option_menu)
        option_menu.add_command(label="Settings", command=self.open_option_window)
        # File 1
        ttk.Label(main_frame, text="Old file (Excel/CSV) :").grid(row=0, column=0, pady=0, padx=5, sticky="w")
        ttk.Entry(main_frame, textvariable=self.file1_path, width=70).grid(row=1, column=0, pady=0, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.select_file1).grid(row=1, column=1, pady=5, padx=5)
        # File 2
        ttk.Label(main_frame, text="New file (Excel/CSV) :").grid(row=2, column=0, pady=0, padx=5, sticky="w")
        ttk.Entry(main_frame, textvariable=self.file2_path, width=70).grid(row=3, column=0, pady=0, padx=5)
        ttk.Button(main_frame, text="Browse...", command=self.select_file2).grid(row=3, column=1, pady=5, padx=5)
        # Output buttons
        ttk.Button(main_frame, text="START", command=self.select_reference_column).grid(row=4, column=1, pady=(20, 0),
                                                                                        padx=(0, 5), sticky="E")
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate', maximum=100)
        self.progress.grid(row=5, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))
        # Separator
        separator_line = tk.Canvas(main_frame, height=1, bg="#D4D4D4")
        separator_line.grid(row=6, column=0, columnspan=6, sticky="ew", pady=0)
        # Global info
        ttk.Label(main_frame, text="").grid(row=7, column=0, columnspan=3, padx=2, pady=0,
                                                                 sticky="NSW")
        ttk.Label(main_frame, text="v1.1").grid(row=7, column=1, columnspan=2, padx=5, pady=0, sticky="NSE")

    def select_file1(self):
        self.file1_path.set(filedialog.askopenfilename(title="Select File 1", filetypes=[("All files", "*.*")]))

    def select_file2(self):
        self.file2_path.set(filedialog.askopenfilename(title="Select File 2", filetypes=[("All files", "*.*")]))

    def get_file_paths(self):
        return self.file1_path.get(), self.file2_path.get()

    def read_file(self, file_path):
        try:
            if file_path.endswith('.xlsx'):
                return pd.read_excel(file_path)
            else:
                return pd.read_csv(file_path, sep=';', encoding='utf-8')
        except UnicodeDecodeError:
            try:
                return pd.read_csv(file_path, sep=';', encoding='latin1')
            except Exception as e:
                raise e

    def select_reference_column(self, col=None):
        old_file, new_file = self.get_file_paths()
        if not old_file or not new_file:
            messagebox.showerror("Error", "Please select both files.")
            return
        try:
            self.old_df = self.read_file(old_file)
            self.new_df = self.read_file(new_file)
            all_columns = set(self.old_df.columns).union(set(self.new_df.columns))
            self.reference_selection_window = Toplevel(self)
            self.reference_selection_window.title("Select Reference Column")

            # Calculate and set the height based on the number of checkbuttons
            num_rows = len(all_columns) + 2  # +2 for the label and button
            row_height = 30  # Approximate height of each row
            height = num_rows * row_height
            self.reference_selection_window.geometry(f"300x{height}")

            reference_frame = Frame(self.reference_selection_window)
            reference_frame.pack(pady=10, padx=10, anchor='w')

            Label(reference_frame, text="Select the reference column:").pack(anchor=tk.W)

            self.reference_vars = {}
            for col in all_columns:
                var = IntVar()
                chk = Checkbutton(reference_frame, text=col, variable=var,
                                  command=lambda col=col: self.set_reference_column(col))
                chk.pack(anchor=tk.W)
                self.reference_vars[col] = var

            ttk.Button(self.reference_selection_window, text="Next", command=self.show_column_selection).pack(pady=(5,0),padx =10,
                                                                                                              anchor='e')

        except Exception as e:
            messagebox.showerror("Error", f"Error reading files: {e}")

    def set_reference_column(self, col):
        for other_col in self.reference_vars:
            if other_col != col:
                self.reference_vars[other_col].set(0)
        self.reference_column = col

    def show_column_selection(self):
        if not self.reference_column:
            messagebox.showerror("Error", "Please select a reference column.")
            return
        self.reference_selection_window.destroy()
        all_columns = set(self.old_df.columns).union(set(self.new_df.columns))
        self.column_selection_window = Toplevel(self)
        self.column_selection_window.title("Select Mandatory Fields")

        # Calculate and set the height based on the number of checkbuttons
        num_rows = len(all_columns) + 2  # +2 for the label and button
        row_height = 30  # Approximate height of each row
        height = num_rows * row_height
        self.column_selection_window.geometry(f"300x{height}")

        self.column_vars = {}
        column_frame = Frame(self.column_selection_window)
        column_frame.pack(pady=10, padx=10, anchor='w')

        Label(column_frame, text="Select Mandatory Fields :").pack(anchor=tk.W)

        for col in all_columns:
            var = IntVar()
            color = "green" if col in self.mandatory_columns else "red"
            symbol = "✔" if col in self.mandatory_columns else "✖"
            chk = Checkbutton(column_frame, text=f"{symbol} {col}", variable=var, fg=color)
            chk.pack(anchor=tk.W)
            self.column_vars[col] = var

        ttk.Button(self.column_selection_window, text="Compare", command=self.set_mandatory_columns).pack(pady=(5,0),padx =10,
                                                                                                              anchor='e')

    def set_mandatory_columns(self):
        self.mandatory_columns = [col for col, var in self.column_vars.items() if var.get() == 1]
        self.column_selection_window.destroy()
        self.progress.start()
        self.after(1000, self.compare_files_step_1)

    def compare_files_step_1(self):
        old_df = self.old_df
        new_df = self.new_df
        ref_col = self.reference_column
        if ref_col not in old_df.columns or ref_col not in new_df.columns:
            messagebox.showerror("Error", f"The reference column '{ref_col}' does not exist in both files.")
            self.progress.stop()
            return
        self.missing = old_df[~old_df[ref_col].isin(new_df[ref_col])]
        self.added = new_df[~new_df[ref_col].isin(old_df[ref_col])]
        self.merged = pd.merge(old_df.set_index(ref_col), new_df.set_index(ref_col), on=ref_col,
                               suffixes=('_old', '_new'),
                               how='outer')
        self.progress['value'] = 20
        self.after(1000, self.compare_files_step_2)

    def compare_files_step_2(self):
        edited_rows = []
        for index, row in self.merged.iterrows():
            if index in self.missing[self.reference_column].values or index in self.added[self.reference_column].values:
                continue  # Skip if the row is in both missing and added
            diff_cols = []
            for col in self.mandatory_columns:
                if col != self.reference_column:
                    old_col = f"{col}_old"
                    new_col = f"{col}_new"
                    old_value = row[old_col] if old_col in row else None
                    new_value = row[new_col] if new_col in row else None
                    if pd.isna(old_value) and pd.isna(new_value):
                        continue  # Treat both NaN values as equal
                    elif old_value != new_value:
                        diff_cols.append({'Column': col, 'Old Value': old_value, 'New Value': new_value})
            if diff_cols:
                edited_rows.append({'Index': index, 'Differences': diff_cols})
        self.edited = pd.DataFrame(edited_rows)
        self.progress['value'] = 50
        self.after(1000, self.compare_files_step_3)

    def compare_files_step_3(self):
        mandatory_check = pd.DataFrame(columns=['Mandatory Field', 'Old Version (N)', 'New Version (N+1)'])
        all_columns = set(self.old_df.columns).union(set(self.new_df.columns))
        for col in all_columns:
            if col != self.reference_column:
                old_exists = 'Exist' if col in self.old_df.columns else "Doesn't Exist"
                new_exists = 'Exist' if col in self.new_df.columns else "Doesn't Exist"
                mandatory_check = pd.concat([
                    mandatory_check,
                    pd.DataFrame([{
                        'Mandatory Field': col,
                        'Old Version (N)': old_exists,
                        'New Version (N+1)': new_exists
                    }])
                ], ignore_index=True)
        self.mandatory_check = mandatory_check
        self.progress['value'] = 80
        self.after(1000, self.save_results)

    def save_results(self):
        # Get the current date and time
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d_%H%M%S")
        # Create a file name with the current date and time
        output_filename = f"Comparison_{timestamp}.xlsx"
        with ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            self.missing.to_excel(writer, sheet_name='Missing', index=False)
            self.added.to_excel(writer, sheet_name='Added', index=False)
            self.edited.to_excel(writer, sheet_name='Edited', index=False)
            self.mandatory_check.to_excel(writer, sheet_name='Mandatory Fields Check', index=False)
        self.progress.stop()
        messagebox.showinfo("Success", f"Dataframes exported to {output_filename}")

    def open_option_window(self):
        self.option_window = Toplevel(self)
        self.option_window.title("Option Settings")
        self.option_window.geometry("220x400")

        option_frame = ttk.Frame(self.option_window, padding="10 10 10 10")
        option_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        Label(option_frame, text="New Mandatory :", font=('Arial', 10)).grid(row=1, column=0, pady=(5, 5), sticky=tk.W)
        ttk.Entry(option_frame, textvariable=self.new_mandatory, width=25).grid(row=2, column=0, columnspan=2,
                                                                                pady=(0, 5),
                                                                                padx=(0, 5), sticky="ew")

        ttk.Button(option_frame, text="Add", command=self.add_mandatory_column).grid(row=3, column=0, columnspan=2,
                                                                                     pady=(3, 0), padx=(0, 5),
                                                                                     sticky="ew")
        ttk.Button(option_frame, text="Delete", command=self.delete_mandatory_column).grid(row=4, column=0,
                                                                                           columnspan=2,
                                                                                           pady=(0, 3), padx=(0, 5),
                                                                                           sticky="ew")


        # Use Treeview instead of Listbox
        self.mandatory_treeview = ttk.Treeview(option_frame, columns=("Column"), show="headings", selectmode="extended",
                                               height=11)
        self.mandatory_treeview.heading("Column", text="Current Mandatory Columns")
        self.mandatory_treeview.grid(row=6, column=0, columnspan=2, pady=(10, 10), padx=(0, 5), sticky=(tk.W, tk.E))

        self.update_mandatory_listbox()

    def update_mandatory_listbox(self):
        for item in self.mandatory_treeview.get_children():
            self.mandatory_treeview.delete(item)
        for col in self.mandatory_columns:
            self.mandatory_treeview.insert("", "end", values=(col,))

    def add_mandatory_column(self):
        new_column = self.new_mandatory.get()
        if new_column and new_column not in self.mandatory_columns:
            self.mandatory_columns.append(new_column)
            self.save_mandatory_columns()
            self.update_mandatory_listbox()

    def delete_mandatory_column(self):
        selected_items = self.mandatory_treeview.selection()
        for item in selected_items:
            col = self.mandatory_treeview.item(item, "values")[0]
            if col in self.mandatory_columns:
                self.mandatory_columns.remove(col)
        self.save_mandatory_columns()
        self.update_mandatory_listbox()

    def load_mandatory_columns(self):
        try:
            with open('mandatory_columns.txt', 'r') as file:
                return [line.strip() for line in file]
        except FileNotFoundError:
            return []

    def save_mandatory_columns(self):
        with open('mandatory_columns.txt', 'w') as file:
            for col in self.mandatory_columns:
                file.write(f"{col}\n")


app = GUI()
app.mainloop()
