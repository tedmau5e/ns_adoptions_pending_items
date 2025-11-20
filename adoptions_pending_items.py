import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import threading
import sys

cols_to_delete = ['ID', 'School', 'Academic Department Code', 'Section Code', 'Section Name', 'Class Number', 'Class Section', 'Estimated Enrollment', 'Book Cover Type', 'Course Material Requirement', 'Item', 'Item Notes', 'Previously Adopted', 'OK to Substitute', 'Internal Notes', 'Item Type', 'Adoption Status Date', 'Item Buyer Statistic']

current_date = datetime.now()
date_string = current_date.strftime("%m%d%y")

def load_file(root):
    file_path = filedialog.askopenfilename(
        title="Select input file",
        filetypes=[("Excel files", "*.xlsx, *.xls"), ("CSV files", "*.csv")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected. Exiting")
    
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, header=0)
            messagebox.showinfo("Success", "Successfully loaded Excel file.")
            return df
        elif file_path.endswith('.xls'):
            try:
                df = pd.read_xls(file_path, header=0)
                messagebox.showinfo("Success", "Successfully loaded Excel file.")
                return df
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load. Please try opening the file and saving as .xls or .xlsx, then run this script again.")
                return None
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path, header=0)
            df['ISBN'] = pd.to_numeric(df['ISBN'], errors="coerce")
            messagebox.showinfo("Success", "Successfully loaded CSV file.")
            return df
        else:
            messagebox.showerror("Error", "Unsupported file format. Please provide a valid Excel file (.xlsx).")
    except Exception as e:
        messagebox.showerror("Error", f"Error loading file: {e}")
        return None

def delete_junk_cols(df):
    # function to delete unneeded columns
    df.drop(columns=cols_to_delete, inplace=True)
    df.columns = ['ISBN', 'UPC Code', 'Item Display Name', 'Short Title', 'Long Title', 'Book Author', 'Book Publisher']
    print(f"Columns {cols_to_delete} deleted.")
    return df

def reorder_cols(df):
    # function to reorder remaining columns after deletion; must be run before rename_and_add_cols
    current_cols = df.columns.tolist()
    print(current_cols)
    new_column_order = ['Item Display Name', 'ISBN', 'UPC Code', 'Long Title', 'Short Title', 'Book Author', 'Book Publisher']
    return new_column_order

def rename_and_add_cols(df):
    # function to rename columns and add needed columns
    current_cols = df.columns.tolist()
    df1 = pd.DataFrame(df)
    df2 = pd.DataFrame()
    first_row_idx = 0
    last_row_idx = len(df)
    print(current_cols)
    df1.rename(columns={'Item Display Name': 'DISPLAYNAME', 'ISBN' : 'ISBN', 'UPC Code' : 'UPC', 'Long Title' : 'LONG_TITLE', 'Short Title' : 'SHORT_TITLE', 'Book Author' : 'AUTHOR', 'Book Publisher' : 'IMPRINT'}, inplace=True)
    df1['DISPLAYNAME'] = df1['DISPLAYNAME'].fillna(df1['LONG_TITLE'])
    new_columns = {'EXTERNAL_ID' : '', 'ITEM_NAME_NUMBER' : '', 'MATRIX_TYPE' : 'Parent Matrix Item', 'CS_PRODUCT_TYPE' : '', 'PARENT' : '', 'UNIT_TYPE' : 'Each', 'STOCK_UNIT' : 'EA', 'PURCHASE_UNIT' : 'EA', 'SALE_UNIT' : 'EA', 'DO_NOT_ALLOW_DISCOUNT' : 'T', 'SUBSIDIARY' : 'Brown University Bookstore', 'INCLUDE_CHILDREN' : 'T', 'DEPARTMENT' : 'Textbooks : TX Textbook (TXT)', 'PRODUCT_CATEGORY' : 'Course Materials Physical (CMP) : Print New', 'RENTAL_AVAILABLE' : '', 'BOOK_CONDITION' : '', 'EDITION' : 'NA', 'WEB_DISP_BEHAVIOR' : 'ATP based add and remove from web display', 'COSTING_METHOD' : 'Average', 'PREFERRED_LOCATION' : 'Main Campus Bookstore', 'COGS_ACCOUNT' : '219', 'INCOME_ACCOUNT' : '324', 'ASSET_ACCOUNT' : '218', 'TAX_SCHEDULE' : 'Taxable RI'}
    for index in range(first_row_idx, last_row_idx):
        for col, val in new_columns.items():
            df2.loc[index, col] = val
    alt_df = pd.merge(df1, df2, left_index=True, right_index=True)
    print(alt_df)
    return alt_df

def reorder_all_cols(df):
    # function to settle final column order
    current_cols = df.columns.tolist()
    print(current_cols)
    final_col_order = ['EXTERNAL_ID', 'ITEM_NAME_NUMBER', 'DISPLAYNAME', 'MATRIX_TYPE', 'CS_PRODUCT_TYPE', 'PARENT', 'ISBN', 'UPC', 'UNIT_TYPE', 'STOCK_UNIT', 'PURCHASE_UNIT', 'SALE_UNIT', 'DO_NOT_ALLOW_DISCOUNT', 'SUBSIDIARY', 'INCLUDE_CHILDREN', 'DEPARTMENT', 'PRODUCT_CATEGORY', 'RENTAL_AVAILABLE', 'BOOK_CONDITION', 'LONG_TITLE', 'SHORT_TITLE', 'AUTHOR', 'EDITION', 'IMPRINT', 'WEB_DISP_BEHAVIOR', 'COSTING_METHOD', 'PREFERRED_LOCATION', 'COGS_ACCOUNT', 'INCOME_ACCOUNT', 'ASSET_ACCOUNT', 'TAX_SCHEDULE']
    return df[final_col_order]

def fill_easy_cells(df):
    df['EXTERNAL_ID'] = df.apply(lambda row: f"{row['ISBN']}{'parent'}", axis=1)
    print(f"{df['EXTERNAL_ID']}")
    df['ITEM_NAME_NUMBER'] = df.apply(lambda row: f"{row['DISPLAYNAME']}-{row['ISBN']}", axis=1)
    print(f"{df['ITEM_NAME_NUMBER']}")

def label_dupes(row):
    if row['occurrence'] == 1:
        row['EXTERNAL_ID'] = row['EXTERNAL_ID'].replace('parent', 'new')
        row['ITEM_NAME_NUMBER'] = row['ITEM_NAME_NUMBER'] + '-New'
        row['MATRIX_TYPE'] = 'Child Matrix Type'
        row['CS_PRODUCT_TYPE'] = 'New'
        row['PARENT'] = str(row['ISBN']) + 'parent'
        row['UPC'] = row['ISBN']
        row['INCLUDE_CHILDREN'] = 'F'
        row['RENTAL_AVAILABLE'] = 'F'
        row['BOOK_CONDITION'] = 'New'
    elif row['occurrence'] == 2:
        row['EXTERNAL_ID'] = row['EXTERNAL_ID'].replace('parent', 'used')
        row['ITEM_NAME_NUMBER'] = row['ITEM_NAME_NUMBER'] + '-Used'
        row['MATRIX_TYPE'] = 'Child Matrix Type'
        row['CS_PRODUCT_TYPE'] = 'Used'
        row['PARENT'] = str(row['ISBN']) + 'parent'
        row['UPC'] = str(row['ISBN']).replace('978', '290')
        row['INCLUDE_CHILDREN'] = 'F'
        row['PRODUCT_CATEGORY'] = 'Course Materials Physical (CMP) : Print Used'
        row['RENTAL_AVAILABLE'] = 'F'
        row['BOOK_CONDITION'] = 'Used'
        row['TAX_SCHEDULE'] = 'Not Taxable'
    return row

def make_children(df):
    repeat_index = df.index.repeat(3)
    df_duplicated = df.loc[repeat_index].reset_index(drop=True)

    df_duplicated['original_index'] = repeat_index.to_numpy() # array([[External ID value 1, External ID value 2], [Item Name Number Value 1, Item Name Number Value 2]])

    df_duplicated['occurrence'] = df_duplicated.groupby('original_index').cumcount() # adds 'occurence' column and counts from 0 to last occurence - 1 per row grouping
    current_cols = df.columns.tolist()
    print(current_cols)
    
    df_duplicated = df_duplicated.apply(label_dupes, axis=1) # use func to apply values in each row in given column based on conditional in func
    print(df_duplicated)

    df_duplicated = df_duplicated.drop(columns=['occurrence', 'original_index']) # remove 'occurrence' and 'original_index' columns
    return df_duplicated

def loading_popup(root, task_functions, title="Working", message="Please wait..."):
    popup = tk.Toplevel(root)
    popup.title(title)
    tk.Label(popup, text=message).pack(padx=20, pady=10)

    progress_bar = ttk.Progressbar(popup, mode='indeterminate', length=200)
    progress_bar.pack(padx=20, pady=10)
    progress_bar.start(10)

    def task_wrapper():
        change_books(*task_functions)
        root.after(1000, popup.destroy)
    
    worker_thread = threading.Thread(target=task_wrapper, daemon=True)
    worker_thread.start()

def change_books(load_file, delete_junk_cols, reorder_cols, rename_and_add_cols, reorder_all_cols, fill_easy_cells, make_children):
    try:
        df = load_file(root)
        if df is None:
            return
        
        delete_junk_cols(df)
        
        new_column_order = reorder_cols(df)
        try:
            df = df[new_column_order]
        except KeyError as e:
            messagebox.showerror("Error", f"Error reordering columns: {e}")
            return
        
        df = rename_and_add_cols(df)

        df = reorder_all_cols(df)
        
        fill_easy_cells(df)
        print(df)
        df = make_children(df)
        print(df)

        save_directory = filedialog.askdirectory(parent=root, title="Select Save Location")
        default_name = f'Adptd_Text_Items-{date_string}.xlsx'
        full_save_path = os.path.join(save_directory, default_name)

        with pd.ExcelWriter(full_save_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=default_name, index=False)
            format_cols = ['ISBN', 'UPC']
            for col_name in format_cols:
                col_idx = df.columns.get_loc(col_name) + 1
                for row in writer.sheets[default_name].iter_rows(min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = '0'
        messagebox.showinfo("Success", f"Processed file saved to {full_save_path}")
        root.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occured: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Loading...")

    loading_popup(root, [load_file, delete_junk_cols, reorder_cols, rename_and_add_cols, reorder_all_cols, fill_easy_cells, make_children])

    root.withdraw()

    root.mainloop()