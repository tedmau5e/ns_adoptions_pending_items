import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import sys
import threading
from dotenv import load_dotenv
import requests as req
from requests.exceptions import HTTPError
from PIL import Image
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path
import numpy as np
import time
import unicodedata

# def resource_path(relative_path):
#     try:
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.abspath(".")
#     return os.path.join(base_path, relative_path)


# dotenv_path = resource_path("ISBNdb_API_Key.env")

load_dotenv(dotenv_path="./ISBNdb_API_Key.env")
# load_dotenv(dotenv_path=dotenv_path)
api_key = os.getenv("API_KEY")
print(api_key)

cols_to_delete = [
    "ID",
    "School",
    "Academic Department Code",
    "Section Code",
    "Section Name",
    "Class Number",
    "Class Section",
    "Estimated Enrollment",
    "Book Cover Type",
    "Course Material Requirement",
    "Item",
    "Item Notes",
    "Previously Adopted",
    "OK to Substitute",
    "Internal Notes",
    "Item Type",
    "Adoption Status Date",
    "Item Buyer Statistic",
]

NS = {
    "ss": "urn:schemas-microsoft-com:office:spreadsheet"
}  # define namespace for .xml file parsing

current_date = datetime.now()
date_string = current_date.strftime("%m%d%y")

# define location for image download folders
home_directory = os.path.expanduser("~")
desktop_path = os.path.join(home_directory, "Desktop")
dl_folder_name = f"book_covers-{date_string}"
dl_folder_home = os.path.join(desktop_path, dl_folder_name)
resized_images = f"{date_string}-resized_images"


def load_file(root):  # function to load an Excel (.xlsx or .xls) or CSV (.csv) file.
    file_path = None
    while file_path is None:
        file_path = filedialog.askopenfilename(
            title="Select input file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")],
        )
        if not file_path:
            retry_dialog_result = messagebox.askretrycancel(
                "No file selected", "Please select a file to proceed."
            )  # messagebox to prompt user to select file if not selected, or cancel script
            if retry_dialog_result:
                continue
            else:
                return None

        try:
            if file_path.endswith(".xlsx"):
                df = pd.read_excel(file_path, header=0)
                messagebox.showinfo("Success", "Successfully loaded Excel file.")
                return df
            elif file_path.endswith(".xls"):
                try:  # attempt to load .xls file
                    df = pd.read_excel(file_path, header=0, engine="xlrd")
                    messagebox.showinfo("Success", "Successfully loaded Excel file.")
                    return df
                except Exception as e:
                    conversion_attempt_box = messagebox.askquestion(
                        "Error", f"Failed to load. Attempt converting to .xlsx?"
                    )  # prompt to convert .xls file to .xlsx if unsuccessful
                    if conversion_attempt_box == "yes":
                        try:
                            df = parse_xml_data_to_df(file_path)
                            messagebox.showinfo("Success", "Successfully loaded file.")
                            return df
                        except Exception as e:
                            retry_dialog_result = messagebox.askretrycancel(
                                "Error",
                                f"An error occurred while reading the file: {e}. Try again?",
                            )
                            if (
                                not retry_dialog_result
                            ):  # prompt user to try again, or cancel
                                break
                            file_path = None
                    else:
                        break
            elif file_path.endswith(".csv"):
                df = pd.read_csv(file_path, header=0)
                messagebox.showinfo("Success", "Successfully loaded CSV file.")
                return df
            else:
                retry_dialog_result = messagebox.askretrycancel(
                    "Error",
                    "Unsupported file format. Please provide a valid Excel (.xlsx or .xls) or CSV (.csv) file.",
                )  # prompt user to re-select file
                if not retry_dialog_result:
                    break
                file_path = None
        except FileNotFoundError:
            retry_dialog_result = messagebox.askretrycancel(
                "Error", f"File not found: {file_path}. Select a different file?"
            )
            if not retry_dialog_result:
                break
            file_path = None
        except Exception as e:
            retry_dialog_result = messagebox.askretrycancel(
                "Error", f"An error occurred while reading the file: {e}. Try again?"
            )
            if not retry_dialog_result:  # prompt user to re-select file
                break
            file_path = None
    return None


def parse_xml_data_to_df(file_path):  # function to parse .xml data to .xlsx dataframe
    tree = ET.parse(file_path)
    root = tree.getroot()

    data = []
    columns = []

    table = root.find(".//ss:Table", NS)

    if table is None:
        raise ValueError("could not find 'Table' element in sheet.")

    header_row = table.find("ss:Row", NS)
    if header_row is not None:
        for cell in header_row.findall("ss:Cell", NS):
            data_element = cell.find("ss:Data", NS)
            if data_element is not None:
                columns.append(data_element.text)

    for row in table.findall("ss:Row", NS)[1:]:
        row_data = []
        for cell in row.findall("ss:Cell", NS):
            data_element = cell.find("ss:Data", NS)
            if data_element is not None:
                row_data.append(data_element.text)
            else:
                row_data.append(None)
        data.append(row_data)

    df = pd.DataFrame(data, columns=columns, index=None)
    excel_file_name = Path(file_path).stem + ".xlsx"
    excel_save_location = os.path.join(desktop_path, excel_file_name)
    df.to_excel(excel_save_location, index=False)
    print(df)
    return df


def delete_junk_cols(df):
    # function to delete unneeded columns/keep required original columns
    df = df[
        [
            "ISBN",
            "UPC Code",
            "Item Display Name",
            "Short Title",
            "Long Title",
            "Book Author",
            "Book Publisher",
        ]
    ]
    print(f"Columns {cols_to_delete} deleted.")
    return df


def reorder_cols(df):
    # function to reorder remaining columns after deletion; must be run before rename_and_add_cols
    current_cols = df.columns.tolist()
    print(current_cols)
    new_column_order = [
        "Item Display Name",
        "ISBN",
        "UPC Code",
        "Long Title",
        "Short Title",
        "Book Author",
        "Book Publisher",
    ]
    return new_column_order


def rename_and_add_cols(df):
    # function to rename columns and add needed columns
    current_cols = df.columns.tolist()
    df1 = pd.DataFrame(df)  # original dataframe
    df2 = pd.DataFrame()  # empty dataframe to hold new columns and data
    first_row_idx = 0
    last_row_idx = len(df)
    print(current_cols)
    df1.rename(
        columns={
            "Item Display Name": "DISPLAYNAME",
            "ISBN": "ISBN",
            "UPC Code": "UPC",
            "Long Title": "LONG_TITLE",
            "Short Title": "SHORT_TITLE",
            "Book Author": "AUTHOR",
            "Book Publisher": "IMPRINT",
        },
        inplace=True,
    )  # renames existing columns inplace
    df1["DISPLAYNAME"] = df1["DISPLAYNAME"].fillna(
        df1["LONG_TITLE"]
    )  # assigns 'DISPLAYNAME' field contents to match 'LONG_TITLE'
    new_columns = {
        "EXTERNAL_ID": "",
        "ITEM_NAME_NUMBER": "",
        "MATRIX_TYPE": "Parent Matrix Item",
        "CS_PRODUCT_TYPE": "",
        "PARENT": "",
        "UNIT_TYPE": "Each",
        "STOCK_UNIT": "EA",
        "PURCHASE_UNIT": "EA",
        "SALE_UNIT": "EA",
        "DO_NOT_ALLOW_DISCOUNT": "T",
        "SUBSIDIARY": "Brown University Bookstore",
        "INCLUDE_CHILDREN": "T",
        "DEPARTMENT": "Textbooks : TX Textbook (TXT)",
        "PRODUCT_CATEGORY": "Course Materials Physical (CMP) : Print New",
        "RENTAL_AVAILABLE": "",
        "BOOK_CONDITION": "",
        "EDITION": "NA",
        "WEB_DISP_BEHAVIOR": "ATP based add and remove from web display",
        "COSTING_METHOD": "Average",
        "PREFERRED_LOCATION": "Main Campus Bookstore",
        "COGS_ACCOUNT": "219",
        "INCOME_ACCOUNT": "324",
        "ASSET_ACCOUNT": "218",
        "TAX_SCHEDULE": "Taxable RI",
        "Base Price": "",
        "Webstore Image Name": "",
    }  # defines new columns and default values
    for index in range(
        first_row_idx, last_row_idx
    ):  # loop to go through every row (index) with data in dataframe
        for col, val in new_columns.items():
            df2.loc[index, col] = (
                val  # loop to assign val to cells for each column per row (index)
            )
    alt_df = pd.merge(
        df1, df2, left_index=True, right_index=True
    )  # combine dataframe 1 (original data) and dataframe 2 (new columns and data)
    print(alt_df)
    return alt_df


def reorder_all_cols(df):
    # function to settle final column order
    current_cols = df.columns.tolist()
    print(current_cols)
    final_col_order = [
        "EXTERNAL_ID",
        "ITEM_NAME_NUMBER",
        "DISPLAYNAME",
        "MATRIX_TYPE",
        "CS_PRODUCT_TYPE",
        "PARENT",
        "ISBN",
        "UPC",
        "UNIT_TYPE",
        "STOCK_UNIT",
        "PURCHASE_UNIT",
        "SALE_UNIT",
        "DO_NOT_ALLOW_DISCOUNT",
        "SUBSIDIARY",
        "INCLUDE_CHILDREN",
        "DEPARTMENT",
        "PRODUCT_CATEGORY",
        "RENTAL_AVAILABLE",
        "BOOK_CONDITION",
        "LONG_TITLE",
        "SHORT_TITLE",
        "AUTHOR",
        "EDITION",
        "IMPRINT",
        "Base Price",
        "WEB_DISP_BEHAVIOR",
        "COSTING_METHOD",
        "PREFERRED_LOCATION",
        "COGS_ACCOUNT",
        "INCOME_ACCOUNT",
        "ASSET_ACCOUNT",
        "TAX_SCHEDULE",
        "Webstore Image Name",
    ]
    return df[final_col_order]


def fill_easy_cells(df):  # function to fill cells with pre-existing data
    df["EXTERNAL_ID"] = df.apply(
        lambda row: f"{row['ISBN']}{'parent'}", axis=1
    )  # fills 'EXTERNAL_ID' column with each row's 'ISBN' and string 'parent'
    print(f"{df['EXTERNAL_ID']}")
    df["ITEM_NAME_NUMBER"] = df.apply(
        lambda row: f"{row['DISPLAYNAME']}-{row['ISBN']}", axis=1
    )  # fills 'ITEM_NAME_NUMBER' column with each row's 'DISPLAYNAME' and 'ISBN' with '-' between each
    print(f"{df['ITEM_NAME_NUMBER']}")
    df["Webstore Image Name"] = df.apply(
        lambda row: f"{row['ISBN']}", axis=1
    )  # fills 'Webstore Image Name' field with each row's 'ISBN'
    print(f"{df['Webstore Image Name']}")


def label_dupes(row):  # function to label duplicates as 'New' or 'Used' children items
    if (
        row["occurrence"] == 1
    ):  # if first duplicate (not original), assign values as a 'New' item
        row["EXTERNAL_ID"] = row["EXTERNAL_ID"].replace("parent", "new")
        row["ITEM_NAME_NUMBER"] = row["ITEM_NAME_NUMBER"] + "-New"
        row["MATRIX_TYPE"] = "Child Matrix Item"
        row["CS_PRODUCT_TYPE"] = "New"
        row["PARENT"] = str(row["ISBN"]) + "parent"
        row["UPC"] = row["ISBN"]
        row["INCLUDE_CHILDREN"] = "F"
        row["RENTAL_AVAILABLE"] = "F"
        row["BOOK_CONDITION"] = "New"
        row["Webstore Image Name"] = ""
    elif row["occurrence"] == 2:  # if second duplicate, assign values as a 'Used' item
        row["EXTERNAL_ID"] = row["EXTERNAL_ID"].replace("parent", "used")
        row["ITEM_NAME_NUMBER"] = row["ITEM_NAME_NUMBER"] + "-Used"
        row["MATRIX_TYPE"] = "Child Matrix Item"
        row["CS_PRODUCT_TYPE"] = "Used"
        row["PARENT"] = str(row["ISBN"]) + "parent"
        row["UPC"] = str(row["ISBN"]).replace("978", "290")
        row["INCLUDE_CHILDREN"] = "F"
        row["PRODUCT_CATEGORY"] = "Course Materials Physical (CMP) : Print Used"
        row["RENTAL_AVAILABLE"] = "F"
        row["BOOK_CONDITION"] = "Used"
        row["TAX_SCHEDULE"] = "Not Taxable"
        try:
            row["Base Price"] = round(row["Base Price"] * 0.75, 2)
        except:
            row["Base Price"] = ""
        row["Webstore Image Name"] = ""
    return row


def make_children(
    df,
):  # function to create child items, making 2 extra copies of each row
    repeat_index = df.index.repeat(3)
    df_duplicated = df.loc[repeat_index].reset_index(
        drop=True
    )  # creates new dataframe for original items, duplicate items, and resets index

    df_duplicated["original_index"] = (
        repeat_index.to_numpy()
    )  # creates array([[External ID value 1, External ID value 2], [Item Name Number Value 1, Item Name Number Value 2]]) of duplicates from each 'original_index' item

    df_duplicated["occurrence"] = df_duplicated.groupby(
        "original_index"
    ).cumcount()  # adds 'occurence' column and counts from 0 to last occurence - 1 per row grouping
    current_cols = df.columns.tolist()
    print(current_cols)

    df_duplicated = df_duplicated.apply(
        label_dupes, axis=1
    )  # use func to apply values in each row in given column based on conditional in label_dupes func
    print(df_duplicated)

    df_duplicated = df_duplicated.drop(
        columns=["occurrence", "original_index"]
    )  # remove 'occurrence' and 'original_index' columns
    return df_duplicated


def strip_accents(text):
    nfd_form = unicodedata.normalize(
        "NFD", text
    )  # Normalize string to decomposed form (NFD) where accents are separate characters
    ascii_form = "".join(
        c for c in nfd_form if unicodedata.category(c) != "Mn"
    )  # Filter out non-spacing characters (Mn category, includes accents)
    return str(ascii_form)


def get_images(isbn):  # function to call ISBNdb and retrieve book images
    base_url = f"https://api2.isbndb.com/book/"
    headers = {"accept": "application/json", "Authorization": api_key}

    os.makedirs(dl_folder_home, exist_ok=True)
    book_cover = f"{isbn}.jpg"
    images_save_path = os.path.join(dl_folder_home, book_cover)

    try:  # attempt to call ISBNdb and receive json response
        url = f"{base_url}{isbn}"
        response = req.get(url, headers=headers, stream=True)
        print(headers)
        response.raise_for_status()  # return status code if call unsuccessful
        book_data = response.json()  # assign response to book_data object
        if (
            "image_original" in book_data["book"] and "msrp" in book_data["book"]
        ):  # check if 'image_original' exists in json response's 'book' object
            new_msrp = book_data["book"]["msrp"]
            image = book_data["book"][
                "image_original"
            ]  # assign 'image_original' to image object
            image_link = req.get(
                image, stream=True
            )  # store image url from image object
            image_link.raise_for_status()  # return status code if call unsuccessful
            with open(
                images_save_path, "wb"
            ) as file:  # loop to save image binary data into chunks at images_save_path
                for chunk in image_link.iter_content(chunk_size=8192):
                    file.write(chunk)
            return new_msrp
        elif "msrp" not in book_data["book"]:
            print(f"No MSRP available for {isbn}. Leaving blank.")
            new_msrp = 0
        elif "image" not in book_data["image_original"]:
            messagebox.showerror(
                "Error", f"No image found for {isbn}."
            )  # display message if no image found
        time.sleep(3)
    except HTTPError as e:
        print(f"HTTP Error occurred: {e}")
        print(f"Status code: {e.response.status_code}")
        if e.response.status_code == 404:  # API call responds with 404 error
            print(
                f"No result for ISBN {isbn}. Try at a later date to check if a record exists."
            )
            messagebox.showerror(
                "Error",
                f"No result for ISBN {isbn}. Try at a later date to check if a record exists.",
            )
            return
        elif (
            e.response.status_code == 429
        ):  # API call responds with 429 error indicating 2000 call per day limit reached
            print(f"Daily requests limit met. Please try tomorrow.")
            messagebox.showerror(
                "Error", "Daily requests limit met. Please try again tomorrow."
            )
            return
    except req.exceptions.RequestException as e:
        messagebox.showerror("Error", f"API call failed for {isbn}: \n{e}")
        return None


def resize_covers(
    input_folder, output_folder, canvas_size
):  # function to resize book images to meet NS standards, filling white background if smaller than required dimensions
    output_dir = Path(desktop_path) / output_folder

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for filename in os.listdir(input_folder):
        if filename.lower().endswith((".jpg", ".jpeg")):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_dir, filename)

            try:
                img = Image.open(input_path).convert(
                    "RGB"
                )  # convert image to 'RGB' format
            except Exception as e:
                print(f"Error processing {filename}: \n{e}")
                return

            canvas_width, canvas_height = (
                canvas_size  # define width and height of canvas from passed-in 'canvas_size' object
            )
            img_width, img_height = img.size  # define width and height of image

            canvas = Image.new(
                "RGB", canvas_size, (255, 255, 255)
            )  # create new canvas in 'RGB' format at dimensions from 'canvas_size' object, and fill with white background

            if img_width > canvas_width or img_height > canvas_height:
                ratio_w = (
                    canvas_width / img_width
                )  # defines difference ratio if image is wider than canvas
                ratio_h = (
                    canvas_height / img_height
                )  # defines difference ratio if image is taller than canvas
                scale_factor = min(
                    ratio_w, ratio_h
                )  # determines the smaller of the two ratios to create scale factor

                new_width = int(
                    img_width * scale_factor
                )  # define new image width using the current width multiplied by 'scale_factor'
                new_height = int(
                    img_height * scale_factor
                )  # define new image height using the current height multiplied by 'scale_factor'
                resized_img = img.resize(
                    (new_width, new_height), Image.Resampling.LANCZOS
                )  # create resized image with new width and height and adjust resolution as needed

                x_offset = (
                    canvas_width - new_width
                ) // 2  # define position for image on x-axis (width difference divided by 2)
                y_offset = (
                    canvas_height - new_height
                ) // 2  # define position for image on y-axis (height difference divided by 2)
                canvas.paste(
                    resized_img, (x_offset, y_offset)
                )  # paste resized image on canvas, using x and y offset values to determine center placement
            else:  # determine center placement for image if smaller than or equal to 'canvas_size' dimensions
                x_offset = (canvas_width - img_width) // 2
                y_offset = (canvas_width - img_height) // 2
                canvas.paste(img, (x_offset, y_offset))

            canvas.save(output_path)
            print(f"Processed {filename} saved to {output_path}")


def loading_popup(
    root, task_functions, title="Working", message="Please wait..."
):  # function to display "Loading" window while 'change_books' function runs
    popup = tk.Toplevel(root)
    popup.title(title)
    tk.Label(popup, text=message).pack(padx=20, pady=10)

    progress_bar = ttk.Progressbar(popup, mode="indeterminate", length=200)
    progress_bar.pack(padx=20, pady=10)
    progress_bar.start(10)

    def task_wrapper():
        change_books(*task_functions)
        root.after(1000, popup.destroy)

    worker_thread = threading.Thread(
        target=task_wrapper, daemon=True
    )  # allows 'Loading' window to run on main thread and functions to run on different thread
    worker_thread.start()


def change_books(
    load_file,
    delete_junk_cols,
    reorder_cols,
    rename_and_add_cols,
    # strip_accents,
    reorder_all_cols,
    fill_easy_cells,
    get_images,
    resize_covers,
    make_children,
):  # function to run other functions
    try:
        df = load_file(root)
        if df is None:
            return

        delete_junk_cols(df)

        new_column_order = reorder_cols(df)
        try:  # attempt to assign new_column_order dataframe to main dataframe
            df = df[new_column_order]
        except KeyError as e:
            messagebox.showerror("Error", f"Error reordering columns: \n{e}")
            return

        df = rename_and_add_cols(df)
        # df = df.apply(strip_accents)

        df = reorder_all_cols(df)

        fill_easy_cells(df)
        print(df)

        df["Base Price"] = df["ISBN"].apply(
            get_images
        )  # get book images and MSRP based on ISBN column

        resize_covers(dl_folder_home, resized_images, (600, 600))

        df = make_children(df)
        # identify parent item rows and delete value in Base Price for each parent item
        parent_rows = df["MATRIX_TYPE"] == "Parent Matrix Item"
        df.loc[parent_rows, "Base Price"] = np.nan

        # define compressed folder name and format, then create archive folder of images
        archive_name = os.path.join(desktop_path, f"{resized_images}-compressed")
        archive_format = "zip"
        shutil.make_archive(archive_name, archive_format)
        messagebox.showinfo(
            "Compressed",
            f"Image folder compressed for upload and saved to {dl_folder_home}.",
        )

        try:  # attempt to delete original downloaded image folder, keeping resized images
            shutil.rmtree(dl_folder_home)
            print(f"\nSuccessfully deleted original image folder {dl_folder_home}.")
        except OSError as e:
            print(f"\nError deleting folder: {dl_folder_home}: \n{e}")

        # ask user to select save location for Excel and CSV output
        save_directory = filedialog.askdirectory(
            parent=root, title="Select Save Location"
        )
        excel_name = f"Adptd_Text_Items-{date_string}.xlsx"
        csv_name = f"Adptd_Text_Items-{date_string}.csv"
        excel_save_path = os.path.join(save_directory, excel_name)
        csv_save_path = os.path.join(save_directory, csv_name)

        with pd.ExcelWriter(excel_save_path, engine="openpyxl") as writer:
            df.to_excel(
                writer, sheet_name=excel_name, index=False
            )  # save as Excel and ensure ISBN and UPC columns are 'number' format
            format_cols = ["ISBN", "UPC"]
            for col_name in format_cols:
                col_idx = df.columns.get_loc(col_name) + 1
                for row in writer.sheets[excel_name].iter_rows(
                    min_col=col_idx, max_col=col_idx
                ):
                    for cell in row:
                        cell.number_format = "0"
        df.to_csv(csv_save_path, index=False)  # save as CSV without index column
        messagebox.showinfo(
            "Success",
            f"Processed Excel file saved to {excel_save_path}. \n\nProcessed CSV file saved to {csv_save_path}.",
        )
        root.destroy()  # close script windows and end
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occured: \n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Loading...")

    loading_popup(
        root,
        [
            load_file,
            delete_junk_cols,
            reorder_cols,
            rename_and_add_cols,
            reorder_all_cols,
            fill_easy_cells,
            get_images,
            resize_covers,
            make_children,
            # strip_accents,
        ],
    )

    root.withdraw()

    root.mainloop()
