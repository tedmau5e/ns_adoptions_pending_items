# ns_adoptions_pending_items
Python script to create import file from NS of textbook adoptions pending items

1/31 Edit: Implemented MSRP retrieval as part of API call and logic to determine used book price (75% of new price)

12/3 Edit: Now supports opening .xls files from NS that are actually .xml files, converting to .xlsx dataframe and continuing script. ISBNdb API connection to retrieve images for books. Implemented function to resize images for NS standard, and then automatically compress folder for upload. Now exports .xlsx and .csv file once complete. Enhanced messaging for users, including more user friendly options.
