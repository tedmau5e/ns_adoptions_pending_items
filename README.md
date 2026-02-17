# ns_adoptions_pending_items
Python script to create import file from NS of textbook adoptions pending items

API Key will need to be requested from Ted in order to successfully run this script.

1. Download and save all files in one folder on your Desktop.
2. Download and install Python: https://www.python.org/downloads/
3. Open Command Prompt (on Windows) or Terminal (on Mac) and run 'pip install pyinstaller' (Windows) or 'pip3 install pyinstaller' (Mac)
4. Verify installation with 'pyinstaller --version'
5. In Command Prompt or Terminal, navigate to the downloaded folder (on Windows, you can right click on the folder in File Explorer and choose 'Open in Terminal'; on Mac, open containing folder in Finder, right click and choose Services, New Terminal at Folder)
6. Run 'pip freeze > requirements.txt' (or 'pip3 freeze > requirements.txt' on Mac) to generate external package requirements.
7. Run 'pip install -r requirements.txt' (or 'pip3 freeze > requirements.txt' on Mac) to install required external packages.
8. After getting API Key from Ted, run 'pyinstaller --add-data "ISBNdb_API_Key.env" --onefile adoptions_pending_items.py' to create packaged executable file.
9. In Finder/File Explorer, open new folder named 'adoptions_pending_items', open 'dist' folder and locate App (for MacOS) or executable (.exe file, for Windows).
9a. For MacOS, right click on App and choose "Copy" to place App on Desktop or Applications, etc.
9b. For Windows, right click on executable and select "Show more options", then "Send to > Desktop (create shortcut).
10. Double click on App or executable to run, selecting prompts for file to open and location to save re-formatted Excel and CSV file. (BE PATIENT. The script can take up to a minute before loading file selection window.) Image folder and compressed images will be saved to same location as application folder.

1/31 Edit: Implemented MSRP retrieval as part of API call and logic to determine used book price (75% of new price)

12/3 Edit: Now supports opening .xls files from NS that are actually .xml files, converting to .xlsx dataframe and continuing script. ISBNdb API connection to retrieve images for books. Implemented function to resize images for NS standard, and then automatically compress folder for upload. Now exports .xlsx and .csv file once complete. Enhanced messaging for users, including more user friendly options.
