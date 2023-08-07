import importlib
import subprocess
import os
import re
from tkinter import messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
import xlwings as xw
import pandas as pd

def check_and_install_library(library_name):
    """
    Check if a library is installed and install it if not.

    Args:
        library_name (str): The name of the library to check and install.
    """
    try:
        importlib.import_module(library_name)
    except ImportError:
        print(f"Installing {library_name}...")
        subprocess.check_call(["pip", "install", library_name])

def scrape_dynamic_website(url):
    """
    Scrape dynamic content from a website and save it to a text file.

    Args:
        url (str): The URL of the website to scrape.
    """
    # Create a Chrome WebDriver
    try:
        driver = webdriver.Chrome()

        # Navigate to the URL
        driver.get(url)

        # Wait for the dynamic content to load (you may need to adjust the time)
        driver.implicitly_wait(10)

        # Find and extract the dynamic elements
        table_element = driver.find_element(By.CLASS_NAME, "timeline-lines")
        table_content = table_element.get_attribute('outerHTML')

        # Get the root folder of the script
        save_path = os.path.dirname(os.path.abspath(__file__))

        # Create the "bin" folder if it doesn't exist
        if not os.path.exists(save_path):
            os.makedirs(save_path)

        # Save the dynamic content to a .txt file in the "bin" folder
        file_path = os.path.join(save_path, 'raw_data.txt')
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(table_content)

        print("Dynamic content has been copied to: ", file_path)

    except Exception as e:
        print(f"An error occurred while scraping the website: {e}")   

    finally:
        # Close the WebDriver
        driver.quit()

def read_file(file_path):
    """
    Read the content of a file.

    Args:
        file_path (str): The path to the file to read.

    Returns:
        str: The content of the file.
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        return content
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return ""

def extract(file_content, search_string, char_length):
    """
    Extract occurrences of a search string from the file content.

    Args:
        file_content (str): The content of the file to search in.
        search_string (str): The search string to find occurrences of.
        char_length (int): The length of characters to extract after the search string.

    Returns:
        list: A list of extracted occurrences.
    """
    timestamps = []
    start_index = 0
    
    while True:
        # Find the occurrence of the search string in the content
        start_index = file_content.find(search_string, start_index)
        
        if start_index == -1:
            # If no more occurrences are found, break the loop
            break
        
        # Extract the timestamp (next char_length characters after the search string)
        timestamp = file_content[start_index + len(search_string): start_index + len(search_string) + char_length]
        
        # Remove commas from the timestamp and append to the list
        #timestamps.append(timestamp.replace(',!@#$%^&', ''))
        timestamps.append(re.sub(r'[,!@#$%^&]', '', timestamp))
        
        # Move the start_index to continue searching from the next position
        start_index += len(search_string)
    
    return timestamps

def main():
    # List of required libraries
    required_libraries = ["lxml", "requests", "bs4", "pandas", "openpyxl", "selenium"]

    # Check and install required libraries
    for library in required_libraries:
        check_and_install_library(library)

    # Connect to the active Excel application
    #pp = xw.App(visible=False)
    #wb = xw.books.active
    # Get the URL from cell Q5 of the "Intro" sheet
    #url = wb.sheets['Intro'].range('Q5').value
    url = 'https://www.warcraftlogs.com/reports/JNpFkaVh3Wbv7HZn/#fight=last&type=casts&source=76&start=18429&end=332138&view=timeline'
    scrape_dynamic_website(url)

    # Get the current working directory
    root_folder = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(root_folder, "raw_data.txt")
    file_content = read_file(file_path)
    search_timestamp = '<div onmouseover="showTimelineTooltipText(this, 1, printEvent({&quot;timestamp&quot;:'
    search_src = 'abilities/'
    timestamps = extract(file_content, search_timestamp, 7)
    src = extract(file_content, search_src, 25)

    #print(timestamps)
    #timestamps
    #print(src)
    data = {'timestamps': timestamps, 'Name': src}
    df = pd.DataFrame(data)
    #df.head(10)
    print("Got timestamps and spell names")

    #sort the data
    df['timestamps'] = pd.to_numeric(df['timestamps'])

    # Sort the DataFrame by the "timestamps" column in ascending order
    df_sorted = df.sort_values(by='timestamps', ascending=True)

    #df_sorted.head(10)
    df_spell_density=df.groupby("Name")
    df_spell_density.size()
    df_spell_counts = df_spell_density.size().reset_index(name='Count')
    df_spell_counts_sorted = df_spell_counts.sort_values(by='Count', ascending=False)
    df_spell_counts_sorted.head(10)
    print("Data sorted")



    # Specify the new Excel file name
    excel_file_new = os.path.join(root_folder, "Rotation_WarcraftLogs.xlsx")

    # Save the df_sorted DataFrame to the new Excel file with the specified sheet name and starting from cell A1
    with pd.ExcelWriter(excel_file_new, engine='openpyxl') as writer:
        df_sorted.to_excel(writer, sheet_name="Warcraft_Logs", index=False)

        # Specify the starting cell for df_spell_counts_sorted
        start_row = 0
        start_col = 5  # This will be the column after the last column of df_sorted

        # Save the df_spell_counts_sorted DataFrame to the same Excel file and sheet
        df_spell_counts_sorted.to_excel(writer, sheet_name="Warcraft_Logs", startrow=start_row, startcol=start_col, index=False)

    print("Created/updated excel in:", root_folder)
    #-------------------------------------------------------------POPO A MESSAGEBOX WHEN DONE---------------------------------------------------------------

    # Show a message box to inform the user
    #messagebox.showinfo("Python Script Executed", f"Dynamic content has been copied to: {file_path}\nData extracted and sorted in: {root_folder}")

if __name__ == "__main__":
    main()