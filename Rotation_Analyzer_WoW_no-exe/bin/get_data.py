import importlib
import subprocess
import os
import re
import tkinter as tk
from tkinter import simpledialog
import tkinter.messagebox as mbox
import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook

def check_and_install_library(library_name):
    try:
        importlib.import_module(library_name)
    except ImportError:
        print(f"Installing {library_name}...")
        subprocess.check_call(["pip", "install", library_name])

def scrape_dynamic_website(url, save_path):
    try:
        driver = webdriver.Chrome()
        driver.get(url)
        driver.implicitly_wait(10)
        table_element = driver.find_element(By.CLASS_NAME, "timeline-lines")
        table_content = table_element.get_attribute('outerHTML')

        if not os.path.exists(save_path):
            os.makedirs(save_path)

        file_path = os.path.join(save_path, 'raw_data.txt')
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(table_content)

        print("Dynamic content has been copied to:", file_path)

    except Exception as e:
        print(f"An error occurred while scraping the website: {e}")
    finally:
        driver.quit()

def read_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
        return content
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return ""

def extract(file_content, search_string, char_length):
    timestamps = []
    start_index = 0

    while True:
        start_index = file_content.find(search_string, start_index)

        if start_index == -1:
            break

        timestamp = file_content[start_index + len(search_string): start_index + len(search_string) + char_length]
        timestamps.append(re.sub(r'[,!@#$%^&]', '', timestamp))
        start_index += len(search_string)

    return timestamps

def get_url_from_user():
    root = tk.Tk()
    root.withdraw()

    # Prompt the user for the URL using a simple dialog box
    while True:
        # Prompt the user for the URL using a simple dialog box
        url = simpledialog.askstring("Input", "Enter the URL:")
        if url:
            return url  # Return the URL if it is provided
# ...

def main():
    required_libraries = ["lxml", "requests", "bs4", "pandas", "openpyxl", "selenium"]
    for library in required_libraries:
        check_and_install_library(library)

    # Get the directory where the .exe file is located
    exe_directory = os.path.dirname(sys.argv[0])

    script_path_filename = os.path.abspath(__file__)
    script_path = os.path.dirname(script_path_filename)
    excel_path = os.path.dirname(script_path)
    file_name = "Rotation_Analyzer.xlsm"

    temp_dir = sys._MEIPASS
    full_excel_path = os.path.join(temp_dir, file_name)

    intro_sheet = "INTRO"
    url_cell = "Q5"

    wb = load_workbook(filename=full_excel_path, read_only=True)
    sheet = wb[intro_sheet]
    #url = sheet[url_cell].value
    #url = input("Enter the URL: ")
    url = get_url_from_user()
    
    # Change the save_dir to be within the exe_directory
    save_dir = os.path.join(exe_directory, "WarcraftLogs_Data")
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    scrape_dynamic_website(url, save_dir)

    file_path = os.path.join(save_dir, "raw_data.txt")
    file_content = read_file(file_path)

    search_timestamp = '<div onmouseover="showTimelineTooltipText(this, 1, printEvent({&quot;timestamp&quot;:'
    search_src = 'abilities/'
    timestamps = extract(file_content, search_timestamp, 7)
    src = extract(file_content, search_src, 25)

    data = {'timestamps': timestamps, 'Name': src}
    df = pd.DataFrame(data)

    df['timestamps'] = pd.to_numeric(df['timestamps'])

    df_sorted = df.sort_values(by='timestamps', ascending=True)

    df_spell_density = df.groupby("Name")
    df_spell_counts = df_spell_density.size().reset_index(name='Count')
    df_spell_counts_sorted = df_spell_counts.sort_values(by='Count', ascending=False)

    # Save the Rotation_WarcraftLogs.xlsx file inside the WarcraftLogs_Data folder
    excel_file_new = os.path.join(save_dir, "Rotation_WarcraftLogs.xlsx")

    with pd.ExcelWriter(excel_file_new, engine='openpyxl') as writer:
        df_sorted.to_excel(writer, sheet_name="Warcraft_Logs", index=False)
        start_row = 0
        start_col = 5
        df_spell_counts_sorted.to_excel(writer, sheet_name="Warcraft_Logs", startrow=start_row, startcol=start_col,
                                        index=False)

    # Show the message box at the end of the main function
    mbox.showinfo("Finished", "You can proceed")
if __name__ == "__main__":
    main()

