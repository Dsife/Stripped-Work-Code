import time
import tkinter as tk
import pyautogui
import os
import datetime
import openpyxl
import win32com.client as win32
import re
from datetime import date, datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException

# Create pop-up for user input and allow user input to be pasted into Chrome
def get_user_input():
    def submit_input(event=None):
        user_input = entry.get()
        year_input = year_entry.get()

        # Make sure that user text is encoded
        user_input = user_input.encode('utf-8').decode('utf-8')
        year_input = year_input.encode('utf-8').decode('utf-8')
        print(f"User input: {user_input}")
        print(f"Year input: {year_input}")
        # Zfill pads zeroes to the left
        user_input = str(user_input).zfill(4)
        # Getting the last two digits of the year
        year_last_two = year_input[-2:]
        
        formatted_string = f"Ex.{user_input}"

        user_input = formatted_string
        print(f"Updated user input: {user_input}")

        open_chrome_and_fill_form(user_input, year_input)

        status_label.config(text="Processing complete. You can close this window.")

    # Create user input window
    root = tk.Tk()
    root.title("Enter Text")
    root.geometry("400x300")
    root.configure(bg='#f0f0f0')
    
    label = tk.Label(root, text="Enter the text to paste:", font=('Verdana', 12, 'bold'), fg='#333333', bg='#f0f0f0')
    label.pack(pady=10)
    
    entry = tk.Entry(root, width=40, font=('Verdana', 12), fg='#333333', bg='#e0e0e0', relief='solid')
    entry.pack(pady=5)
    entry.focus()

    label_year = tk.Label(root, text="Enter the year:", font=('Verdana', 12, 'bold'), fg='#333333', bg='#f0f0f0')
    label_year.pack(pady=10)
    # Getting current year
    current_year = str(datetime.now().year)
    year_entry = tk.Entry(root, width=40, font=('Verdana', 12), fg='#333333', bg='#e0e0e0', relief='sunken')
    year_entry.insert(0, current_year)
    year_entry.pack(pady=5)

    status_label = tk.Label(root, text="", font=('Verdana', 10, 'italic'), fg='#008000', bg='#f0f0f0')
    status_label.pack(pady=10)

    # Be able to click enter to submit
    root.bind("<Return>", submit_input)
    # Keeps window open and listens for events(clicks & button strokes)
    root.mainloop()

def sanitize_string(input_string):
    invalid_characters = r'[/\:*?"<>|]'
    sanitized_string = re.sub(invalid_characters, '-', input_string)
    return sanitized_string

# Main section (Deals with Chrome & Excel)
def open_chrome_and_fill_form(user_input, year_input):
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # Open Chrome
    try:
        chromedriver_path = r""
    except:
        chromedriver_path = r"Ex."
    
    service = Service(chromedriver_path)

    driver = webdriver.Chrome(service=service, options=options)

    # Go to the website
    url = "Ex."   
    driver.get(url)
    
    time.sleep(5)

    try:
        # Find the textbox and paste the input
        text_box = driver.find_element(By.XPATH, 'Ex.')
        text_box.clear()  
        text_box.send_keys(user_input)  
        text_box.send_keys(Keys.RETURN)
        print("Input submitted into the search box")

        time.sleep(3)
        # Click icon inside of table
        pyautogui.click(444, 728)
        print("Clicked at coordinates (444, 728)")

        time.sleep(5)

    except Exception as e:
        print(f"Error occurred: {e}")
    # Secondary portion to get into Ex. (if applicable) and gather some info before returning to the main page
    try:
        Ex. = driver.find_element(By.XPATH, 'Ex.')
        Ex._text = Ex.text if Ex. else ""
        print("Ex.: " + Ex._text)
        Ex. = driver.find_element(By.XPATH, 'Ex.')
        Ex._text = Ex.text if Ex. else ""
        print("Ex.: " + Ex.text)
        Ex. = driver.find_element(By.XPATH, 'Ex.')
        Ex._text = Ex.text if Ex. else ""
        print("Ex.: " + Ex.text)
    except NoSuchElementException as e:
        print(f"Ex. element not found: {e}")
        Ex. = ""
    try:
        on_Ex. = False
        # Find Ex.
        try:
            Ex._button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, 'Ex.'))
            )
            Ex._button.click()
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, 'Ex.'))
            )
            on_request_page = True
            print("Button clicked")
        except Exception as inner_e:
            print(f"Button not found")
        # Gathers Ex.
        try:
            Ex. = None
            Ex. = None

            Ex. = ""
            Ex. = ""
            
            if Ex.:
                # Elements on request page
                Ex. = driver.find_element(By.XPATH, 'Ex.')
                Ex. = driver.find_element(By.XPATH, 'Ex.')
            else:
                # Elements on main page
                Ex. = driver.find_element(By.XPATH, 'Ex.')
                Ex. = None
                
            Ex. = Ex. if Ex. else ""
            print("Ex.: " + Ex.)
            
            if Ex.:
                Ex. = Ex.text if Ex. else ""
                print("Ex.: " + Ex.text)
            if on_request_page:
                driver.back()
                time.sleep(3)
            
        except Exception as e:
            print(f"Error finding Ex.: {e}")

    except Exception as e:
        print(f"Error occurred: {e}")

    # Check if table containing the samples is visible
    def is_table_visible():
        try:
            # Find the table
            table = driver.find_element(By.XPATH, 'Ex.')
            return table.is_displayed()
        except:
            return False

        # Get the total number of Ex.
    def Ex.():
        try:
            Ex. = driver.find_element(By.XPATH, 'Ex.')
            text = Ex.

            match = re.search(r'of (\d+)', text)
            if match:
                Ex. = int(match.group(1))
                print("Ex.:", Ex.)
                return Ex.
            else:
                print("Ex.")
                return None
        except NoSuchElementException:
            print("Element not found.")
            return None
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def click_search_buttons():
        try:
            search_buttons = driver.find_elements(By.XPATH, 'Ex.')

            if not search_buttons:
                print("No search buttons found")
                return []

            click_count = 0
            all_scraped_data = []

            for idx, button in enumerate(search_buttons):
                try:
                    driver.execute_script("arguments[0].scrollIntoView(true);", button)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(button))
                    button.click()
                    click_count += 1
                    print(f"Clicked button {click_count}")

                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, 'Ex.'))
                    )

                    scraped_data = scrape_info_from_page()
                    if scraped_data:
                        all_scraped_data.append(scraped_data)

                    driver.back()

                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, 'Ex.'))
                    )

                except Exception as e:
                    print(f"Error in the loop while processing button {idx + 1}: {e}")

            print(f"Total buttons clicked: {click_count}")
            return all_scraped_data

        except Exception as e:
            print(f"An error occurred: {e}")
            return []

    def scrape_info_from_page():
        try:
            try:
                Ex. = driver.find_element(By.XPATH, 'Ex.')
            except: 
                Ex. = driver.find_element(By.XPATH, 'Ex.')

            Ex. = Ex.
            print(Ex.)

            data = {}

            if "Ex." in Ex.:
                print("Ex.")
                dropdown_1 = driver.find_element(By.ID, 'Ex.')
                selected_option = dropdown_1.find_element(By.XPATH, './/option[@selected]')
                data['Ex.'] = selected_option.text

                data['Ex.'] = driver.find_element(By.XPATH, '//*[@id="Ex."]').get_attribute("value")
                data['Ex.'] = driver.find_element(By.XPATH, '//*[@id="Ex."]').get_attribute("value")
                data['Ex.'] = driver.find_element(By.XPATH, '//*[@id="Ex."]').get_attribute("value")

                dropdown_2 = driver.find_element(By.ID, 'Ex.')
                Ex. = dropdown_2.find_element(By.XPATH, './/option[@selected]')
                data['Ex.'] = Ex..text

                data['Ex.'] = driver.find_element(By.XPATH, '//*[@id="Ex."]').get_attribute("value")

                dropdown_3 = driver.find_element(By.ID, 'Ex.')
                Ex. = dropdown_3.find_element(By.XPATH, './/option[@selected]')
                data['Ex.'] = Ex..text

            else:
                print("Ex.")
                data['Ex.'] = driver.find_element(By.XPATH, 'Ex.').text
                data['Ex.'] = driver.find_element(By.XPATH, 'Ex.').text
                data['Ex.'] = driver.find_element(By.XPATH, 'Ex.').text
                data['Ex.'] = driver.find_element(By.XPATH, 'Ex.').text
                data['Ex.'] = driver.find_element(By.XPATH, 'Ex.').text
                data['Ex.'] = driver.find_element(By.XPATH, 'Ex.').text
                data['Ex.'] = driver.find_element(By.XPATH, 'Ex.').text

            return data

        except NoSuchElementException as e:
            print(f"Error extracting data: {e}")
            return None

        try:
            row_element = driver.find_element(By.XPATH, 'Ex.')
            column_elements = row_element.find_elements(By.XPATH, './*')
            if column_elements:
                print("Ex.")
                Ex. = []
                for column in column_elements:
                    Ex..append(column.text.strip())

                Ex. = ", ".join(Ex.)
                print(Ex.)
            else:
                print("Ex.")
                Ex. = []

        except Exception as e:
            print("Error")

        try:
            total_sample_count += 1

            driver.get(current_url)
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located(By.XPATH, 'Ex.'))

        except Exception as e:
            print(f"Error processing {idx + 1}: {e}")

        # Check if we need to go to the next page
        try:
            next_page_button = driver.find_element(By.XPATH, 'Ex.')
            if next_page_button.is_enabled():
                print(f"Moving to the next page.")
                next_page_button.click()
                time.sleep(3)
            else:
                print("No next page button found.")
        except NoSuchElementException:
            print("No next page button found.")

        print("Saving data to Excel")
        save_data_to_excel(all_data, Ex.)

        return all_data
        
    # Takes all data gathered and exports it to Excel
    def save_data_to_excel(all_data, Ex.):
        try:
           # excel_path = r"Ex."
           # excel_path = r"Ex."
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = True
           # workbook = excel.Workbooks.Open(excel_path)
           # sheet = workbook.Sheets(1)

           # start_row = 2

            # Write the collected data
           # for i, row_data in enumerate(all_data, start=2):
            #    for j, cell_data in enumerate(row_data, start=1):
             #       sheet.Cells(i, j).Value = cell_data

          # print("Data saved to first Excel sheet")

            excel_path_2 = r"Ex."
            workbook_2 = excel.Workbooks.Open(excel_path_2)
            sheet_2 = workbook_2.Sheets("Ex.")
            # Get today's date in MM/DD/YYYY format
            today_date = datetime.combine(date.today(), datetime.min.time()).strftime('%m/%d/%Y')
            
            last_row_2 = 1
            # Loop through rows until the first empty cell in column A
            while sheet_2.Cells(last_row_2, 1).Value != None:
                last_row_2 += 1

            next_row_2 = last_row_2

            print(f"Next available row: {next_row_2}")
            
            # Define a mapping between column indexes in the data to specific column letters in the Excel sheet
            column_mapping = {
                0: 'D', 
                1: 'M', 
                3: 'E', 
                4: 'R', 
                7: 'B', 
                8: 'O', 
                9: 'P', 
                10: 'Q', 
                12: 'F',
                13: 'C',
            }

            Ex._pasted = False

            print(f"All Data: {all_data}")

            # Loop through each row of all_data and paste values into the mapped columns
            for i in range(len(all_data)):
                if len(all_data[i]) < max(column_mapping.keys()) + 1:
                    continue
                for j in range(len(all_data[i])):
                    if j in column_mapping:
                        target_column = column_mapping[j]
                        target_cell = f"{target_column}{next_row_2}"
                        if j == 0 and not Ex.:
                            if all_data[i][12] == "Ex.":
                                Ex. = f"Ex.{all_data[i][j]}"
                            else:
                                Ex. = all_data[i][j]
                                
                            sheet_2.Range(target_cell).Value = all_data[i][j]
                            product_pasted = True
                        elif j != 0:
                            sheet_2.Range(target_cell).Value = all_data[i][j]
                            
            sheet_2.Cells(next_row_2, 1).Value = user_input
            sheet_2.Cells(next_row_2, 7).Value = today_date

            print("Successfully pasted into second Excel sheet")

            excel_path_3 = r"Ex."
            workbook_3 = excel.Workbooks.Open(excel_path_3)
            print("Workbook 3 opened successfully.")
            sheet_3 = workbook_3.Sheets("2025")
            print("Sheet 'Ex.' loaded successfully.")

            print(f"Ex.: {Ex.}")

            #Find the first empty cell in column A starting from next_row_3
            last_row_3 = 1

            #Loop through each row to find the first empty cell in column A
            while sheet_3.Cells(last_row_3, 1).Value is not None:
                print(f"Row {last_row_3} is not empty: {sheet_3.Cells(last_row_3, 1).Value}")
                last_row_3 += 1

            #Now, last_row_3 should hold the first empty row
            next_row_3 = last_row_3

            # Debugging: Print the result
            print(f"Next available row in sheet_3: {next_row_3}")
            
            # Define a function to convert a number to a base-26 suffix (A, B, ..., Z, AA, AB, etc.)
            def number_to_suffix(index):
                #Convert a number to a base-26 string (e.g., 0 -> 'A', 25 -> 'Z', 26 -> 'AA', etc.).
                suffix = ""
                while index >= 0:
                    suffix = chr(index % 26 + 65) + suffix # Get the character for the current "digit" (A-Z)
                    index = index // 26 - 1 # Move to the next digit
                return suffix

            for i in range(total_samples):
                Ex. = all_data[i][0]
                Ex. = all_data[i][8]
                Ex. = all_data[i][2]
                Ex. = all_data[i][3]
                Ex. = all_data[i][12]

                Ex. = sanitize_string(Ex.)
                Ex. = sanitize_string(Ex.)

                if total_samples > 1:
                    sample_suffix = number_to_suffix(i)  # Generate the suffix based on the index
                    unique_user_input = f"{user_input}{sample_suffix}" # Combine them
                else:
                    unique_user_input = user_input
                    
                # Paste the into Cell A & B
                sheet_3.Cells(next_row_3 + i, 1).Value = unique_user_input
                print(f"Ex.")
                if Ex. == "Ex.":
                    Ex. = f"Ex. {Ex.()} {Ex.()} {Ex.()}"
                else:
                    Ex. = f"{Ex.()} {Ex.()} {Ex.()}"

                sheet_3.Cells(next_row_3 + i, 2).Value = clean_product_info

                cell = sheet_3.Cells(next_row_3 + i, 1)
                # Add comment
                cell.AddComment(f"Aim: {Ex.}")
                comment = cell.Comment
                if comment:
                    comment.Text("")
                    comment.Text(f"Aim: {Ex.}")
                print("Added comment")
                
            print("Finished processing and writing data to third workbook.")

        except Exception as e:
            print(f"Error occurred while saving data to Excel: {e}")

        time.sleep(3)

    if is_table_visible():
        print("Table is visible")
        get_total_samples()
        scrape_info_from_page()
    else:
        print("Table is not visible.")

get_user_input()
