# Chrome Automation with Excel Integration

This Python script is a complete automation solution for interacting with a specific web application and exporting the collected data into Excel spreadsheets. It leverages a combination of Tkinter, Selenium, pyautogui, and win32com to streamline a repetitive process that involves web navigation, data scraping, and Excel entry.

User Input Collection
  Launches a Tkinter GUI where the user enters:
    A numeric ID (e.g., 1234)
    A year (e.g., 2025)
  Formats the ID into a specific structure like Ex.1234.

Automated Web Interaction
  Opens Google Chrome with Selenium and navigates to a target website.
  Inputs the formatted ID into a search box and submits the query.
  Clicks on specific UI elements using both Selenium and pyautogui.
  
Data Extraction
  Attempts to gather relevant textual data from various elements on the web page.
  Navigates into deeper pages (e.g., detail or request pages) if necessary to gather additional information.
  Collects dropdown values, text fields, and any structured data within rows/tables.
  
Handling Multiple Entries
  Iteratively clicks through a list of items (e.g., search results).
  For each item:
    Extracts structured information.
    Navigates back to the main list to process the next item.
    Supports multi-page navigation if multiple search result pages exist.
    
Excel Automation
  Opens one or more Excel workbooks via win32com.client.
  Pastes the scraped data into specific columns based on a defined mapping.
  Dynamically identifies the next available row to avoid overwriting existing entries.
  Adds metadata such as:
    Today’s date
    User-input ID
    Comments generated from the data
    
Data Sanitization
  Ensures extracted strings are sanitized for invalid characters before saving.
  Converts multi-sample entries to unique identifiers (e.g., Ex.1234A, Ex.1234B, etc.)
