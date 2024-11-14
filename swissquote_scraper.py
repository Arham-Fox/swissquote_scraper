#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import time
import pandas as pd
from bs4 import BeautifulSoup
import undetected_chromedriver as uc
from datetime import datetime
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

start_time = time.time()

# Set the path to your ChromeDriver executable
chromedriver_path = '/usr/bin/chromedriver'

# Configure ChromeOptions to simulate a regular browser
chrome_options = Options()
chrome_options.binary_location='/usr/bin/chromium'
# Disables Web Notifications and Push APIs
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--disable-popup-blocking")
# Skip First Run wizards
chrome_options.add_argument("--no-first-run")
# Stores password in plain text
chrome_options.add_argument("--password-store=basic")
chrome_options.add_argument("--enable-automation")
# Disable sandboxing features. Should avoid problems on Linux
chrome_options.add_argument("--no-sandbox")
# Disable the setuid sandbox
chrome_options.add_argument("--disable-setuid-sandbox")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--sisable-accelerated-2d-canvas')
chrome_options.add_argument('--disable-gpu-compositing')
# Disable automation detection by the website
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
# Disable animations for shorter load times
chrome_options.add_argument("-disable-animations")
# Disabe Chrome Sync
chrome_options.add_argument("--disable-sync")
# Disable shared memory
chrome_options.add_argument("--disable-dev-shm-usage")
# Could reduce the change of detection
chrome_options.add_argument("--headless")

# Create a ChromeDriver instance
driver = uc.Chrome(version_main=130, service=Service(chromedriver_path), options=chrome_options)
# driver = uc.Chrome(service=Service(chromedriver_path), options=chrome_options)
driver.maximize_window()


def click_button_by_class_name(button_class, timeout=10):
        btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, f".{button_class.replace(' ', '.')}"))
        )
        btn.click()
        time.sleep(0.25)
        
def click_button_by_custom_attribute(attribute, value, timeout=10):
        try:
            # Wait for the element to be visible, then clickable
            btn = WebDriverWait(driver, timeout).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, f"[{attribute}='{value}']"))
            )
            WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, f"[{attribute}='{value}']"))
            )
            btn.click()
            time.sleep(0.25)
    
            
        except Exception as e:
            print(f"Failed to click element with [{attribute}='{value}']: {e}")


def click_button_by_xpath(xpath, timeout=10):
        try:
            btn = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            btn.click()
            time.sleep(0.25)
            
        except Exception as e:
            print(f"Failed to click element with XPath '{xpath}': {e}")  
        

def select_radio_button_by_xpath(xpath, timeout=10):
        try:
            radio_btn = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            
            # Click the radio button if it's not already selected
            if not radio_btn.is_selected():
                radio_btn.click()
                time.sleep(0.25)
                
        except Exception as e:
            print(f"Failed to select radio button with XPath '{xpath}': {e}")
        


def save_dataframe_to_xlsx(df, filename):
        """
        Saves a Pandas DataFrame to an XLSX file with autofit columns.
        
        Parameters:
        - df: The DataFrame to save.
        - filename: The name of the file to save (without extension).
        """
        if not isinstance(df, pd.DataFrame):
            raise ValueError("Input must be a Pandas DataFrame")
        
        # Define the file path
        file_path = f"{filename}.xlsx"
        
        # Save the DataFrame to an XLSX file using XlsxWriter as the engine
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=filename, index=False, freeze_panes=(1, 0))
            
            # Access the workbook and worksheet objects
            workbook  = writer.book
            worksheet = writer.sheets[filename]
            
            # Autofit the columns based on the max length in each column
            for column_num, column in enumerate(df.columns):
                max_length = max(
                    df[column].astype(str).map(len).max(),  # Length of largest item
                    len(str(column))                        # Length of column header
                ) + 2  # Add a little extra padding
                
                # Set the column width
                worksheet.set_column(column_num, column_num, max_length)
        
        print(f"DataFrame saved as '{file_path}' successfully.")    
    

def read_table_with_header_to_dataframe(timeout=10):
    """
    Reads a table into a DataFrame, optimizing link extraction and using BeautifulSoup for parsing.
    """
    dev = False
    try:
        if dev:
            print("[i] Looking for Table...")
        
        # Locate the table and extract its HTML content in one go
        table = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CLASS_NAME, "s-table.SecuritiesSearchPlugin-Table"))
        )
        
        # Parse the table's HTML with BeautifulSoup
        soup = BeautifulSoup(table.get_attribute('outerHTML'), 'html.parser')
        
        # Extract headers
        headers = [th.get_text(strip=True) for th in soup.select("thead th")]
        
        # Initialize table data and links storage
        table_data = []
        links = {}
        
        # Extract rows and process each row
        rows = soup.select("tbody tr")
        for row_idx, row in enumerate(rows):
            cells = row.find_all("td")
            row_data = [cell.get_text(strip=True) for cell in cells]
            
            # Extract links if any and store them by row index
            link_element = row.find("a", href=True)
            links[row_idx] = link_element['href'] if link_element else None
            
            # Append the row data to table_data
            table_data.append(row_data)
        
        # Append the 'Link' column to headers
        headers.append("Link")
        
        # Add link data to table_data
        for row_idx, row_data in enumerate(table_data):
            row_data.append(links[row_idx])  # Append link to the end of the row data
        
        # Create the DataFrame
        df = pd.DataFrame(table_data, columns=headers)
        if dev:
            print("[i] DataFrame created.")
        
        return df

    except Exception as e:
        print(f"Failed to read table with header into DataFrame: {e}")
        return None


def read_fundamentals_to_series(timeout=10):
        """
        Reads fundamental data from a webpage and returns it as a Pandas Series.
    
        The data is structured with headers and values organized in rows within an article element.
    
        Parameters:
        - timeout: Time to wait for elements to become available.
    
        Returns:
        - A Pandas Series containing the extracted data.
        """
        try:
            # Locate the article containing the fundamentals data
            article = WebDriverWait(driver, timeout).until(
                # EC.presence_of_element_located((By.CLASS_NAME, "FullQuote_Card"))
                EC.visibility_of_element_located((By.CSS_SELECTOR, "[data-cardid='FundamentalsCard']"))
            )
    
            # Locate all the data rows within the article
            rows = article.find_elements(By.CLASS_NAME, "FullQuote_DataRow")
    
            # Initialize a dictionary to store the header-value pairs
            data = {}
    
            for row in rows:
                # Extract header and value
                header_span = row.find_elements(By.TAG_NAME, "span")
                if len(header_span) >= 2:
                    header_text = header_span[0].text.strip()  # First span is the header
                    value_text = header_span[1].text.strip()  # Second span is the value
                    
                    # Store in the dictionary
                    data[header_text] = value_text
    
            # Create a Pandas Series from the dictionary
            series = pd.Series(data)
    
            return series
    
        except Exception as e:
            print(f"Failed to read fundamentals data into Series: {e}")
            return None


def press_page_down_n_times(n, wait_time=1):
        """
        Press the Page Down key n times with a wait time between each press.
    
        Args:
        - n (int): The number of times to press the Page Down key.
        - wait_time (float): The amount of time to wait between each key press (in seconds).
        """
        body = driver.find_element(By.TAG_NAME, 'body')
        body.click()  # Click the body to set focus
    
        for _ in range(n):
            body.send_keys(Keys.PAGE_DOWN)  # Press the Page Down key
            time.sleep(wait_time)  # Wait for the specified time

# =============================================================================
# Settings
# =============================================================================
initial_url = 'https://www.swissquote.ch/trading-platform/#scanner'

# =============================================================================
# START
# =============================================================================
if __name__ == '__main__':
        print("Navigate to the specified URL")
        driver.get(initial_url)
        time.sleep(3)
        
        print("Set language to German")
        click_button_by_class_name("Button Button--ghost Languages__triggerButton globalNavigation-Button globalNavigation-Button--ghost globalNavigation-Languages__triggerButton", 10)
        click_button_by_xpath("/html/body/div[6]/div/div/ul/li[3]", 10)
        
        print("Add a new scanner")
        click_button_by_class_name("Button Button--outlined Button--small SecuritiesSearchPlugin-ScannerContainer__addScannerButton", 10)

        print('Select "Höchste Kapitalisierung der Schweiz"')
        select_radio_button_by_xpath("/html/body/div[6]/div/div/div/div[2]/div[2]/div[2]/div/div/div[2]/label", 10)

        print('Click "Hinzufügen"')
        click_button_by_xpath("/html/body/div[6]/div/div/div/div[3]/div/div[2]/button", 10)
        
        print("Scroll down to load full table")
        press_page_down_n_times(17, wait_time=2)
        # idea. Number of shares is known. grab it. scroll down until table has n elements.

        
        print("Read table of securities")
        df = read_table_with_header_to_dataframe(10)

            
        if df is not None and "Link" in df.columns:
            print("Links found in the DataFrame:")
            
            # List to hold DataFrames for each link
            fundamentals_dataframes = []
            
            # Iterate through each link in the DataFrame
            for count, link in enumerate(df['Link'], start=1):
                if link:
                    print(f"Fetching fundamentals for link: {link}")
                    # Navigate to the link using the driver
                    driver.get(link)  # Navigate to the link
                    time.sleep(1)
                    
                    # Fetch the fundamental data
                    fundamentals_series = read_fundamentals_to_series(10)
                    
                    # Combine the DataFrame and Series if fundamentals_series is not None
                    if fundamentals_series is not None:
                        # Convert Series to DataFrame and transpose it
                        fundamentals_df = fundamentals_series.to_frame(name='Value').reset_index()
                        fundamentals_df.columns = ['Header', 'Value']  # Rename columns for clarity
                        
                        # Transpose the fundamentals_df so that each header becomes a column
                        transposed_df = fundamentals_df.set_index('Header').T
                        transposed_df['Link'] = link
                        
                        # Append the transposed DataFrame to the list
                        fundamentals_dataframes.append(transposed_df)
                        
                        # if count == 3:
                        #     break                        
                    else:
                        print(f"Failed to fetch fundamentals data for the link: {link}")
        
            # After collecting data for all links, combine all DataFrames
            if fundamentals_dataframes:
                # Concatenate all transposed DataFrames into a single DataFrame
                combined_fundamentals_df = pd.concat(fundamentals_dataframes, ignore_index=True)
        
                # Merge combined_fundamentals_df with the original df using the Link column
                final_combined_df = pd.merge(df, combined_fundamentals_df, on='Link', how='left')
        
                # Remove the '%' character from column Dividendenrendite for sorting
                final_combined_df['Dividendenrendite'] = final_combined_df['Dividendenrendite'].str.replace('%', '', regex=False)
                
                # Convert to numeric for sorting
                final_combined_df['Dividendenrendite'] = pd.to_numeric(final_combined_df['Dividendenrendite'])
                    
                # Sort by column Dividendenrendite, descending
                sorted_df = final_combined_df.sort_values(by=['Dividendenrendite'], ascending=False)
                
                # Add the '%' character back after sorting
                sorted_df['Dividendenrendite'] = sorted_df['Dividendenrendite'].astype(str) + '%'
                
                del sorted_df[''] # first col, no header, content is always "Trade"
                del sorted_df['Marktkapitalisierung']
                del sorted_df['Bereich']
                del sorted_df['Branche_y']
                del sorted_df['Land']
                del sorted_df['Börse'] # always "SIX"
                del sorted_df['Währung'] # always "CHF"
                
                # Move Link column to the end
                end_pos = len(sorted_df.columns)-1
                sorted_df.insert(end_pos, 'Link', sorted_df.pop('Link'))
                # column_to_move = sorted_df.pop("Link")     
                # sorted_df.insert(len(sorted_df.columns), "Link", column_to_move)
           
                filename = f'{datetime.now().strftime("%Y%m%d-%H%M%S")}-sq-ch-aktien'
                save_dataframe_to_xlsx(sorted_df, filename)
            else:
                print("No fundamental data collected from any links.")
        else:
            print("No links found or DataFrame is empty.")         
        
        
        # Close the browser
        driver.quit()

# Calculate and format execution time.
end_time = time.time()
execution_time = end_time - start_time
hours = int(execution_time // 3600) # Floor division
minutes = int((execution_time % 3600) // 60)
seconds = int((execution_time % 3600) % 60)

print("-"*80)
print(f"----- Duration: {hours} Hrs, {minutes} Mins, {seconds} Secs.")
print("-"*80)


print(sorted_df['Dividendenrendite'])



















