# import pandas as pd
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.common.exceptions import TimeoutException
# import re
# import logging

# # Configure logging
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# # Load the Excel file
# excel_file = "C:/Users/vigne/OneDrive/Documents/Projects/auto/Book1.xlsx"  # Replace with your file path
# data = pd.read_excel(excel_file)

# # Define the column names
# link_column = "short link"  # The column containing short links
# product_name_column = "product name"
# sku_column = "sku"
# status_column = "status"
# old_link_column = "old link"

# # Initialize Selenium WebDriver (using Chrome in this example)
# service = Service("C:/Windows/chromedriver.exe")  # Replace with the path to your ChromeDriver
# options = webdriver.ChromeOptions()
# options.add_argument("--headless")  # Run browser in headless mode for automation
# options.add_argument("--disable-gpu")
# options.add_argument("--no-sandbox")
# driver = webdriver.Chrome(service=service, options=options)

# # Function to find the correct product and update the Excel file
# def process_listing_page(url, expected_product_name, expected_sku):
#     logging.info(f"Processing listing page: {url}")
#     driver.get(url)
#     try:
#         # Wait until the page fully loads
#         WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

#         # Locate product listings on the page (customize selectors based on page structure)
#         products = driver.find_elements(By.CSS_SELECTOR, "div.product-thumb")
#         for product in products:
#             try:
#                 # Extract product name from the <a> tag inside the "name" class
#                 product_name_element = product.find_element(By.CSS_SELECTOR, "div.name a")
#                 product_name = product_name_element.text.lower()

#                 # Extract SKU using the updated selector
#                 product_sku_element = product.find_element(By.CSS_SELECTOR, "#content > div.main-products-wrapper > div.main-products.product-grid > div:nth-child(1) > div > div.caption > div.stats > span.stat-2 > span:nth-child(2)")
#                 product_sku = product_sku_element.text.strip().lower().split('#')[0]

#                 logging.info(f"Found product: Name='{product_name}', SKU='{product_sku}'")

#                 # Normalize and check product name similarity
#                 expected_tokens = re.findall(r'\b\w+\b', expected_product_name.lower())
#                 product_tokens = re.findall(r'\b\w+\b', product_name)
#                 match_count = sum(1 for token in expected_tokens if token in product_tokens)
#                 match_ratio = match_count / len(expected_tokens)

#                 logging.info(f"Match ratio: {match_ratio}, Expected SKU='{expected_sku.lower()}', Found SKU='{product_sku}'")

#                 if match_ratio > 0.8 and product_sku == expected_sku.lower():
#                     # Extract the product page link from the <a href> attribute
#                     product_page_url = product_name_element.get_attribute("href")
#                     logging.info(f"Match found: Product Page URL='{product_page_url}'")
#                     return product_page_url

#             except Exception as e:
#                 logging.error(f"Error processing product: {e}")

#         logging.info("No matching product found on the listing page.")
#         return None  # Return None if no matching product is found
#     except TimeoutException:
#         logging.error("Timeout while loading the page.")
#         return None

# # Iterate through each row in the Excel file
# for index, row in data.iterrows():
#     short_link = row[link_column]
#     product_name = row[product_name_column]
#     sku = row[sku_column]

#     logging.info(f"Processing row {index + 1}: Product Name='{product_name}', SKU='{sku}', Short Link='{short_link}'")

#     # Process the listing page
#     product_page_url = process_listing_page(short_link, product_name, sku)

#     if product_page_url:
#         # Update the old link column and status, overwriting any existing text in the old link column
#         data.at[index, old_link_column] = product_page_url
#         data.at[index, status_column] = "Updated"
#     else:
#         data.at[index, status_column] = "Not Found"

# # Save the updated Excel file
# updated_excel_file = "updated_file.xlsx"
# data.to_excel(updated_excel_file, index=False)
# logging.info(f"Updated Excel file saved to: {updated_excel_file}")

# # Quit the driver
# driver.quit()


import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import re
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Load the Excel file
excel_file = "C:/Users/vigne/OneDrive/Documents/Projects/auto/Book1.xlsx"  # Replace with your file path
data = pd.read_excel(excel_file)

# Define the column names
link_column = "short link"  # The column containing short links
product_name_column = "product name"
sku_column = "sku"
status_column = "status"
old_link_column = "old link"

# Initialize Selenium WebDriver (using Chrome in this example)
service = Service("C:/Windows/chromedriver.exe")  # Replace with the path to your ChromeDriver
options = webdriver.ChromeOptions()
#options.add_argument("--headless")  # Run browser in headless mode for automation
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
driver = webdriver.Chrome(service=service, options=options)

# Function to find the correct product and update the Excel file
def process_listing_page(url, expected_product_name, expected_sku):
    logging.info(f"Processing listing page: {url}")
    driver.get(url)
    try:
        # Wait until the page fully loads
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # Locate product listings on the page (customize selectors based on page structure)
        products = driver.find_elements(By.CSS_SELECTOR, "div.product-thumb")
        for product in products:
            try:
                # Extract product name from the <a> tag inside the "name" class
                product_name_element = product.find_element(By.CSS_SELECTOR, "div.name a")
                product_name = product_name_element.text.lower()

                # Extract SKU using the updated selector
                product_sku_element = product.find_element(By.CSS_SELECTOR, "#content > div.main-products-wrapper > div.main-products.product-grid > div:nth-child(1) > div > div.caption > div.stats > span.stat-2 > span:nth-child(2)")
                product_sku = product_sku_element.text.strip().lower().split('#')[0]

                logging.info(f"Found product: Name='{product_name}', SKU='{product_sku}'")

                # Normalize and check product name similarity
                expected_tokens = re.findall(r'\b\w+\b', expected_product_name.lower())
                product_tokens = re.findall(r'\b\w+\b', product_name)
                match_count = sum(1 for token in expected_tokens if token in product_tokens)
                match_ratio = match_count / len(expected_tokens)

                logging.info(f"Match ratio: {match_ratio}, Expected SKU='{expected_sku.lower()}', Found SKU='{product_sku}'")

                if match_ratio > 0.8 and product_sku == expected_sku.lower():
                    # Extract the product page link from the <a href> attribute
                    product_page_url = product_name_element.get_attribute("href")
                    logging.info(f"Match found: Product Page URL='{product_page_url}'")
                    return product_page_url

            except Exception as e:
                logging.error(f"Error processing product: {e}")

        logging.info("No matching product found on the listing page.")
        return None  # Return None if no matching product is found
    except TimeoutException:
        logging.error("Timeout while loading the page.")
        return None

# Iterate through each row in the Excel file
for index, row in data.iterrows():
    short_link = row[link_column]
    product_name = row[product_name_column]
    sku = row[sku_column]

    logging.info(f"Processing row {index + 1}: Product Name='{product_name}', SKU='{sku}', Short Link='{short_link}'")

    # Process the listing page
    product_page_url = process_listing_page(short_link, product_name, sku)

    if product_page_url:
        # Update the old link column and status, overwriting any existing text in the old link column
        data.at[index, old_link_column] = product_page_url
        data.at[index, status_column] = "Updated"
    else:
        data.at[index, status_column] = "Not Found"

# Save the updated Excel file to the same file
data.to_excel(excel_file, index=False)
logging.info(f"Updated Excel file saved to: {excel_file}")

# Quit the driver
driver.quit()
