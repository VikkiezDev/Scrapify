import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from sentence_transformers import SentenceTransformer, util
import time

# Load Excel data
file_path = 'your_file.xlsx'  # Replace with your Excel file path
data = pd.read_excel(file_path)

# Initialize the semantic similarity model
model = SentenceTransformer('all-MiniLM-L6-v2')

# Set up Selenium WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

def check_match(row):
    """
    Opens the URL, extracts page content, and compares it with the Excel data.
    """
    url = row['URL']
    expected_sku = row['SKU']
    expected_family = row['Product Family']
    expected_color = row['Color']
    expected_description = row['Description']

    try:
        # Open the URL
        driver.get(url)
        time.sleep(3)  # Wait for the page to load (adjust as needed)

        # Extract details from the page (update locators based on your webpage structure)
        page_sku = driver.find_element(By.ID, 'sku').text.strip()
        page_family = driver.find_element(By.ID, 'product_family').text.strip()
        page_color = driver.find_element(By.ID, 'color').text.strip()
        page_description = driver.find_element(By.ID, 'description').text.strip()

        # Perform exact matches
        sku_match = page_sku == expected_sku
        family_match = page_family.lower() == expected_family.lower()
        color_match = page_color.lower() == expected_color.lower()

        # Perform semantic similarity for the description
        description_embedding_excel = model.encode(expected_description, convert_to_tensor=True)
        description_embedding_page = model.encode(page_description, convert_to_tensor=True)
        similarity = util.pytorch_cos_sim(description_embedding_excel, description_embedding_page).item()

        # Threshold for semantic similarity
        description_match = similarity > 0.8

        # Determine overall match status
        if sku_match and family_match and color_match and description_match:
            return 'Correct'
        else:
            return 'Incorrect'

    except Exception as e:
        print(f"Error processing URL {url}: {e}")
        return 'Error'

# Loop through each row in the Excel data
for index, row in data.iterrows():
    print(f"Processing SKU: {row['SKU']} at URL: {row['URL']}")
    data.loc[index, 'Status'] = check_match(row)

# Save the results to a new Excel file
output_file = 'verification_results.xlsx'
data.to_excel(output_file, index=False)
print(f"Results saved to {output_file}")

# Close the browser
driver.quit()
