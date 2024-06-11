import shutil

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd
import os
import requests

# Load the Excel file
df = pd.read_excel('SKUs_To_Be_Scraped.xlsx') # Assuming 'part_numbers.xlsx' is the file name
part_numbers = df['SKU'] # Assuming 'Part Number' is the column name
attribute_df = pd.read_excel('SKUs_To_Be_Scraped.xlsx',index_col=0)
mapping_df = pd.read_excel('SKUs_To_Be_Scraped.xlsx',index_col=0)

attribute_df['Website SKU'] = ''
attribute_df['Description'] = ''
# Initialize the  browser
browser = webdriver.Firefox()
browser.maximize_window()

# Website To Be Searched
url = "https://supplierwebsite.com/"

# Not found text/results text
not_found_text = "You might want to check that URL again or head over to"
found_text = "Search results for"
browser.get(url)
# Cookies accept, if necessary
try:
    cookiesAccept = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, "btn-cookie-allow")))
    cookiesAccept.click()
except:
    print("No cookies or wrong cookie accept button ")
mapping_names = []
attribute_names = []
# Iterate through part numbers
for part_number in part_numbers:

    # Open the website
    browser.get(url+part_number)

    # Check if page is search results page or no results available
    if not_found_text in browser.page_source:
        print(f"'{part_number}' search returned no results")
        attribute_df.loc[part_number, 'Website SKU'] = "Wrong URL or Not Found"
    elif found_text in browser.page_source:
        print(f"'{part_number}' search returned some results")
    else:
        #print("Redirected")

        # Images
        image_URLS = []
        time.sleep(5)
        primary_image = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CLASS_NAME,"fotorama__img")))
        product_media = browser.find_elements(By.CLASS_NAME,"fotorama__img")
        primary_url = str(primary_image.get_attribute("src")).rsplit('/',1)[0]
        for image_thumb in product_media:
            image_URLS.extend([primary_url+"/"+str(image_thumb.get_attribute("src")).rsplit('/',1)[1]])
        # Keep uniques
        image_URLS = list(set(image_URLS))

        # Download Images
        folder_path = "downloaded_images/"
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        image_no = 1
        for i, image_url in enumerate(image_URLS):
            response = requests.get(image_url)
            if response.status_code == 200:
                # filename
                filename = os.path.join(folder_path, f'{part_number}_{i + 1}.jpg')

                # Save the image
                with open(filename, 'wb') as file:
                    file.write(response.content)
                print(f'Saved {filename}')
                if "Image_" + str(image_no) not in mapping_names:
                    mapping_names.append("Image_" + str(image_no))
                    mapping_df["Image_" + str(image_no)] = ''
                mapping_df.loc[part_number, "Image_" + str(image_no)] = f'{part_number}_{i + 1}.jpg'
            else:
                print(f'Failed to download {image_url}')
            image_no += 1
            mapping_df.to_excel("Mapping_File.xlsx")

        # Attributes
        product_info_main = WebDriverWait(browser, 10).until(EC.presence_of_element_located((
            By.CLASS_NAME,"product-info-main")))
        website_sku = product_info_main.find_element(By.CLASS_NAME,"page-title")
        attribute_df.loc[part_number, 'Website SKU'] = website_sku.text
        print(website_sku.text)

        product_description = browser.find_element(By.ID,"description")
        attribute_df.loc[part_number, 'Description'] = product_description.text
        #print(product_description.text)

        attributes_button = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((
            By.ID,"tab-label-additional-title")))
        browser.execute_script("arguments[0].click();", attributes_button)
        time.sleep(2)

        attributes_table = browser.find_element(By.ID,"product-attribute-specs-table").find_elements(By.TAG_NAME,"tr")

        for row in attributes_table:
            attribute_name = row.find_element(By.TAG_NAME,"th").text
            attribute_value = row.find_element(By.TAG_NAME,"td").text
            if attribute_name not in attribute_names:
                attribute_names.append(attribute_name)
                attribute_df[attribute_name] = ''
            attribute_df.loc[part_number, attribute_name] = attribute_value

        attribute_df.to_excel("Scraped_Attributes.xlsx")

# Close the browser when done
browser.quit()