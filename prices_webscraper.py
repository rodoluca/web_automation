#!/usr/bin/env python
# coding: utf-8

# In[6]:


# Create a web browser
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import pandas as pd

# Create the web browser
browser = webdriver.Chrome()

# Import/visualize the database
products_table = pd.read_excel("searches.xlsx")
display(products_table)


# ### Definition of search functions in Google and Buscape

# In[2]:


import time

def search_google_shopping(browser, product, banned_terms, min_price, max_price):
    # Go to Google
    browser.get("https://www.google.com/")
    
    # Process values from the table
    product = product.lower()
    banned_terms = banned_terms.lower()
    banned_terms_list = banned_terms.split(" ")
    product_terms_list = product.split(" ")
    max_price = float(max_price)
    min_price = float(min_price)

    # Search for the product name on Google
    browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(product)
    browser.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

    # Click on the shopping tab
    elements = browser.find_elements(By.CLASS_NAME, 'hdtb-mitem')
    for item in elements:
        if "Shopping" in item.text:
            item.click()
            break

    # Get the list of results from Google Shopping search
    results_list = browser.find_elements(By.CLASS_NAME, 'sh-dgr__grid-result')

    # For each result, check if it meets all our conditions
    offers_list = []  # List that the function will return as the response
    for result in results_list:
        name = result.find_element(By.CLASS_NAME, 'Xjkr3b').text
        name = name.lower()

        # Name verification - if the name contains any banned terms
        has_banned_terms = False
        for word in banned_terms_list:
            if word in name:
                has_banned_terms = True

        # Check if the name contains all the terms from the product name
        has_all_product_terms = True
        for word in product_terms_list:
            if word not in name:
                has_all_product_terms = False

        if not has_banned_terms and has_all_product_terms:  # Check the name
            try:
                price = result.find_element(By.CLASS_NAME, 'a8Pemb').text
                price = price.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                price = float(price)
                # Check if the price is within the minimum and maximum
                if min_price <= price <= max_price:
                    link_element = result.find_element(By.CLASS_NAME, 'aULzUe')
                    parent_element = link_element.find_element(By.XPATH, '..')
                    link = parent_element.get_attribute('href')
                    offers_list.append((name, price, link))
            except:
                continue


    return offers_list
    
    

def search_buscape(browser, product, banned_terms, min_price, max_price):
    # Process function values
    max_price = float(max_price)
    min_price = float(min_price)
    product = product.lower()
    banned_terms = banned_terms.lower()
    banned_terms_list = banned_terms.split(" ")
    product_terms_list = product.split(" ")

    
#     Go to Buscape
    browser.get("https://www.buscape.com.br/")

    Search for the product on Buscape
    browser.find_element(By.CLASS_NAME, 'search-bar__text-box').send_keys(product, Keys.ENTER)

#     Get the list of search results from Buscape
    time.sleep(5)
    results_list = browser.find_elements(By.CLASS_NAME, 'Cell_Content__1630r')

#     For each result
    offers_list = []
    for result in results_list:
        try:
            price = result.find_element(By.CLASS_NAME, 'CellPrice_MainValue__3s0iP').text
            name = result.get_attribute('title')
            name = name.lower()
            link = result.get_attribute('href')
            
            # Name verification - if the name contains any banned terms
            has_banned_terms = False
            for word in banned_terms_list:
                if word in name:
                    has_banned_terms = True  

            # Check if the name contains all the terms from the product name
            has_all_product_terms = True
            for word in product_terms_list:
                if word not in name:
                    has_all_product_terms = False            

            if not has_banned_terms and has_all_product_terms:
                price = price.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                price = float(price)
                if min_price <= price <= max_price:
                    offers_list.append((name, price, link))
        except:
            pass
    return offers_list


# ### Building our list of found offers

# In[8]:


offers_table = pd.DataFrame()

for row in products_table.index:
    product = products_table.loc[row, "Name"]
    banned_terms = products_table.loc[row, "Banned terms"]
    min_price = products_table.loc[row, "Minimum price"]
    max_price = products_table.loc[row, "Maximum price"]
    
    google_shopping_offers_list = search_google_shopping(browser, product, banned_terms, min_price, max_price)
    if google_shopping_offers_list:
        google_shopping_table = pd.DataFrame(google_shopping_offers_list, columns=['product', 'price', 'link'])
        offers_table = offers_table.append(google_shopping_table)
    else:
        google_shopping_table = None

    buscape_offers_list = search_buscape(browser, product, banned_terms, min_price, max_price)
    if buscape_offers_list:
        buscape_table = pd.DataFrame(buscape_offers_list, columns=['product', 'price', 'link'])
        offers_table = offers_table.append(buscape_table)
    else:
        buscape_table = None

display(offers_table)


# ### Exporting the offers database to Excel

# In[9]:


# Export to Excel
offers_table = offers_table.reset_index(drop=True)
offers_table.to_excel("Offers.xlsx", index=False)


# ### Sending the email

# In[ ]:


# Sending the email
import win32com.client as win32

Checking if there are any offers in the offers table
if len(offers_table.index) > 0:
# Sending email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'pythonimpressionador@gmail.com'
    mail.Subject = 'Product(s) Found within Desired Price Range'
    mail.HTMLBody = f"""
    <p>Dear Sir/Madam,</p>
    <p>We have found some products on offer within the desired price range. Please find the details in the table below:</p>
    {offers_table.to_html(index=False)}
    <p>If you have any questions, feel free to reach out.</p>
    <p>Best regards,</p>
    """
    mail.Send()
    
browser.quit()


# In[ ]:




