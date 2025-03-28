from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import tkinter as tk
import pandas as pd
import pyautogui
import time
import sys
import os
import keyboard
import re

# Excel file path
EXCEL_FILE = "tcg_prices.csv"

# Connect to ChromeDriver Path
if getattr(sys, 'frozen', False):
    chrome_driver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
else:
    chrome_driver_path = r"C:\Program Files\Google\Chrome\Application\chromedriver.exe"

service = Service(chrome_driver_path)
chrome_options = Options()
driver = webdriver.Chrome(service=service, options=chrome_options)

# Launch Chrome and go full screen
driver.get("https://www.tcgplayer.com/")
time.sleep(1)
pyautogui.hotkey('f11')

def get_current_tab_data():
    """Scrapes and returns price data from the TCGPlayer page."""
    driver.switch_to.window(driver.window_handles[0])
    soup = BeautifulSoup(driver.page_source, "html.parser")

    # Market Price
    field1_elem = soup.select_one(".price-guide__near-mint td:nth-child(2) span") 
    market_price = field1_elem.text.strip() if field1_elem and field1_elem.text.strip() else "N/A"

    # Median Price
    field2_elem = soup.select_one(".price-guide__points .price-points__lower td.price-points__lower__right-padding span")
    median_price = field2_elem.text.strip() if field2_elem and field2_elem.text.strip() else "N/A"

    # Card Name (and splitting)
    field3_elem = soup.select_one("#app > div > div > section.marketplace__content > section > div.product-details-container > div.product-details__product > div > h1")
    card_name_full = field3_elem.text.strip() if field3_elem else "N/A"
    match = re.match(r"^(.*?) - (.*?) \((.*?)\)$", card_name_full)
    if match:
        card_name, set_name, set_code = match.groups()
    else:
        card_name, set_name, set_code = "N/A", "N/A", "N/A"
    
    # Rarity
    rarity_elem = soup.select_one("div > div > ul > li:nth-child(1) > div > span")
    rarity = rarity_elem.text.strip() if rarity_elem else "N/A"
    
    #Quantity
    quantity_elem = soup.select_one("#app > div > div > section.marketplace__content > section > div.product-details-container > div.product-details__product > section.product-details__price-guide > div > section.price-guide__points > div > div.price-points__lower > table > tr.price-points__lower__top-padding > td.price-points__lower__right-padding > span")
    quantity = quantity_elem.text.strip() if quantity_elem else "N/A"
    
    #Num Sellers
    numsellers_elem = soup.select_one("#app > div > div > section.marketplace__content > section > div.product-details-container > div.product-details__product > section.product-details__price-guide > div > section.price-guide__points > div > div.price-points__lower > table > tr.price-points__lower__top-padding > td:nth-child(4) > span")
    num_sellers = numsellers_elem.text.strip() if numsellers_elem else "N/A"

    #Num Sold
    numsold_elem = soup.select_one("#app > div > div > section.marketplace__content > section > div.product-details-container > div.product-details__product > section.product-details__price-guide > div > section.price-guide__latest-sales > section > table > tr.sales-data__top-padding > td.sales-data__right-padding > span")
    num_sold = numsold_elem.text.strip() if numsold_elem else "N/A"
    
    #Daily Sold
    dailysold_elem = soup.select_one("#app > div > div > section.marketplace__content > section > div.product-details-container > div.product-details__product > section.product-details__price-guide > div > section.price-guide__latest-sales > section > table > tr.sales-data__top-padding > td:nth-child(4) > span")
    daily_sold = dailysold_elem.text.strip() if dailysold_elem else "N/A"

    return market_price, median_price, card_name_full, rarity, card_name, set_name, set_code, quantity, num_sellers, num_sold, daily_sold

def save_to_excel():
    try:
        value1, value2, value3, value4, value5, value6, value7, value8, value9, value10, value11 = get_current_tab_data()
        new_data = pd.DataFrame([[value5, value6, value7, value4, value1, value2, value8, value9, value10, value11]], 
                                columns=["Card Name", "Set Name", "Set Code", "Rarity", "Market Price", "Median Price", "Quantity", "Num Sellers", "3 Month Sold", "Daily Sold"])

        if os.path.exists(EXCEL_FILE):
            existing_data = pd.read_csv(EXCEL_FILE)
            updated_data = pd.concat([existing_data, new_data], ignore_index=True)
        else:
            updated_data = new_data
        
        updated_data.to_csv(EXCEL_FILE, index=False)
        label.config(text=f" Data saved: {value5}") # Update GUI to show data has been saved.
    
    except Exception as e:
        print(f"Error saving data: {e}")

def update_display():
    try:
        value1, value2, value3, value4, value5, value6, value7, value8, value9, value10, value11 = get_current_tab_data()
        result = f"{value3}\nMarket Price: {value1}\nMedian Price: {value2}\n{value4}"
        label.config(text=result)
    except Exception as whoops:
        label.config(text=f"Error: {whoops}")
    root.after(1000, update_display)  # Refresh data every second

# Setup the GUI
root = tk.Tk()
root.title("TCGPlayer Buying Guide")
root.attributes("-topmost", 1)
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 800
window_height = 200
x_pos = screen_width - window_width - 10
y_pos = 10
root.geometry(f'{window_width}x{window_height}+{x_pos}+{y_pos}')

# Label to display scraped data
label = tk.Label(root, text="", font=("Arial", 14))
label.pack(pady=20)

# Start updating display
update_display()

# Bind data save hotkey
keyboard.add_hotkey("shift+s", save_to_excel)

# Start the Tkinter event loop
root.mainloop()