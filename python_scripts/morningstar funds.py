import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

# Set up Selenium WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--headless")  # Run in background mode
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920x1080")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# URL of the Morningstar Fund Screener
base_url = "https://sg.morningstar.com/sg/screener/fund.aspx"

# Open the page
driver.get(base_url)
time.sleep(5)  # Allow JavaScript to load

### **🔹 Step 1: Remove Any Overlay** ###
try:
    overlay = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CLASS_NAME, "body_blocked"))
    )
    driver.execute_script("arguments[0].remove();", overlay)
    print("✅ Removed page overlay.")
except:
    print("⚠️ No overlay detected. Continuing...")

### **🔹 Step 2: Set 'Show Rows' to 50 Using JavaScript** ###
try:
    dropdown_id = "ec-screener-input-page-size-select"
    driver.execute_script("""
        let dropdown = document.getElementById(arguments[0]);
        dropdown.value = '50';
        dropdown.dispatchEvent(new Event('change', { bubbles: true }));
    """, dropdown_id)

    time.sleep(5)  # Wait for table update
    print("✅ Successfully set 'Show Rows' to 50 using JavaScript!")
except Exception as e:
    print(f"⚠️ Failed to set 'Show Rows' to 50: {e}")

# Dictionary to store data per tab
data_per_tab = {}

# Tabs Mapping
tabs = {
    "Overview": "ec-screener-view-tabs-tab0",
    "Short Term Performance": "ec-screener-view-tabs-tab1",
    "Long Term Performance": "ec-screener-view-tabs-tab2",
    "Fees": "ec-screener-view-tabs-tab3",
    "Portfolio": "ec-screener-view-tabs-tab4",
    "Risk": "ec-screener-view-tabs-tab5",
}

# **Function to Scrape Table Data**
def scrape_table():
    soup = BeautifulSoup(driver.page_source, "html.parser")
    headers = [th.text.strip() for th in soup.find_all("th") if th.text.strip()]
    rows = soup.select("tbody tr")
    page_data = []

    for row in rows:
        cols = row.find_all("td")
        if len(cols) > 1:
            row_data = [col.text.strip() for col in cols]
            page_data.append(row_data)

    return headers, page_data

# **Function to Reset to Page 1 via URL Reload**
def reset_to_page_1():
    print("🔄 Resetting to Page 1 via URL reload...")
    driver.get(base_url)  # Reload the page to ensure we're at Page 1
    time.sleep(5)  # Allow page to load

    # Verify reset worked
    try:
        current_page_element = driver.find_element(By.CSS_SELECTOR, "a.mds-pagination__link.mds-pagination__link--selected")
        current_page = int(current_page_element.text.strip())

        if current_page == 1:
            print("✅ Successfully reset to Page 1.")
        else:
            print(f"⚠️ Page did not reset properly. Expected Page 1, but got {current_page}.")
    
    except Exception as e:
        print(f"⚠️ Failed to reset pagination: {e}")

# **Function to Click Next Page**
def click_next_page():
    try:
        # Find the currently active page number
        current_page_element = driver.find_element(By.CSS_SELECTOR, "a.mds-pagination__link.mds-pagination__link--selected")
        current_page = int(current_page_element.text.strip())
        print(f"🔹 Before Clicking: Current Page: {current_page}")

        # Find all pagination buttons and get the next one
        page_buttons = driver.find_elements(By.CSS_SELECTOR, "a.mds-pagination__link")

        for btn in page_buttons:
            try:
                page_num = int(btn.text.strip())
                if page_num == current_page + 1:
                    # Click the next page button using JavaScript
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(5)  # Ensure new data loads
                    
                    # Verify page change
                    new_page_element = driver.find_element(By.CSS_SELECTOR, "a.mds-pagination__link.mds-pagination__link--selected")
                    new_page = int(new_page_element.text.strip())

                    if new_page == page_num:
                        print(f"✅ Successfully moved to Page {new_page}.")
                        return True
                    else:
                        print(f"⚠️ Page did not change properly. Expected {page_num}, but got {new_page}.")
                        return False
            except:
                continue

    except Exception as e:
        print(f"⚠️ Failed to navigate to the next page: {e}")
    return False

# **Scraping Logic for Each Tab**
for tab_name, tab_id in tabs.items():
    print(f"\n🔄 Switching to {tab_name} tab...")

    # Reset to Page 1 before switching tabs
    reset_to_page_1()

    # Click the tab
    try:
        tab_element = driver.find_element(By.ID, tab_id)
        driver.execute_script("arguments[0].click();", tab_element)
        time.sleep(5)  # Allow data to load
        print(f"✅ Switched to {tab_name} tab.")
    except Exception as e:
        print(f"❌ Failed to switch to {tab_name}: {e}")
        continue

    funds_data = []
    page_count = 0

    # **Scrape first 5 pages**
    while page_count < 5:
        page_count += 1
        print(f"🔄 Scraping page {page_count} of {tab_name}...")

        headers, page_data = scrape_table()

        if not page_data:
            print("⚠️ No data detected, stopping.")
            break

        funds_data.extend(page_data)

        success = click_next_page()
        if not success:
            print(f"⚠️ No more pages in {tab_name}. Moving to next tab.")
            break

        time.sleep(5)  # Wait for the table to update

    # Save data
    data_per_tab[tab_name] = pd.DataFrame(funds_data, columns=headers)

# Save all data into Excel
with pd.ExcelWriter("morningstar_funds.xlsx") as writer:
    for sheet_name, df in data_per_tab.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("🎉 Scraping complete! Data saved as 'morningstar_funds.xlsx'.")

# Close the Selenium driver
driver.quit()
