import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ===== CONFIG =====
FILE_PATH = "cities.xlsx"
URL = "https://www.icaionlineregistration.org/launchbatchdetail.aspx"

# ===== LOAD EXCEL =====
df_cities = pd.read_excel(FILE_PATH)
if 'City' not in df_cities.columns:
    raise Exception("Excel must contain 'City' column")

# ===== SETUP SELENIUM =====
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)
driver.get(URL)
print("Opening site... waiting 20 sec")
time.sleep(20)

output_rows = []

for city_name in df_cities['City']:
    city = str(city_name).strip()
    print(f"\nProcessing city: {city}")
    try:
        # Select Southern
        region = wait.until(EC.presence_of_element_located((By.ID, "ddl_reg")))
        Select(region).select_by_visible_text("Southern")
        time.sleep(6)

        # Select city
        city_dropdown = wait.until(EC.presence_of_element_located((By.ID, "ddlPou")))
        select_city = Select(city_dropdown)
        found = False
        for option in select_city.options:
            if city.lower() in option.text.lower():
                select_city.select_by_visible_text(option.text)
                found = True
                break
        if not found:
            print("City not found in dropdown")
            output_rows.append([city, "No records", "No records", "No records", "No records"])
            driver.refresh()
            time.sleep(5)
            continue
        time.sleep(2)

        # Click Get List
        get_btn = wait.until(EC.element_to_be_clickable((By.ID, "btn_getlist")))
        driver.execute_script("arguments[0].click();", get_btn)
        print("Clicked Get List")
        time.sleep(6)

        # Extract multiple batches dynamically
        idx = 0
        while True:
            try:
                batch_size = driver.find_element(By.ID, f"GridView1_lblPublishSize_{idx}").text
                from_date = driver.find_element(By.ID, f"GridView1_lblFromDate_{idx}").text
                to_date = driver.find_element(By.ID, f"GridView1_lblTodate_{idx}").text
                timing = driver.find_element(By.ID, f"GridView1_lblBatchTiming_{idx}").text
                output_rows.append([city, batch_size, from_date, to_date, timing])
                idx += 1
            except:
                break

        if idx == 0:
            print("No records found for city")
            output_rows.append([city, "No records", "No records", "No records", "No records"])

    except Exception as e:
        print("Error:", e)
        output_rows.append([city, "Error", "Error", "Error", "Error"])

    driver.refresh()
    print("Page refreshed")
    time.sleep(8)

# ===== SAVE BACK TO SAME EXCEL =====
df_output = pd.DataFrame(output_rows, columns=['City', 'Batch Size', 'From Date', 'To Date', 'Timing'])
df_output.to_excel(FILE_PATH, index=False)
driver.quit()
print("\nAll cities processed. Data saved back to", FILE_PATH)
