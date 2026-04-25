import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ---------------- DATA ---------------- #
file_path = "product list.xlsx"
df = pd.read_excel(file_path)

# ---------------- SETUP ---------------- #
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)

driver.get("https://www.amazon.in/")

# =========================================================
# 1️⃣ CLICK SIGN IN
# =========================================================
sign_in = wait.until(
    EC.element_to_be_clickable((By.ID, "nav-link-accountList"))
)
sign_in.click()

# =========================================================
# 2️⃣ ENTER EMAIL / PHONE
# =========================================================
email_box = wait.until(
    EC.presence_of_element_located((By.ID, "ap_email_login"))
)
email_box.clear()
email_box.send_keys("Your Mail or Phone Number")

# fallback if needed
try:
    continue_btn = driver.find_element(By.ID, "continue")
except:
    continue_btn = wait.until(
        EC.element_to_be_clickable((By.ID, "continue"))
    )

continue_btn.click()

# =========================================================
# 3️⃣ ENTER PASSWORD
# =========================================================
password_box = wait.until(
    EC.presence_of_element_located((By.ID, "ap_password"))
)
password_box.clear()
password_box.send_keys("Your Password")

sign_in_btn = wait.until(
    EC.element_to_be_clickable((By.ID, "signInSubmit"))
)
sign_in_btn.click()

# ✅ Confirm login success
wait.until(
    EC.presence_of_element_located((By.ID, "nav-link-accountList"))
)

print("✅ Login Successful")

# =========================================================
# 4️⃣ PRODUCT LOOP (SEARCH + ADD TO CART)
# =========================================================
for product in df['product']:
    try:
        driver.get("https://www.amazon.in/")

        # Search product
        search_box = wait.until(
            EC.presence_of_element_located((By.ID, "twotabsearchtextbox"))
        )
        search_box.clear()
        search_box.send_keys(str(product))
        search_box.send_keys(Keys.RETURN)

        # Click first product
        first_product = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "h2 a"))
        )
        first_product.click()

        # Switch tab if opened
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])

        print(f"🔍 Opened product page: {product}")

        # ⏸️ MANUAL STEP (YOU CLICK ADD TO CART)
        input("👉 Click 'Add to Cart' manually, then press ENTER to continue...")

        print(f"🛒 Confirmed manual add to cart: {product}")

        time.sleep(2)

        # Close product tab safely
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

    except Exception as e:
        print(f"❌ Error with {product}: {e}")
        driver.get("https://www.amazon.in/")
        time.sleep(2)
# =========================================================
# 5️⃣ SIGN OUT
# =========================================================
try:
    menu = wait.until(
        EC.element_to_be_clickable((By.ID, "nav-hamburger-menu"))
    )
    menu.click()

    sign_out = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Sign Out')]"))
    )
    sign_out.click()

    print("🚪 Signed out successfully")

except Exception as e:
    print(f"❌ Sign out failed: {e}")

# ---------------- CLOSE ---------------- #
driver.quit()

