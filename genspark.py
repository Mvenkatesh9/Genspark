import streamlit as st
from pptx.util import Pt 
from pptx.dml.color import RGBColor
from pptx import Presentation
from pptx.util import Inches
from pathlib import Path   
import os
import re
from io import BytesIO
import time


# Conceptual Selenium Code
import tempfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException

user_data_dir = tempfile.mkdtemp()

options = Options()
options.add_argument(f'--user-data-dir={user_data_dir}')
options.add_argument('--headless')  # Optional, if running on server
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

st.set_page_config(page_title="AI PPT Generator", layout="centered")
st.title("🎯 AI Presentation Generator")

topic = st.text_input("Enter your topic:")
presentation_type = st.selectbox("Select presentation type:", ["Academic", "Corporate", "Research"])
detail_level = st.selectbox("Select detail level:", ["High-level", "Detailed"])
slide_count = st.slider("Number of slides:", 3, 15, 5)

def robust_login(driver):
    """Handles locating and interacting with the login form elements."""
    wait = WebDriverWait(driver, MAX_WAIT_TIME)

    try:
        # 1. Locate and fill Email Field
        print("1. Locating and filling Email field...")
        # Use visibility_of_element_located to ensure it's on screen and enabled
        email_field = wait.until(
            EC.visibility_of_element_located((By.ID, EMAIL_ID)),
            message=f"Timeout waiting for Email field (ID: {EMAIL_ID})"
        )
        email_field.send_keys(USER_EMAIL)
        print("✅ Email entered.")

        # 2. Locate and fill Password Field
        print("2. Locating and filling Password field...")
        password_field = wait.until(
            EC.visibility_of_element_located((By.ID, PASSWORD_ID)),
            message=f"Timeout waiting for Password field (ID: {PASSWORD_ID})"
        )
        password_field.send_keys(USER_PASSWORD)
        print("✅ Password entered.")
        time.sleep(5)
        # 3. Locate and click Sign In Button
        print("3. Locating and clicking 'Sign in' button...")
        # Use element_to_be_clickable as it's the strongest check for interaction
        sign_in_button = wait.until(
            EC.element_to_be_clickable((By.ID, SIGN_IN_ID)),
            message=f"Timeout waiting for Sign in button (ID: {SIGN_IN_ID})"
        )
        sign_in_button.click()
        print("✅ 'Sign in' button clicked. Login sequence complete.")

    except (TimeoutException, ElementNotInteractableException) as e:
        print(f"❌ Automation failed during login. Element not found or interactable: {e}")
    except Exception as e:
        print(f"❌ An unexpected error occurred: {e}")

# --- Credentials (Replace with your actual data) ---
USER_EMAIL = "info@cliniv.in" 
USER_PASSWORD = "ClinIV@810" 
MAX_WAIT_TIME = 60 # Increased wait time for safety

# --- Locators ---
EMAIL_ID = "email"
PASSWORD_ID = "password"
SIGN_IN_ID = "next"

if st.button("Generate Presentation"):
    if not topic.strip():
        st.warning("⚠️ Please enter a topic first.")
    else:
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),options=options)
        driver.get("https://www.genspark.ai/") 
        
# Define the maximum wait time
        MAX_WAIT_TIME = 70
# Target the outer div with both class names joined by a dot
        BUTTON_LOCATOR = (By.CSS_SELECTOR, "div.button.signin")

        try:
                print(f"Waiting for the 'Sign in' button to be clickable...")
                wait = WebDriverWait(driver, MAX_WAIT_TIME)
    
    # Wait until the element is clickable
                sign_in_button = wait.until(EC.element_to_be_clickable(BUTTON_LOCATOR))
    
    # Click the button
                sign_in_button.click()
                print("✅ 'Sign in' button clicked successfully.")

        except Exception as e:
                print(f"❌ Error: The 'Sign in' button was not found or clicked. Details: {e}")

        wait = WebDriverWait(driver, 10)
        BUTTON_LOCATOR = (By.ID, "loginWithEmailWrapper")
        login_button = wait.until(EC.element_to_be_clickable(BUTTON_LOCATOR))
    
        # Click the button
        login_button.click()
        robust_login(driver)
        time.sleep(90)
        
        driver.refresh()
        time.sleep(30)

        AI_SLIDE_XPATH = "//div[contains(@style, 'background: rgb(255, 250, 245)')]"

        ai_slide_button = driver.find_element(By.XPATH, AI_SLIDE_XPATH)
    
    # Click the button
        ai_slide_button.click()

        window_handles = driver.window_handles 

# 2. Assume the last handle in the list is the new tab
        new_tab_handle = window_handles[-1]

        

# 3. Switch the driver's focus to the new tab
        driver.switch_to.window(new_tab_handle)
        print("✅ Switched focus to the new tab.")

# Conceptual Selenium Code
        topic_input = topic

# Locate the topic input field (inspect the actual site for ID, name, or XPath)
        
        topic_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "textarea.j-search-input")))        
        topic_field.send_keys(topic_input)

# Click the generation button
        SUBMIT_BUTTON_LOCATOR = (By.CSS_SELECTOR, ".enter-icon-wrapper")

# Wait up to 10 seconds until the button is clickable
        try:
            print("Waiting for the Generate/Send button to be clickable...")
    
            submit_button = wait.until(
            EC.element_to_be_clickable(SUBMIT_BUTTON_LOCATOR))
    
    # Click the button to initiate generation
            submit_button.click()
            print("✅ Generate/Send button clicked successfully.")

        except Exception as e:
            print(f"❌ Error: Could not click the Generate/Send button. Details: {e}")

        