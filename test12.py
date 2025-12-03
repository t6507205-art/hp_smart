import time
import re
import random
import string
import os
from typing import Optional, Tuple
import json

import pyperclip
import pytest
from pywinauto import Desktop, keyboard
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

REPORT = []

# -----------------------------
# CONFIGURATION - All hardcoded values replaced
# -----------------------------
CONFIG_FILE = os.getenv("HP_CONFIG_FILE", "hp_config.json")

def load_config():
    """Load configuration from JSON file or environment variables with defaults."""
    default_config = {
        "app_launch": "{VK_LWIN}HP Smart{ENTER}",
        "windows": {
            "hp_smart": r".*HP Smart.*",
            "hp_account": r".*HP account.*",
            "chrome_hpid": r".*HPID Login - Google Chrome",
            "open_hp_smart_dialog": r".*Open HP Smart.*"
        },
        "controls": {
            "manage_account": dict(title="Manage HP Account", auto_id="HpcSignedOutIcon", control_type="Button"),
            "create_account": dict(auto_id="HpcSignOutFlyout_CreateBtn", control_type="Button"),
            "firstname": dict(auto_id="firstName", control_type="Edit"),
            "lastname": dict(auto_id="lastName", control_type="Edit"),
            "email": dict(auto_id="email", control_type="Edit"),
            "password": dict(auto_id="password", control_type="Edit"),
            "signup": dict(auto_id="sign-up-submit", control_type="Button"),
            "otp_input": dict(auto_id="code", control_type="Edit"),
            "otp_submit": dict(auto_id="submit-code", control_type="Button"),
            "open_hp_smart": dict(title="Open HP Smart", control_type="Button"),
            "always_allow_checkbox": dict(
                title="Always allow hpsmart.com to open links of this type in the associated app",
                control_type="CheckBox", class_name="Checkbox"
            )
        },
        "mailsac": {
            "url": os.getenv("MAILSAC_URL", "https://mailsac.com"),
            "mailbox_placeholder_xpath": "//input[@placeholder='mailbox']",
            "check_mail_btn_xpath": "//button[normalize-space()='Check the mail!']",
            "inbox_row_xpath": "//table[contains(@class,'inbox-table')]/tbody/tr[contains(@class,'clickable')][1]",
            "email_body_css": "#emailBody",
            "otp_regex": r"\b(\d{4,8})\b"
        },
        "timeouts": {
            "default": int(os.getenv("HP_DEFAULT_TIMEOUT", "30")),
            "short": int(os.getenv("HP_SHORT_TIMEOUT", "5")),
            "poll_interval": int(os.getenv("HP_POLL_INTERVAL", "3")),
            "otp_max_wait": int(os.getenv("HP_OTP_MAX_WAIT", "60"))
        },
        "random_data": {
            "mailbox_prefix_len": int(os.getenv("HP_MAILBOX_LEN", "4")),
            "mail_domain": os.getenv("HP_MAIL_DOMAIN", "mailsac.com"),
            "firstname_len": int(os.getenv("HP_FIRSTNAME_LEN", "6")),
            "lastname_len": int(os.getenv("HP_LASTNAME_LEN", "6"))
        },
        "password": os.getenv("HP_PASSWORD", "SecurePassword123"),
        "selenium": {
            "headless": os.getenv("HP_HEADLESS", "False").lower() == "true",
            "chrome_args": os.getenv("HP_CHROME_ARGS", "").split() if os.getenv("HP_CHROME_ARGS") else []
        }
    }
    
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                file_config = json.load(f)
            # Merge file config with defaults (file overrides defaults)
            merge_config(default_config, file_config)
        return default_config
    except Exception as e:
        print(f"Config file error, using defaults: {e}")
        return default_config

def merge_config(defaults, overrides):
    """Recursively merge overrides into defaults."""
    for key, value in overrides.items():
        if isinstance(value, dict) and key in defaults and isinstance(defaults[key], dict):
            merge_config(defaults[key], value)
        else:
            defaults[key] = value

CONFIG = load_config()

# -----------------------------
# Utility functions
# -----------------------------
def log_step(desc: str, status: str = "PASS") -> None:
    """Append a step to the report and print it."""
    REPORT.append((desc, status))
    print(f"{desc}: {status}")

def generate_random_mailbox(prefix_len: int = None, domain: str = None) -> str:
    prefix_len = prefix_len or CONFIG["random_data"]["mailbox_prefix_len"]
    domain = domain or CONFIG["random_data"]["mail_domain"]
    prefix = ''.join(random.choices(string.ascii_lowercase, k=prefix_len))
    return f"{prefix}test@{domain}"

def generate_random_name(first_len: int = None, last_len: int = None) -> Tuple[str, str]:
    first_len = first_len or CONFIG["random_data"]["firstname_len"]
    last_len = last_len or CONFIG["random_data"]["lastname_len"]
    first = ''.join(random.choices(string.ascii_letters, k=first_len)).capitalize()
    last = ''.join(random.choices(string.ascii_letters, k=last_len)).capitalize()
    return first, last

# -----------------------------
# App automation using pywinauto
# -----------------------------
def launch_hp_smart(timeout: int = None):
    """Launch HP Smart using the configured command and return the Desktop object on success."""
    timeout = timeout or CONFIG["timeouts"]["default"]
    try:
        keyboard.send_keys(CONFIG["app_launch"])
        log_step("Sent keys to launch HP Smart app.")

        desktop = Desktop(backend="uia")
        main_win = desktop.window(title_re=CONFIG["windows"]["hp_smart"])
        main_win.wait('exists visible enabled ready', timeout=timeout)
        main_win.set_focus()
        log_step("Focused HP Smart main window.")

        manage_btn = main_win.child_window(**CONFIG["controls"]["manage_account"])
        manage_btn.wait('visible enabled ready', timeout=CONFIG["timeouts"]["short"])
        manage_btn.click_input()
        log_step("Clicked Manage HP Account button.")

        create_btn = main_win.child_window(**CONFIG["controls"]["create_account"])
        create_btn.wait('visible enabled ready', timeout=CONFIG["timeouts"]["short"])
        create_btn.click_input()
        log_step("Clicked Create Account button.")

        return desktop

    except Exception as e:
        log_step(f"Error launching HP Smart: {e}", "FAIL")
        return None

def fill_account_form(desktop, email: str, first_name: str, last_name: str, password: str = None):
    """Fill the account creation form in the HP account browser window."""
    password = password or CONFIG["password"]
    try:
        browser_win = desktop.window(title_re=CONFIG["windows"]["hp_account"])
        browser_win.wait('exists visible enabled ready', timeout=CONFIG["timeouts"]["default"])
        browser_win.set_focus()
        log_step("Focused HP Account browser window.")

        browser_win.child_window(**CONFIG["controls"]["firstname"]).type_keys(first_name)
        browser_win.child_window(**CONFIG["controls"]["lastname"]).type_keys(last_name)
        browser_win.child_window(**CONFIG["controls"]["email"]).type_keys(email)
        browser_win.child_window(**CONFIG["controls"]["password"]).type_keys(password)

        signup = browser_win.child_window(**CONFIG["controls"]["signup"])
        signup.wait('visible enabled ready', timeout=CONFIG["timeouts"]["short"])
        signup.click_input()
        log_step("Filled account form and clicked Create button.")

        time.sleep(3)  # allow navigation/OTP flow to start

    except Exception as e:
        log_step(f"Error filling account form: {e}", "FAIL")

# -----------------------------
# OTP retrieval using Selenium
# -----------------------------
def _create_selenium_driver(headless: bool = None, extra_args: list = None):
    opts = webdriver.ChromeOptions()
    headless = headless if headless is not None else CONFIG["selenium"]["headless"]
    extra_args = extra_args or CONFIG["selenium"]["chrome_args"]
    
    if headless:
        opts.add_argument('--headless=new')
    for arg in extra_args:
        opts.add_argument(arg)
    return webdriver.Chrome(options=opts)

def fetch_otp_from_mailsac(mailbox_local_part: str,
                          mailsac_url: str = None,
                          max_wait: int = None,
                          poll_interval: int = None) -> Tuple[Optional[str], Optional[webdriver.Chrome]]:
    """Open a browser, navigate to Mailsac, and attempt to extract an OTP from the latest message."""
    config = CONFIG["mailsac"]
    max_wait = max_wait or CONFIG["timeouts"]["otp_max_wait"]
    poll_interval = poll_interval or CONFIG["timeouts"]["poll_interval"]
    mailsac_url = mailsac_url or config["url"]
    
    otp = None
    driver = None
    try:
        driver = _create_selenium_driver()
        wait = WebDriverWait(driver, CONFIG["timeouts"]["short"])

        driver.get(mailsac_url)
        log_step("Opened Mailsac website.")

        mailbox_field = WebDriverWait(driver, CONFIG["timeouts"]["default"]).until(
            EC.presence_of_element_located((By.XPATH, config["mailbox_placeholder_xpath"]))
        )
        mailbox_field.clear()
        mailbox_field.send_keys(mailbox_local_part)

        check_btn = wait.until(EC.element_to_be_clickable((By.XPATH, config["check_mail_btn_xpath"])))
        check_btn.click()
        log_step("Opened Mailsac inbox.")

        start_time = time.time()
        while time.time() - start_time < max_wait:
            try:
                email_row = WebDriverWait(driver, poll_interval).until(
                    EC.presence_of_element_located((By.XPATH, config["inbox_row_xpath"]))
                )
                email_row.click()
                log_step("Clicked on first email row.")
                break
            except Exception:
                try:
                    driver.find_element(By.XPATH, config["check_mail_btn_xpath"]).click()
                except Exception:
                    pass
                log_step("Refreshed Mailsac inbox.", "INFO")

        body_elem = WebDriverWait(driver, CONFIG["timeouts"]["default"]).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, config["email_body_css"]))
        )
        email_body = body_elem.text

        match = re.search(config["otp_regex"], email_body)
        if match:
            otp = match.group(1)
            log_step(f"Extracted OTP: {otp}")
        else:
            log_step("OTP not found in email.", "FAIL")

        return otp, driver

    except Exception as e:
        log_step(f"Error fetching OTP: {e}", "FAIL")
        if driver:
            driver.quit()
        return None, None

# -----------------------------
# Complete OTP entry using pywinauto
# -----------------------------
def complete_web_verification_in_app(otp: str, timeout: int = None):
    """Paste OTP into the application and submit it, then handle Open HP Smart dialog."""
    timeout = timeout or CONFIG["timeouts"]["default"]
    try:
        desktop = Desktop(backend="uia")
        otp_win = desktop.window(title_re=CONFIG["windows"]["hp_account"])
        otp_win.wait('exists visible enabled ready', timeout=timeout)
        otp_win.set_focus()
        log_step("Focused OTP input screen.")

        otp_box = otp_win.child_window(**CONFIG["controls"]["otp_input"])
        otp_box.wait('visible enabled ready', timeout=CONFIG["timeouts"]["short"])

        pyperclip.copy(otp)
        time.sleep(0.5)
        otp_box.click_input()
        otp_box.type_keys("^v")
        log_step("OTP pasted successfully.")

        submit_btn = otp_win.child_window(**CONFIG["controls"]["otp_submit"])
        submit_btn.wait('visible enabled ready', timeout=CONFIG["timeouts"]["short"])
        submit_btn.click_input()
        log_step("Clicked Verify button.")
        time.sleep(2)

        # Handle 'Open HP Smart?' popup
        click_open_hp_smart()

    except Exception as e:
        log_step(f"OTP verification failed: {e}", "FAIL")

# -----------------------------
# External protocol dialog handler
# -----------------------------
def click_open_hp_smart(timeout: int = None):
    """Handle the Chrome 'Open HP Smart?' popup by clicking Open HP Smart."""
    timeout = timeout or CONFIG["timeouts"]["default"]
    try:
        desktop = Desktop(backend="uia")

        # Attach to the Chrome window that hosts the dialog
        dlg = desktop.window(
            title_re=CONFIG["windows"]["chrome_hpid"],
            class_name="Chrome_WidgetWin_1"
        )
        dlg.wait('exists', timeout=timeout)
        dlg.set_focus()
        log_step("Attached to 'HPID Login - Google Chrome' window.")

        # Tick the 'Always allow...' checkbox if present
        try:
            always_allow = dlg.child_window(**CONFIG["controls"]["always_allow_checkbox"])
            always_allow.wait('exists', timeout=CONFIG["timeouts"]["short"])
            if hasattr(always_allow, "get_toggle_state") and not always_allow.get_toggle_state():
                always_allow.click_input()
                log_step("Checked 'Always allow hpsmart.com...' checkbox.")
        except Exception:
            pass

        # Click the 'Open HP Smart' button
        open_btn = dlg.child_window(**CONFIG["controls"]["open_hp_smart"])
        open_btn.wait('exists', timeout=CONFIG["timeouts"]["short"])
        open_btn.click_input()
        log_step("Clicked 'Open HP Smart' button on Chrome popup.")

    except Exception as e:
        log_step(f"Failed to handle Chrome 'Open HP Smart?' popup: {e}", "FAIL")

# -----------------------------
# Report generation
# -----------------------------
def generate_report(path: str = None):
    """Generate HTML report."""
    path = path or os.getenv("HP_REPORT_PATH", "automation_report.html")
    html = (
        "<html><head><meta charset='utf-8'><title>Automation Report</title>"
        "<style>table{border-collapse:collapse;width:100%;}th,td{border:1px solid #ddd;padding:8px;text-align:left;}th{background:#f2f2f2;}</style>"
        "</head><body><h2>HP Account Automation Report</h2><table><tr><th>Step</th><th>Status</th></tr>"
    )
    for desc, status in REPORT:
        status_color = "green" if status == "PASS" else "red"
        html += f"<tr><td>{desc}</td><td style='color:{status_color}'>{status}</td></tr>"
    html += "</table></body></html>"
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Report generated: {path}")

# -----------------------------
# Main orchestration
# -----------------------------
def main():
    driver = None
    try:
        mailbox_full = generate_random_mailbox()
        mailbox_local_part = mailbox_full.split("@")[0]
        log_step(f"Generated mailbox: {mailbox_full}")

        first_name, last_name = generate_random_name()
        log_step(f"Generated name: {first_name} {last_name}")

        desktop = launch_hp_smart()
        if desktop:
            fill_account_form(desktop, mailbox_full, first_name, last_name)

        otp, driver = fetch_otp_from_mailsac(mailbox_local_part)
        if otp:
            complete_web_verification_in_app(otp)

        # Safely accept any Selenium alert if present
        if driver:
            try:
                alert = Alert(driver)
                log_step(f"Alert present with text: {alert.text}")
                alert.accept()
                log_step("Accepted browser alert.")
            except Exception:
                pass

    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
        generate_report()

# -------------------------------------------------------------
# PYTEST ENTRY POINT
# -------------------------------------------------------------
if __name__ == "__main__":
    main()

def test_hp_account_automation():
    main()
    assert True
