import os
import sys
import time
import re
from playwright.sync_api import sync_playwright, expect

USERNAME = "tam.doan@myteamsolution.com.vn"
PASSWORD = "Doantam@123"
DOWNLOAD_DIR = r"G:\My Drive\Dữ liệu nhân sự\Workload\Construction Team"

SESSION_FILE = os.path.join(os.path.dirname(__file__), "auth_state.json")

def main():
    print("Starting automated data download from MyDaily...")

    # Ensure the download directory exists
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)

        # Load session state if it exists
        if os.path.exists(SESSION_FILE):
            print("Found existing session file, loading state...")
            context = browser.new_context(storage_state=SESSION_FILE, accept_downloads=True)
        else:
            context = browser.new_context(accept_downloads=True)

        page = context.new_page()

        try:
            print("Navigating to MyDaily...")
            page.goto("https://mydaily.myteamsolution.com.vn/login", wait_until="domcontentloaded")

            # Target locator for either login form (#email) or main page (Data link)
            target = page.locator("#email").or_(page.get_by_role("link", name="Data", exact=True))

            # Check if initial load renders blank; if so, trigger a reload (F5)
            try:
                target.wait_for(timeout=1000)
            except Exception:
                print("Page blank on initial load, performing page reload (F5)...")
                page.reload(wait_until="domcontentloaded")
                target.wait_for(timeout=15000)

            # Check if login form is present
            if page.locator("#email").is_visible():
                print("Logging in to MyDaily...")
                page.locator("#email").fill(USERNAME)
                page.locator("#password").fill(PASSWORD)
                page.get_by_role("button", name="Login").click()

                # Wait for main page element after login
                page.get_by_role("link", name="Data", exact=True).wait_for(timeout=20000)

                # Save session state after successful login
                context.storage_state(path=SESSION_FILE)
                print("Session saved successfully!")
            else:
                print("Already logged in using saved session!")

            print("Navigating to export section...")
            page.get_by_role("link", name="Data", exact=True).click()
            page.get_by_role("button", name="All Reports Export jobs, sub-").click()
            page.get_by_role("button", name="Next: Select Time Period →").click()

            print("Selecting time period: past 6 months...")
            page.get_by_role("combobox").nth(1).select_option("past_six_months")
            page.get_by_role("button", name="Next: Preview Data →").click()
            page.get_by_role("button", name="All").click()

            print("Exporting to Excel...")
            with page.expect_download(timeout=60000) as download_info:
                page.get_by_role("button", name="Export to Excel").first.click(timeout=60000)

            download = download_info.value
            ext = os.path.splitext(download.suggested_filename)[1] or ".xlsx"
            filename = f"Report_Past Month{ext}"
            save_path = os.path.join(DOWNLOAD_DIR, filename)

            print(f"Saving file to: {save_path}")
            download.save_as(save_path)

            print("✅ Download completed successfully!")

        except Exception as e:
            print(f"❌ An error occurred: {e}")
            if os.path.exists(SESSION_FILE):
                try:
                    os.remove(SESSION_FILE)
                    print("Cleared saved session file so next attempt can login fresh.")
                except Exception:
                    pass
        finally:
            time.sleep(2)
            context.close()
            browser.close()

if __name__ == "__main__":
    main()
