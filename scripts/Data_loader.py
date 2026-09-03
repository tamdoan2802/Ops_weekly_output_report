import os
import sys
import time
from playwright.sync_api import sync_playwright

USERNAME = "tam.doan@myteamsolution.com.vn"
PASSWORD = "Doantam@123"
DOWNLOAD_DIR = r"G:\My Drive\Dữ liệu nhân sự\Workload\Construction Team"
SESSION_FILE = os.path.join(os.path.dirname(__file__), "auth_state.json")

def download_report(context, report_type_name, output_filename):
    page = context.new_page()
    page.set_default_timeout(90000)
    
    try:
        page.goto("https://mydaily.myteamsolution.com.vn/login", wait_until="domcontentloaded", timeout=60000)
        
        target = page.locator("#email").or_(page.get_by_role("link", name="Data", exact=True))
        try:
            target.wait_for(timeout=5000)
        except Exception:
            page.reload(wait_until="domcontentloaded", timeout=60000)
            target.wait_for(timeout=20000)

        if page.locator("#email").is_visible():
            print("   Logging in to MyDaily...")
            page.locator("#email").fill(USERNAME)
            page.locator("#password").fill(PASSWORD)
            page.get_by_role("button", name="Login").click(timeout=30000)
            page.get_by_role("link", name="Data", exact=True).wait_for(timeout=30000)
            context.storage_state(path=SESSION_FILE)

        # Go to Data
        page.get_by_role("link", name="Data", exact=True).click(timeout=30000)
        time.sleep(2)

        # Select Report Type
        btn_report = page.locator(f"button:has-text('{report_type_name}')").first
        btn_report.wait_for(timeout=30000)
        btn_report.click()

        # Select Past 6 Months
        page.get_by_role("button", name="Next: Select Time Period →").click(timeout=30000)
        page.get_by_role("combobox").nth(1).select_option("past_six_months")

        # Preview Data (Wait for aggregation)
        page.get_by_role("button", name="Next: Preview Data →").click(timeout=30000)
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        time.sleep(5) # Buffer for complete server data assembly

        try:
            page.get_by_role("button", name="All").click(timeout=5000)
            time.sleep(2)
        except Exception:
            pass

        # Export with 180s timeout
        with page.expect_download(timeout=180000) as dl_info:
            page.get_by_role("button", name="Export to Excel").first.click(timeout=180000)

        download = dl_info.value
        ext = os.path.splitext(download.suggested_filename)[1] or ".xlsx"
        save_path = os.path.join(DOWNLOAD_DIR, f"{output_filename}{ext}")
        download.save_as(save_path)
        print(f"   ✅ Saved: {output_filename}{ext} ({os.path.getsize(save_path):,} bytes)")
        return True
    finally:
        page.close()

def main():
    print("Starting automated data download from MyDaily...")
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        if os.path.exists(SESSION_FILE):
            context = browser.new_context(storage_state=SESSION_FILE, accept_downloads=True)
        else:
            context = browser.new_context(accept_downloads=True)

        try:
            print("\n[1/2] Downloading All Reports -> Report_Past Month.xlsx...")
            download_report(context, "All Reports", "Report_Past Month")

            print("\n[2/2] Downloading Task Pauses -> Pause Profile.xlsx...")
            download_report(context, "Task Pauses", "Pause Profile")

            print("\n🎉 Both files downloaded successfully!")
        except Exception as e:
            print(f"❌ An error occurred: {e}")
            if os.path.exists(SESSION_FILE):
                try:
                    os.remove(SESSION_FILE)
                except Exception:
                    pass
        finally:
            context.close()
            browser.close()

if __name__ == "__main__":
    main()
