from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import csv

service = Service('chromedriver.exe')
driver = webdriver.Chrome(service=service)
wait = WebDriverWait(driver, 20)

driver.get('https://bolpatra.gov.np/egp/searchOpportunity')

def wait_for_overlay_to_disappear():
    wait.until(lambda d: d.execute_script(
        "const ov = document.getElementById('overlay');"
        "return !ov || ov.style.display === 'none' || ov.style.visibility === 'hidden';"
    ))

try:
    search_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.button[onclick*='searchBidsByCriteria']")))
    search_btn.click()
    print("✅ Clicked the Search button.")

    wait.until(EC.presence_of_element_located((By.ID, "dashBoardBidResult")))
    wait_for_overlay_to_disappear()

    all_rows_data = []
    seen_keys = set()
    max_pages = 30      
    max_rows = 1000       

    for page_num in range(1, max_pages + 1):
        rows = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#dashBoardBidResult tbody tr")))
        print(f"✅ Found {len(rows)} rows on page {page_num}")

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            data = [col.text.strip() for col in cols]
            unique_key = data[0] + "|" + data[1] + "|" + data[2]
            if unique_key not in seen_keys:
                seen_keys.add(unique_key)
                all_rows_data.append(data)
            else:
                print(f"⚠️ Duplicate row skipped: Sl. No. {data[0]}, IFB No. {data[1]}")

            if len(all_rows_data) >= max_rows:
                print(f"✅ Reached max rows: {max_rows}")
                break

        if len(all_rows_data) >= max_rows:
            break

        if page_num < max_pages:
            try:
                wait_for_overlay_to_disappear()
                goto_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.gotoPage")))
                driver.execute_script("arguments[0].value = '';", goto_input)
                goto_input.send_keys(str(page_num + 1))
                goto_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img.goto")))
                goto_btn.click()
                print(f"✅ Triggered goto page {page_num + 1}")
                wait.until(EC.staleness_of(rows[0]))
                wait_for_overlay_to_disappear()
                time.sleep(1)
            except Exception as e:
                print(f"❌ Could not go to next page after page {page_num}: {e}")
                break
        else:
            print("✅ Reached max pages or last page.")

finally:
    csv_file = 'bolpatra_tenders.csv'
    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow([
            "Sl. No.",
            "IFB / RFP / EOI / PQ No.",
            "Project Title",
            "Public Entity Name",
            "Procurement Type",
            "Status",
            "Notice Published Date",
            "Last Date of Bid Submission",
            "No of Days Left"
        ])
        writer.writerows(all_rows_data)

    print(f"📄 Saved data to {csv_file}")
    driver.quit()
