# -*- coding: utf-8 -*-
"""
Created on Wed Dec 31 08:41:09 2025

@author: Sabin
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import csv
import os

class BraveScraper:
    def __init__(self, chromedriver_path='chromedriver.exe'):
        self.driver = self.setup_brave(chromedriver_path)
        self.wait = WebDriverWait(self.driver, 20)
        
    def setup_brave(self, chromedriver_path):
        """Set up Brave browser with Selenium"""
        options = Options()
        brave_paths = [
            r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
            r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
            os.path.expanduser(r"~\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe"),
            r"D:\Brave\Application\brave.exe",
        ]
        
        brave_found = None
        for path in brave_paths:
            if os.path.exists(path):
                brave_found = path
                break
        
        if brave_found:
            print(f"🚀 Using Brave browser: {brave_found}")
            options.binary_location = brave_found
        else:
            print("⚠️ Brave not found. Using Chrome instead.")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        service = Service(chromedriver_path)
        driver = webdriver.Chrome(service=service, options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        return driver
    
    def wait_for_overlay_to_disappear(self):
        self.wait.until(lambda d: d.execute_script(
            "const ov = document.getElementById('overlay');"
            "return !ov || ov.style.display === 'none' || ov.style.visibility === 'hidden';"
        ))
    
    def scrape_tenders(self, max_pages=60, max_rows=1200):
        """Main scraping function"""
        
        print("🌐 Opening Bolpatra website...")
        self.driver.get('https://bolpatra.gov.np/egp/searchOpportunity')
        
        all_rows_data = []
        seen_keys = set()
        
        try:
            search_btn = self.wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input.button[onclick*='searchBidsByCriteria']")
            ))
            search_btn.click()
            print("✅ Clicked the Search button.")

            self.wait.until(EC.presence_of_element_located((By.ID, "dashBoardBidResult")))
            self.wait_for_overlay_to_disappear()

            for page_num in range(1, max_pages + 1):
                rows = self.wait.until(EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "#dashBoardBidResult tbody tr")
                ))
                print(f"📊 Found {len(rows)} rows on page {page_num}")

                for row in rows:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    data = [col.text.strip() for col in cols]
                    
                    if len(data) < 3:
                        continue
                    
                    unique_key = data[0] + "|" + data[1] + "|" + data[2]
                    if unique_key not in seen_keys:
                        seen_keys.add(unique_key)
                        all_rows_data.append(data)
                    else:
                        print(f"⚠️ Duplicate row skipped: Sl. No. {data[0]}")

                    if len(all_rows_data) >= max_rows:
                        print(f"✅ Reached max rows: {max_rows}")
                        break

                if len(all_rows_data) >= max_rows:
                    break

                if page_num < max_pages:
                    try:
                        self.wait_for_overlay_to_disappear()
                        goto_input = self.wait.until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "input.gotoPage"))
                        )
                        self.driver.execute_script("arguments[0].value = '';", goto_input)
                        goto_input.send_keys(str(page_num + 1))
                        
                        goto_btn = self.wait.until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "img.goto"))
                        )
                        goto_btn.click()
                        print(f"➡️ Going to page {page_num + 1}")
                        
                        self.wait.until(EC.staleness_of(rows[0]))
                        self.wait_for_overlay_to_disappear()
                        time.sleep(1)
                        
                    except Exception as e:
                        print(f"❌ Could not go to page {page_num + 1}: {e}")
                        break
                else:
                    print("✅ Reached final page.")
        
        except Exception as e:
            print(f"❌ Error during scraping: {e}")
        
        return all_rows_data
    
    def save_to_csv(self, data, filename='bolpatra_tenders_brave.csv'):
        """Save data to CSV file"""
        if not data:
            print("❌ No data to save")
            return
        
        with open(filename, 'w', newline='', encoding='utf-8') as f:
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
            writer.writerows(data)
        
        print(f"💾 Saved {len(data)} tenders to {filename}")
        return filename
    
    def close(self):
        """Close the browser"""
        if self.driver:
            self.driver.quit()
            print("👋 Browser closed")

# Usage
if __name__ == "__main__":
    scraper = None
    try:
        scraper = BraveScraper(chromedriver_path='chromedriver.exe')
        tenders_data = scraper.scrape_tenders(max_pages=60, max_rows=1500)
        if tenders_data:
            scraper.save_to_csv(tenders_data)
        
    except Exception as e:
        print(f"❌ Main error: {e}")
    
    finally:
        if scraper:
            scraper.close()
