import re
import os

from playwright.sync_api import sync_playwright
import time
import random
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

# ---------------- READ EXCEL ----------------
def read_asins_from_excel(file_path):
    df = pd.read_excel(file_path, header=None)  # no header
    return df[0].dropna().astype(str).tolist()


# ---------------- SAVE TO EXCEL ----------------
def save_to_excel(data, file_name="amazon_output.xlsx"):
    df = pd.DataFrame(data)
    df.to_excel(file_name, index=False)

def scrape_single(asin):
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
            )
            page = context.new_page()

            url = f"https://www.amazon.in/dp/{asin}"
            page.goto(url, timeout=60000)

            time.sleep(random.uniform(3, 6))

            data = {"asin": asin}

            # Title
            try:
                data["title"] = page.locator("#productTitle").first.inner_text().strip()
            except:
                data["title"] = None

            # Price
            try:
                data["price"] = page.locator(".a-price-whole").first.inner_text()
            except:
                data["price"] = None

            # Rating
            try:
                data["rating"] = page.locator("span.a-icon-alt").first.inner_text()
            except:
                data["rating"] = None

            try:
                elements = page.locator("#acrCustomerReviewText").all()

                if elements:
                    text = elements[0].inner_text()
                    number = re.findall(r'\d[\d,]*', text)[0]
                    data["reviews_count"] = int(number.replace(",", ""))
                else:
                    data["reviews_count"] = 0

            except:
                data["reviews_count"] = 0

            browser.close()
            return data

    except Exception as e:
        return {"asin": asin, "error": str(e)}


# ---------------- YOUR SCRAPER FUNCTION ----------------
def scrape_parallel(asins, workers=5):
    results = []

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [executor.submit(scrape_single, asin) for asin in asins]

        for future in as_completed(futures):
            result = future.result()
            print(f"✅ Done: {result.get('asin')}")
            results.append(result)

    return results

if __name__ == "__main__":
    os.system("playwright install chromium")
    input_file = "asin_list.xlsx"   # your input file
    output_file = "amazon_data.xlsx"

    # Step 1: Read ASINs
    asin_list = read_asins_from_excel(input_file)

    print(f"Total ASINs: {len(asin_list)}")

    # Step 2: Scrape
    data = scrape_parallel(asin_list, workers=3)

    # Step 3: Save to Excel
    save_to_excel(data, output_file)

    print(f"🔥 Data saved to {output_file}")