# vote_scraper.py

import os
import time
from datetime import datetime, timedelta
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# ==== Setup Google Sheets ====
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_json = os.environ.get('CREDS_JSON')  # Ambil dari environment di Render
if creds_json:
    import json
    creds = json.loads(creds_json)
    client = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_dict(creds, scope))
else:
    creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
    client = gspread.authorize(creds)

SHEET_URL = "https://docs.google.com/spreadsheets/d/1ky4L-2L7E5az3yAqglJvWWMmzTCWWZY2sAPmC_5RCkc/edit"
spreadsheet = client.open_by_url(SHEET_URL)

sheet_suara = spreadsheet.worksheet("Suara")
sheet_selisih = spreadsheet.worksheet("Selisih")
sheet_ringkasan = spreadsheet.worksheet("Ringkasan Harian")

target_names = ["KIM HYE YOON", "IU", "LEE HYE RI", "PARK BO GUM", "BYEON WOO SEOK"]

# ==== Function to get voting day ====
def get_voting_day(timestamp):
    now_wib = timestamp + timedelta(hours=7)
    cutoff = now_wib.replace(hour=22, minute=0, second=0, microsecond=0)
    if now_wib < cutoff:
        return now_wib.date()
    else:
        return (now_wib + timedelta(days=1)).date()

# ==== Scrape Data ====
def scrape_votes():
    try:
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        driver = webdriver.Chrome(options=options)

        driver.get("https://global.prizm.co.kr/story/voteforbaeksang")
        time.sleep(5)

        names = driver.find_elements(By.CLASS_NAME, "item-title")
        votes = driver.find_elements(By.CLASS_NAME, "item-count")

        result = {}
        for name_element, vote_element in zip(names, votes):
            name = name_element.text.strip().upper()
            vote = vote_element.text.strip().replace(",", "")
            if name in target_names:
                result[name] = int(vote)

        driver.quit()
        return result
    except Exception as e:
        print(f"Error saat scraping: {e}")
        return None

# ==== Update Sheets ====
def update_sheets(vote_data):
    if not vote_data:
        print("No data fetched.")
        return

    now = datetime.utcnow() + timedelta(hours=7)  # WIB
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")

    # === Update Suara ===
    row = [now_str] + [vote_data.get(name, "") for name in target_names]
    sheet_suara.append_row(row, value_input_option="USER_ENTERED")

    # === Update Selisih ===
    try:
        suara_data = sheet_suara.get_all_values()

        if len(suara_data) >= 3:
            prev_row = suara_data[-2]
            curr_row = suara_data[-1]

            prev_votes = [int(x.replace(",", "")) if x else 0 for x in prev_row[1:]]
            curr_votes = [int(x.replace(",", "")) if x else 0 for x in curr_row[1:]]

            selisih_votes = [curr - prev for curr, prev in zip(curr_votes, prev_votes)]

            diff_row = [now_str] + selisih_votes
            sheet_selisih.append_row(diff_row, value_input_option="USER_ENTERED")
        else:
            print("Belum cukup data untuk hitung selisih.")
    except Exception as e:
        print(f"Error saat update selisih: {e}")

    # === Update Ringkasan Harian ===
    try:
        rows_diff = sheet_selisih.get_all_records()
        hari_ini = get_voting_day(now).strftime("%Y-%m-%d")
        summary = {name: 0 for name in target_names}

        for row in rows_diff:
            waktu_row = datetime.strptime(row["Waktu Ambil"], "%Y-%m-%d %H:%M:%S")
            if get_voting_day(waktu_row).strftime("%Y-%m-%d") == hari_ini:
                for name in target_names:
                    summary[name] += int(row.get(name, 0))

        existing_summary = sheet_ringkasan.get_all_values()
        if existing_summary and existing_summary[-1][0] == hari_ini:
            sheet_ringkasan.delete_rows(len(existing_summary))
        
        final_row = [hari_ini] + [summary[name] for name in target_names]
        sheet_ringkasan.append_row(final_row, value_input_option="USER_ENTERED")
    except Exception as e:
        print(f"Error saat update ringkasan: {e}")

# ==== Main Loop ====
while True:
    votes = scrape_votes()
    update_sheets(votes)
    time.sleep(120)  # 120 seconds = 2 minutes
