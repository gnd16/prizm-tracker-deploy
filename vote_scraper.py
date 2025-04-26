from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
import pytz
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import os
import time

# Set timezone ke WIB
wib = pytz.timezone('Asia/Jakarta')

# Setup Google Sheets credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_json = os.environ.get('CREDS_JSON')
creds_dict = json.loads(creds_json)
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# Buka Spreadsheet
spreadsheet = client.open_by_key("1ky4L-2L7E5az3yAqglJvWWMmzTCWWZY2sAPmC_5RCkc")

# Target aktor yang ingin di-track
target_names = ["KIM HYE YOON", "IU", "LEE HYE RI", "PARK BO GUM", "BYEON WOO SEOK"]

# Buka atau buat Sheet
def get_or_create_sheet(sheet_name):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")

sheet_raw = get_or_create_sheet("Raw Votes")
sheet_diff = get_or_create_sheet("Selisih Votes")
sheet_ringkasan = get_or_create_sheet("Ringkasan Harian")

# Header Raw Votes
header_raw = ["Waktu Ambil"] + target_names
if sheet_raw.row_count == 0:
    sheet_raw.append_row(header_raw)

# Header Selisih
header_diff = ["Waktu Ambil"] + target_names
if sheet_diff.row_count == 0:
    sheet_diff.append_row(header_diff)

# Header Ringkasan
header_ringkasan = ["Tanggal"] + target_names
if sheet_ringkasan.row_count == 0:
    sheet_ringkasan.append_row(header_ringkasan)

# Function ambil voting day
def get_voting_day(dt):
    dt = dt.astimezone(wib)
    batas = dt.replace(hour=22, minute=0, second=0, microsecond=0)
    if dt < batas:
        return dt.date() - timedelta(days=1)
    else:
        return dt.date()

# Mulai loop utama
while True:
    try:
        # Setup WebDriver
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        driver = webdriver.Chrome(options=options)

        # Buka PRIZM Voting Page
        driver.get("https://global.prizm.co.kr/story/voteforbaeksang")
        time.sleep(5)

        # Ambil waktu
        waktu_ambil = datetime.now(wib).strftime("%Y-%m-%d %H:%M:%S")

        # Ambil nama aktor & jumlah suara
        names = driver.find_elements(By.CLASS_NAME, "item-title")
        votes = driver.find_elements(By.CLASS_NAME, "item-count")

        data_aktor = {}
        for name, vote in zip(names, votes):
            actor_name = name.text.strip().upper()
            actor_name = actor_name.replace("LEE HYERI", "LEE HYE RI")  # Koreksi nama
            if actor_name in target_names:
                vote_clean = int(vote.text.replace(",", ""))
                data_aktor[actor_name] = vote_clean

        print("Data aktor yang ketemu:", data_aktor)

        # Siapin data untuk update
        data_row = [waktu_ambil] + [data_aktor.get(name, 0) for name in target_names]
        sheet_raw.append_row(data_row)

        # Hitung selisih
        rows_raw = sheet_raw.get_all_values()
        if len(rows_raw) > 2:
            latest = list(map(int, rows_raw[-1][1:]))
            previous = list(map(int, rows_raw[-2][1:]))
            selisih = [latest[i] - previous[i] for i in range(len(latest))]
        else:
            selisih = [0] * len(target_names)

        data_diff = [waktu_ambil] + selisih
        sheet_diff.append_row(data_diff)

        # Update Ringkasan Harian
        hari_ini = get_voting_day(datetime.now(wib)).strftime("%Y-%m-%d")
        rows_diff = sheet_diff.get_all_records()
        summary = {name: 0 for name in target_names}

        for row in rows_diff:
            waktu_row = datetime.strptime(row["Waktu Ambil"], "%Y-%m-%d %H:%M:%S")
            if get_voting_day(waktu_row).strftime("%Y-%m-%d") == hari_ini:
                for name in target_names:
                    summary[name] += int(row[name])

        try:
            existing_rows = sheet_ringkasan.get_all_records()
            tanggal_list = [r["Tanggal"] for r in existing_rows]
            if hari_ini in tanggal_list:
                idx = tanggal_list.index(hari_ini) + 2
                sheet_ringkasan.update(f"A{idx}", [[hari_ini] + [summary[name] for name in target_names]])
            else:
                sheet_ringkasan.append_row([hari_ini] + [summary[name] for name in target_names])
        except Exception as e:
            print(f"Error update ringkasan: {e}")

        driver.quit()

    except Exception as e:
        print(f"Error saat update: {e}")

    print("Tidur 10 menit...")
    time.sleep(600)

