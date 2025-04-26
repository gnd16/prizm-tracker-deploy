from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# ==== Konfigurasi ====
target_names = ["KIM HYE YOON", "IU", "LEE HYE RI", "PARK BO GUM", "BYEON WOO SEOK"]
spreadsheet_id = "1ky4L-2L7E5az3yAqglJvWWMmzTCWWZY2sAPmC_5RCkc"

def get_voting_day(now):
    cutoff_hour = 22
    if now.hour < cutoff_hour:
        return now.date()
    else:
        return now.date() + timedelta(days=1)

# ==== Ambil Data PRIZM ====
service = Service("/usr/local/bin/chromedriver")
options = webdriver.ChromeOptions()
options.headless = False  # ubah ke True kalau gak mau liat browser
driver = webdriver.Chrome(service=service, options=options)

driver.get("https://global.prizm.co.kr/story/voteforbaeksang")
time.sleep(5)
# Scroll ke bawah supaya semua aktor tampil
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(3)

names = driver.find_elements(By.CLASS_NAME, "item-title")
votes = driver.find_elements(By.CLASS_NAME, "item-count")
waktu_ambil = datetime.now()
waktu_str = waktu_ambil.strftime("%Y-%m-%d %H:%M:%S")

data_dict = {}
for name, vote in zip(names, votes):
    nama = name.text.strip()
    if nama in target_names:
        jumlah = int(vote.text.replace(",", ""))
        data_dict[nama] = jumlah
print("Data aktor yang ketemu:")
print(data_dict)
driver.quit()

# ==== Koneksi ke Google Sheets ====
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("/Users/gitanofiekadwijayati/VotingPRIZM/creds.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(spreadsheet_id)

def get_or_create_sheet(title):
    try:
        return spreadsheet.worksheet(title)
    except:
        return spreadsheet.add_worksheet(title=title, rows="1000", cols="20")

# === Sheet 1: SUARA ===
sheet_suara = get_or_create_sheet("Suara")
header = ["Waktu Ambil"] + target_names
if not sheet_suara.get_all_values():
    sheet_suara.append_row(header)
vote_row = [waktu_str] + [data_dict.get(name, "") for name in target_names]
sheet_suara.append_row(vote_row)

# === Sheet 2: SELISIH ===
sheet_diff = get_or_create_sheet("Selisih")
if not sheet_diff.get_all_values():
    sheet_diff.append_row(header)

rows_suara = sheet_suara.get_all_values()
if len(rows_suara) > 1:
    prev_row = rows_suara[-2]
    prev_votes = {name: int(val.replace(",", "")) if val else 0 for name, val in zip(header[1:], prev_row[1:])}
else:
    prev_votes = {name: 0 for name in target_names}

diff_row = [waktu_str]
for name in target_names:
    now = data_dict.get(name, 0)
    before = prev_votes.get(name, 0)
    selisih = now - before
    diff_row.append(f"{selisih:+,}")
sheet_diff.append_row(diff_row)

# 🎨 Auto-Coloring Selisih
fmt_red = cellFormat(backgroundColor=color(1, 0.85, 0.85))
fmt_green = cellFormat(backgroundColor=color(0.85, 1, 0.85))
row_index = len(sheet_diff.get_all_values())
for col in range(2, len(header) + 1):
    value = sheet_diff.cell(row_index, col).value
    if value.startswith("+"):
        format_cell_range(sheet_diff, f"{chr(64+col)}{row_index}", fmt_green)
    elif value.startswith("-"):
        format_cell_range(sheet_diff, f"{chr(64+col)}{row_index}", fmt_red)

# === Sheet 3: RINGKASAN HARIAN ===
sheet_ringkasan = get_or_create_sheet("Ringkasan Harian")
header_ringkasan = ["Tanggal"] + target_names
rows_diff = sheet_diff.get_all_records()
hari_ini = get_voting_day(waktu_ambil).strftime("%Y-%m-%d")
summary = {name: 0 for name in target_names}

for row in rows_diff:
    waktu_row = datetime.strptime(row["Waktu Ambil"], "%Y-%m-%d %H:%M:%S")
    if get_voting_day(waktu_row).strftime("%Y-%m-%d") == hari_ini:
        for name in target_names:
            selisih = row.get(name)
            if selisih:
                summary[name] += int(selisih.replace(",", "").replace("+", ""))

rows_summary = sheet_ringkasan.get_all_values()
dates = [r[0] for r in rows_summary[1:]] if len(rows_summary) > 1 else []

if "Tanggal" not in dates:
    if not rows_summary:
        sheet_ringkasan.append_row(header_ringkasan)

if hari_ini in dates:
    idx = dates.index(hari_ini) + 2
    sheet_ringkasan.delete_rows(idx)
sheet_ringkasan.append_row([hari_ini] + [summary[name] for name in target_names])

print("✅ Sukses update: Suara, Selisih, dan Ringkasan Harian!")

