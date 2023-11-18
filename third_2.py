import gspread
import os
import sqlite3
import csv
import pandas as pd
# import json
import random
from googletrans import Translator
import time
import datetime
from dotenv import load_dotenv

load_dotenv()

# ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials

# 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# 認証情報設定
# ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
import os

cre = os.getenv("CRE")
credentials = ServiceAccountCredentials.from_json_keyfile_name(cre, scope)

# OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

translator = Translator()


def get_unique_list(seq):
    seen = []
    return [x for x in seq if x not in seen and not seen.append(x)]


wbname = "st_20220916"
sheetname = "total"
PRICE_DIR = "fobprice"
fob_file_name= "fob_09_16.txt"
wb = gc.open(wbname)
wss = wb.worksheets()
try:
    ws = wb.add_worksheet(title=sheetname, rows="100", cols="100")
except Exception:
    print("シートは既に存在します")
    pass
# print(wss)
ws1 = wb.worksheet(sheetname)

sanichi = "SANICHI CORPORATION\n \
NO.602 EIDAI BLDG., 34-21 1-CHOME\n \
NISHI-ku, OSAKA 550 0013, JAPAN\n \
TEL 81-6-6543-6737 FAX: 81-6-6543-6747\n \
e-mail:    sanichi@d4.dion.ne.jp"
ws1.update_cell(1, 1, sanichi)
dt_now = datetime.datetime.now()
ws1.update_cell(1, 9, dt_now.strftime("%Y/%m/%d"))



# headerの書き込み
header = ["id", "maker", "model", "serial_number", "height", "attachment", "year", "hour_meter", "applicable", "fob_price"]
for i, h in enumerate((header)):
    ws1.update_cell(2, i + 1, h)

amount_list = []
with open(f"{PRICE_DIR}/{fob_file_name}") as f:
    amount_list = f.readlines()
amount_list = list(map(lambda x: x.rstrip(), amount_list))

fork_amount_index = [i for i, x in enumerate(amount_list) if not x]
diesel_amount_list = amount_list[:fork_amount_index[0]]
gasoline_amount_list = amount_list[fork_amount_index[0] + 1:fork_amount_index[1]]
battery_amount_list = amount_list[fork_amount_index[1] + 1:fork_amount_index[2]]
shovelloader_amount_list = amount_list[fork_amount_index[2] + 1:]
"""
for i in range(2):
    a_list.append([translator.translate(s).text for  s in wss[i].col_values(3)])
    a_list.append(wss[i].col_values(4))
    a_list.append(wss[i].col_values(5))
    a_list.append([translator.translate(s).text for  s in wss[i].col_values(8)])
    a_list.append([translator.translate(s).text for  s in wss[i].col_values(12)])
    a_list.append([translator.translate(s).text for  s in wss[i].col_values(13)])
    a_list.append([translator.translate(s).text for  s in wss[i].col_values(14)])
    a_list.append([translator.translate(s).text for  s in wss[i].col_values(15)])
    d = 0
    for i in zip(a_list[d],a_list[d+1],a_list[d+2],a_list[d+3],a_list[d+4],a_list[d+5],a_list[d+6],a_list[d+7]):
        b_list.append(row_data[0],list(i))
        a_list = []
"""

# diesellist
diesel_temp = []
diesel_temp_list = wss[0].get_all_values()
diesel_count = len(diesel_temp_list)
diesel_fork_list = []
diesel_fork_list.append(["diesel"])
for i, row_data in enumerate(diesel_temp_list):
    if i < 2:
        continue
    for data in row_data:
        diesel_temp.append(translator.translate(data).text)
    diesel_fork_list.append(
        [
            row_data[0],
            diesel_temp[2],
            row_data[3],
            row_data[4],
            diesel_temp[7],
            diesel_temp[11],
            diesel_temp[12],
            diesel_temp[13],
            diesel_temp[14],
            diesel_amount_list[i - 2],
        ]
    )
    diesel_temp = []

tuple_diesel_list = []
for i, diesel_row in enumerate(diesel_fork_list):
    if i == 0:
        continue
    tuple_diesel_list.append(tuple(diesel_row))
# gasolinelist
gasoline_temp = []
gasoline_temp_list = wss[1].get_all_values()
gasoline_count = len(gasoline_temp_list)
gasoline_fork_list = []
gasoline_fork_list.append(["gasoline"])
for i, row_data in enumerate(gasoline_temp_list):
    if i < 2:
        continue
    for data in row_data:
        gasoline_temp.append(translator.translate(data).text)
    gasoline_fork_list.append(
        [
            row_data[0],
            gasoline_temp[2],
            row_data[3],
            row_data[4],
            gasoline_temp[7],
            gasoline_temp[11],
            gasoline_temp[12],
            gasoline_temp[13],
            gasoline_temp[14],
            gasoline_amount_list[i - 2],
        ]
    )
    gasoline_temp = []
tuple_gasoline_list = []
for i, gasoline_row in enumerate(gasoline_fork_list):
    if i == 0:
        continue
    tuple_gasoline_list.append(tuple(gasoline_row))
time.sleep(10)
# batterylist
battery_temp = []
battery_temp_list = wss[2].get_all_values()
battery_count = len(battery_temp_list)
battery_fork_list = []
battery_fork_list.append(["battery"])
for i, row_data in enumerate(battery_temp_list):
    if i < 2:
        continue
    for data in row_data:
        battery_temp.append(translator.translate(data).text)
    battery_fork_list.append(
        [
            row_data[0],
            battery_temp[2],
            row_data[3],
            row_data[4],
            battery_temp[7],
            battery_temp[11],
            battery_temp[12],
            battery_temp[13],
            battery_temp[14],
            battery_amount_list[i - 2],
        ]
    )
    battery_temp = []
tuple_battery_list = []
for i, battery_row in enumerate(battery_fork_list):
    if i == 0:
        continue
    tuple_battery_list.append(tuple(battery_row))
time.sleep(10)
#shovelloaderlist
shovelloader_temp = []
shovelloader_temp_list = wss[3].get_all_values()
shovelloader_count = len(shovelloader_temp_list)
shovelloader_fork_list = []
shovelloader_fork_list.append(["shovelloader"])
for i, row_data in enumerate(shovelloader_temp_list):
    if i < 2:
        continue
    for data in row_data:
        shovelloader_temp.append(translator.translate(data).text)
    shovelloader_fork_list.append(
        [
            row_data[0],
            shovelloader_temp[2],
            row_data[3],
            row_data[4],
            shovelloader_temp[7],
            shovelloader_temp[11],
            shovelloader_temp[12],
            shovelloader_temp[13],
            shovelloader_temp[14],
            shovelloader_amount_list[i - 2],
        ]
    )
    shovelloader_temp = []
tuple_shovelloader_list = []
for i, shovelloader_row in enumerate(shovelloader_fork_list):
    if i == 0:
        continue
    tuple_shovelloader_list.append(tuple(diesel_row))
time.sleep(10)

#csv書き出し
CSV_DIR = f"csv/{wbname}"
if not os.path.isdir(CSV_DIR):
    os.makedirs(CSV_DIR)

with open(f'{CSV_DIR}/diesel_forklift.csv', 'w', newline='') as diesel_file:
    writer = csv.writer(diesel_file)
    for i , diesel_row in enumerate((diesel_fork_list)):
        if i == 0:
            continue
        writer.writerow(diesel_row)

with open(f'{CSV_DIR}/gasoline_forklift.csv', 'w', newline='') as gasoline_file:
    writer = csv.writer(gasoline_file)
    for i , gasoline_row in enumerate((gasoline_fork_list)):
        if i == 0:
            continue
        writer.writerow(gasoline_row)

with open(f'{CSV_DIR}/battery_forklift.csv', 'w', newline='') as battery_file:
    writer = csv.writer(battery_file)
    for i , battery_row in enumerate((battery_fork_list)):
        if i == 0:
            continue
        writer.writerow(battery_row)

with open(f'{CSV_DIR}/shovelloader_forklift.csv', 'w', newline='') as shovelloader_file:
    writer = csv.writer(shovelloader_file)
    for i , shovelloader_row in enumerate(shovelloader_fork_list):
        if i == 0:
            continue
        writer.writerow(shovelloader_row)


# sqliteデータベース用
SQLITE_DIR = "sqlite"
if not os.path.isdir(SQLITE_DIR):
    os.makedirs(SQLITE_DIR)

save_stocklist = SQLITE_DIR + "/" + wbname + ".db"

# sqlite接続
conn = sqlite3.connect(save_stocklist)
# カーソルを取得
c = conn.cursor()

# dieseltable
c.execute("CREATE TABLE IF NOT EXISTS diesel_stocklist(id int, maker text, model text, serial_number text, height text, attachment text, year text, hour_meter text, applicable text, amount text)")
for diesel_row in tuple_diesel_list:
    c.execute("insert into diesel_stocklist(id, maker, model, serial_number, height, attachment, year, hour_meter, applicable, amount)" f"values {diesel_row}")
#for row in tuple_diesel_list:
#    c.execute("select id, maker, model, serial_number, height, attachment, year, hour_meter, applicable, amount from stocklist")
#    print(row)
conn.commit()

#gasolinetable
c.execute("CREATE TABLE IF NOT EXISTS gasoline_stocklist(id int, maker text, model text, serial_number text, height text, attachment text, year text, hour_meter text, applicable text, amount text)")
for gasoline_row in tuple_gasoline_list:
    c.execute("insert into gasoline_stocklist(id, maker, model, serial_number, height, attachment, year, hour_meter, applicable, amount)" f"values {gasoline_row}")
#for row in tuple_gasoline_list:
#    c.execute("select id, maker, model, serial_number, height, attachment, year, hour_meter, applicable, amount from stocklist")
#    print(row)
conn.commit()

#batterytable
c.execute("CREATE TABLE IF NOT EXISTS battery_stocklist(id int, maker text, model text, serial_number text, height text, attachment text, year text, hour_meter text, applicable text, amount text)")
for battery_row in tuple_battery_list:
    c.execute("insert into battery_stocklist(id, maker, model, serial_number, height, attachment, year, hour_meter, applicable, amount)" f"values {battery_row}")
#for row in tuple_battery_list:
#    c.execute("select id, maker, model, serial_number, height, attachment, year, hour_meter, applicable, amount from stocklist")
#    print(row)
conn.commit()

#shovelloadertable
c.execute("CREATE TABLE IF NOT EXISTS shovelloader_stocklist(id int, maker text, model text, serial_number text, height text, attachment text, year text, hour_meter text, applicable text, amount text)")
for shovelloader_row in tuple_shovelloader_list:
    c.execute("insert into shovelloader_stocklist(id, maker, model, serial_number, height, attachment, year, hour_meter, applicable, amount)" f"values {shovelloader_row}")
#for row in tuple_shovelloader_list:
#    c.execute("select id, maker, model, serial_number, height, attachment, year, hour_meter, applicable, amount from stocklist")
#    print(row)
conn.commit()
conn.close()

"""
temp_list = []
shobel_of_list = wss[1].get_all_values()
for i,row_data in enumerate(shobel_of_list):
    for data in row_data:
        temp_list.append(translator.translate(data).text)
    b_list.append(row_data[0],[temp_list[2],row_data[3],row_data[4], temp_list[7], temp_list[11], temp_list[12], temp_list[13], temp_list[14]])
    temp_list = []


temp_list = []
gasoline_list = wss[2].get_all_values()
for i,row_data in enumerate(gasoline_list):
    for data in row_data:
        temp_list.append(translator.translate(data).text)
    b_list.append(row_data[0],[temp_list[2],row_data[3],row_data[4], temp_list[7], temp_list[11], temp_list[12], temp_list[13], temp_list[14]])
    temp_list = []


temp_list = []
shobel_of_list = wss[3].get_all_values()
for i,row_data in enumerate(shobel_of_list):
    for data in row_data:
        temp_list.append(translator.translate(data).text)
    b_list.append(row_data[0],[temp_list[2],row_data[3],row_data[4], temp_list[7], temp_list[11], temp_list[12], temp_list[13], temp_list[14]])
    temp_list = []



#print(b_list)

"""

"""
import random
a = []
for i in range(100):
    a.append(str(random.randint(900000,1500000)) + '\n')
with open('third.txt', 'w') as f:
    f.writelines(a)

from pathlib import Path
"""


"""
tranlated_list = sorted(translated_list, key=lambda x: x[2])
print(translated_list)

"""


"""
c_list = []
for i, a in enumerate(b_list):
    d = [b.replace('\n','') for b in a]
    c_list.append(d)
c_list = get_unique_list(c_list)
"""

"""
for i,b in enumerate(c_list):
    if i > 1:
        b.append(l[i-2])
    else:
        pass
"""
"""
c_list[0][0] = "maker"
c_list[0][1] =  "model"
c_list[0][2] = "serial number"
c_list[0][3] = "height"
c_list[0][4] = "attachment"
c_list[0][5] = "year"
c_list[0][6] = "hour meter"
c_list[0][7] = "applicable"
c_list[0].append("amount")


#print(c_list)

"""

"""
ws1.update("A1:D8", "SANICHI CORPORATION\nNO.602 EIDAI BLDG., 34-21 1-CHOME\nNISHI-ku, OSAKA 550 0013, JAPAN\nTEL 81-6-6543-6737 FAX: 81-6-6543-6747\ne-mail:    sanichi@d4.dion.ne.jp")
"""


# print(len(translated_list))
"""
for i in range(len(c_list)):
    ws1.append_row(c_list[i])
    if i % 50 == 0:
        time.sleep(30)
"""

for i in range(len(diesel_fork_list)):
    ws1.append_row(diesel_fork_list[i])
    if i % 50 == 0:
        time.sleep(10)

for i in range(len(gasoline_fork_list)):
    ws1.append_row(gasoline_fork_list[i])
    if i % 50 == 0:
        time.sleep(10)

for i in range(len(battery_fork_list)):
    ws1.append_row(battery_fork_list[i])
    if i % 50 == 0:
        time.sleep(10)

for i in range(len(shovelloader_fork_list)):
    ws1.append_row(shovelloader_fork_list[i])
    if i % 50 == 0:
        time.sleep(10)

######
######
######
